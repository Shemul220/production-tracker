import base64
import hmac
import io
import json
import re
from urllib.parse import quote
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="BSU Production Tracker", page_icon="🏭", layout="wide")

# রিপোজিটরিতে এই ফোল্ডারে বা data.xlsx নামে ফাইল রাখলে সাইট নিজে থেকেই পড়ে নেবে।
DATA_DIR = Path("data")
DEFAULT_FILE = Path("data.xlsx")

PROD_COLS = ["Stage", "Date", "Day", "Shift", "Shift Owner", "Line", "Code", "Batch No",
             "Product Name", "Batch Size", "Pack Size", "Run Start", "Run End", "Manpower",
             "Output", "Hours", "Std Capacity", "Actual Capacity", "Down Time", "Remarks"]
TARGET_COLS = ["Date", "Daily Target", "Daily Output"]
BATCH_COLS = ["Date", "Product Name", "Batch No", "Type"]

# মাসিক Target (মাস → সংখ্যা)। এখানে দেওয়া মান ডিফল্ট; সাইট থেকে বদলালে data/monthly_targets.json-এ সেভ হয়।
DEFAULT_MONTHLY_TARGETS = {"2026-11": 17830000}  # নভেম্বর'26 = 1,78,30,000
MT_FILE = DATA_DIR / "monthly_targets.json"
# ভুল মান → সঠিক মান (মাস: (ভুল, ঠিক))। নভেম্বর'26: 17,83,00,000 নয়, হবে 1,78,30,000
TARGET_FIXES = {"2026-11": (178300000, 17830000)}

# লগবুক ফাইলের কলামের অবস্থান (০ থেকে গোনা)। নতুন কলাম লাগলে এখানে যোগ করুন।
SECONDARY_MAP = {0: "Date", 1: "Day", 2: "Shift", 3: "Shift Owner", 4: "Line", 5: "Code",
                 6: "Batch No", 7: "Product Name", 8: "Batch Size", 9: "Pack Size",
                 13: "Run Start", 14: "Run End", 17: "Manpower", 18: "Output", 25: "Hours",
                 27: "Std Capacity", 28: "Actual Capacity", 30: "Down Time", 31: "Remarks"}
PRIMARY_MAP = {0: "Date", 1: "Day", 2: "Shift", 3: "Shift Owner", 4: "Line", 5: "Code",
               6: "Batch No", 7: "Product Name", 8: "Batch Size", 9: "Pack Size",
               13: "Run Start", 14: "Run End", 17: "Output", 24: "Hours",
               26: "Std Capacity", 27: "Actual Capacity", 29: "Down Time", 30: "Remarks"}


# ---------- ছোট সহায়ক ফাংশন ----------
def to_date(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 40000 < v < 60000:
        return datetime(1899, 12, 30) + timedelta(days=float(v))
    return None


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def txt(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def hhmm(v):
    if hasattr(v, "strftime"):
        return v.strftime("%H:%M")
    return txt(v)


def scan_rows(ws, is_data, max_blank=150):
    """ডেটার সারিগুলো বের করে। ডেটা শেষ হয়ে অনেক ফাঁকা সারি এলে থেমে যায় (বড় ফাঁকা শিটে দ্রুত চলে)।"""
    rows, blank = [], 0
    for r in ws.iter_rows(values_only=True):
        if is_data(r):
            rows.append(r)
            blank = 0
        elif rows:
            blank += 1
            if blank > max_blank:
                break
    return rows


# ---------- ফাইল পড়া ----------
def parse_production(wb):
    ws = wb.worksheets[0]
    head = list(ws.iter_rows(min_row=1, max_row=6, values_only=True))
    flat = " ".join(str(c) for row in head for c in row if c is not None)
    stage = "Secondary" if "Manpower" in flat else "Primary"
    mp = SECONDARY_MAP if stage == "Secondary" else PRIMARY_MAP
    recs = []
    for r in scan_rows(ws, lambda r: len(r) > 0 and to_date(r[0]) is not None):
        r = tuple(r) + (None,) * (32 - len(r))
        rec = {name: r[i] for i, name in mp.items()}
        prod = txt(rec["Product Name"])
        if prod is None or prod.startswith("#"):
            continue
        rec["Stage"] = stage
        rec["Date"] = to_date(rec["Date"])
        rec["Product Name"] = prod
        line = txt(rec["Line"])
        rec["Line"] = re.sub(r"\s*-\s*", "-", line) if line else line
        for c in ("Shift", "Shift Owner", "Code", "Batch No", "Day", "Pack Size", "Remarks"):
            rec[c] = txt(rec[c])
        for c in ("Run Start", "Run End", "Down Time"):
            rec[c] = hhmm(rec[c])
        for c in ("Batch Size", "Manpower", "Output", "Hours", "Std Capacity", "Actual Capacity"):
            rec[c] = num(rec.get(c))
        recs.append(rec)
    return recs


def parse_targets(wb):
    recs = []
    for ws in wb.worksheets:
        first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), (None,))[0]
        if not (isinstance(first, str) and first.strip().lower().startswith("daily target")):
            continue
        for r in ws.iter_rows(values_only=True):
            r = tuple(r) + (None,) * 5
            d = to_date(r[0])
            if d is None:
                continue
            tgt = num(r[2])
            recs.append({"Date": d, "Daily Target": tgt if tgt else None, "Daily Output": num(r[3])})
    return recs


def clean_type(v):
    t = re.sub(r"\s+", "", str(v)).upper() if v is not None else ""
    if not t:
        return "OTHER"
    if "EXP" in t:
        return "EXPORT"
    if "PHY" in t:
        return "PHY"
    if t in ("COM+GS", "COM+G/S"):
        return "COM+G/S"
    if t in ("COM", "G/S", "P/S", "REPACK"):
        return t
    return "OTHER"


def find_batch_header(ws):
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True)):
        names = [str(c).strip().upper() if c is not None else "" for c in r]
        if "PRODUCT NAME" in names and "BATCH NO" in names:
            return i, names
    return None, None


def parse_batches(wb, hdr, names):
    ws = wb.worksheets[0]
    i_date = names.index("DATE") if "DATE" in names else 0
    i_prod, i_batch = names.index("PRODUCT NAME"), names.index("BATCH NO")
    i_type = names.index("TYPE") if "TYPE" in names else None
    recs, cur = [], None
    for k, r in enumerate(ws.iter_rows(values_only=True)):
        if k <= hdr:
            continue
        r = tuple(r) + (None,) * 12
        d = to_date(r[i_date])
        if d:
            cur = d  # তারিখ শুধু দিনের প্রথম সারিতে থাকে
        prod, batch = txt(r[i_prod]), txt(r[i_batch])
        if cur is None or prod is None or batch is None:
            continue
        recs.append({"Date": cur, "Product Name": prod, "Batch No": batch,
                     "Type": clean_type(r[i_type]) if i_type is not None else "OTHER"})
    return recs


EXPORT_COLS = ["Source", "Sheet", "Country", "SAP Code", "Order No", "Machine", "Product Name", "Pack Size",
               "Pcs per Box", "Required Box", "Required Pcs", "Batch Size", "Batches Req", "Batch Number",
               "Status", "Order Date", "Completed Date", "Delivery Date", "PLSD", "RM Available",
               "PM Available", "Remarks", "SO Aging"]

COUNTRY_FIX = {"keneya": "Kenya", "sri-lanka": "Sri Lanka", "sri lanka": "Sri Lanka", "srilanka": "Sri Lanka"}


def clean_country(v):
    s = txt(v)
    return COUNTRY_FIX.get(s.lower(), s) if s else s


def to_ddmmyy(v):
    """'25.03.26' বা আসল তারিখ → datetime (না পারলে None)।"""
    d = to_date(v)
    if d:
        return d
    s = txt(v)
    if s:
        m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$", s)
        if m:
            dd, mm, yy = int(m[1]), int(m[2]), int(m[3])
            yy += 2000 if yy < 100 else 0
            try:
                return datetime(yy, mm, dd)
            except ValueError:
                return None
    return None


def _dt(v):
    d = to_ddmmyy(v)
    return d if d and d.year >= 2000 else None


def _note(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%d.%m.%y")
    return txt(v)


def _rec(source, sheet, country, **kw):
    r = dict.fromkeys(EXPORT_COLS)
    r.update(Source=source, Sheet=sheet, Country=country)
    r.update(kw)
    return r


def _country(r):
    """দেশের নামের সারি কিনা (শুধু নাম আছে, অর্ডার নম্বর নেই)।"""
    return (len(r) > 2 and r[0] is None and isinstance(r[1], str) and r[1].strip()
            and r[2] is None and not r[1].strip().lower().startswith("total"))


def _demand_cols(hdr):
    m = {}
    for i, h in enumerate(hdr):
        if not h:
            continue
        if h.startswith("country"):
            m.setdefault("country", i)
        elif "sales document" in h or h.startswith("order no"):
            m.setdefault("so", i)
        elif "name of product" in h or "product name" in h:
            m.setdefault("prod", i)
        elif "product code" in h:
            m.setdefault("code", i)
        elif "qty" in h:
            m.setdefault("qty", i)
            m["unit"] = "Box" if "box" in h else "Pcs"
        elif "document date" in h or "receiving date" in h:
            m.setdefault("odate", i)
        elif "delivery date" in h:
            m.setdefault("ddate", i)
        elif "aging" in h:
            m.setdefault("aging", i)
        elif "pkg status" in h:
            m.setdefault("pkg", i)
        elif "mfg status" in h or h == "status":
            m.setdefault("status", i)
        elif "remark" in h:
            m.setdefault("rem", i)
    return m


def parse_export(wb):
    """Export Order Status ফাইল: Open order, চলতি মাসের Completed, আর সব মাসের Demand শিট।"""
    recs, done_sheet = [], False
    for ws in wb.worksheets:
        title = (ws.title or "").strip()
        low = title.lower()
        rows = list(ws.iter_rows(min_row=1, max_row=150, max_col=21, values_only=True))
        rows = [tuple(r) + (None,) * (21 - len(r)) for r in rows]
        if low.startswith("current export order"):          # চলমান অর্ডার
            ctry = None
            for r in rows:
                if _country(r):
                    ctry = clean_country(r[1])
                    continue
                if not isinstance(r[2], (int, float)) or txt(r[4]) is None or isinstance(r[4], (int, float)):
                    continue
                recs.append(_rec("Open", title, ctry, **{
                    "SAP Code": txt(r[1]), "Order No": txt(r[2]), "Machine": txt(r[3]),
                    "Product Name": txt(r[4]), "Pack Size": txt(r[5]), "Pcs per Box": num(r[6]),
                    "Required Box": num(r[7]), "Required Pcs": num(r[8]), "Batch Size": num(r[9]),
                    "Batches Req": num(r[10]), "Batch Number": txt(r[11]), "Status": txt(r[12]),
                    "Order Date": _dt(r[13]), "Delivery Date": _dt(r[14]), "PLSD": _note(r[15]),
                    "RM Available": _note(r[16]), "PM Available": _note(r[17]),
                    "Completed Date": _dt(r[18]), "Remarks": _note(r[19])}))
        elif low.startswith("completed (") and not done_sheet:   # শুধু সবচেয়ে নতুন Completed শিট
            done_sheet, ctry = True, None
            for r in rows:
                if _country(r):
                    ctry = clean_country(r[1])
                    continue
                if not isinstance(r[2], (int, float)) or not isinstance(r[3], str):
                    continue
                o = 0 if any(ch.isdigit() for ch in r[3]) else 1   # মেশিনের নাম থাকলে এক কলাম সরে যায়
                notes = " | ".join(n for n in (_note(x) for x in r[15 + o:21]) if n) or None
                recs.append(_rec("Completed", title, ctry, **{
                    "SAP Code": txt(r[1]), "Order No": txt(r[2]), "Machine": txt(r[3]) if o else None,
                    "Product Name": txt(r[3 + o]), "Pack Size": txt(r[4 + o]), "Pcs per Box": num(r[5 + o]),
                    "Required Box": num(r[6 + o]), "Required Pcs": num(r[7 + o]), "Batch Size": num(r[8 + o]),
                    "Batches Req": num(r[9 + o]), "Batch Number": txt(r[10 + o]), "Status": txt(r[11 + o]),
                    "Order Date": _dt(r[12 + o]), "Completed Date": _dt(r[13 + o]),
                    "Delivery Date": _dt(r[14 + o]), "Remarks": notes}))
        elif low.startswith("demand") or low.startswith("export order"):   # Demand শিট (সব মাস)
            hdr_i = next((j for j, r in enumerate(rows[:5])
                          if any(str(c or "").strip().lower().startswith("country") for c in r)), None)
            if hdr_i is None:
                continue
            m = _demand_cols([str(c).strip().lower() if c is not None else "" for c in rows[hdr_i]])
            if not all(k in m for k in ("country", "prod", "qty")):
                continue
            unit, ctry = m.get("unit", "Pcs"), None

            def cell(r, key):
                return r[m[key]] if key in m else None
            for r in rows[hdr_i + 1:]:
                if txt(r[m["country"]]):
                    ctry = clean_country(r[m["country"]])
                q = num(r[m["qty"]])
                if q is None or txt(r[m["prod"]]) is None:
                    continue
                rem = " | ".join(n for n in (("PKG: " + txt(cell(r, "pkg"))) if txt(cell(r, "pkg")) else None,
                                             _note(cell(r, "rem"))) if n) or None
                recs.append(_rec("Demand", title, ctry, **{
                    "SAP Code": _note(cell(r, "code")), "Order No": _note(cell(r, "so")), "Pack Size": unit,
                    "Product Name": txt(r[m["prod"]]),
                    "Required Box": q if unit == "Box" else None, "Required Pcs": q if unit == "Pcs" else None,
                    "Status": _note(cell(r, "status")), "Order Date": _dt(cell(r, "odate")),
                    "Delivery Date": _dt(cell(r, "ddate")), "SO Aging": _note(cell(r, "aging")),
                    "Remarks": rem}))
    return recs


def _cell_txt(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:.4g}"
    return re.sub(r"\s+", " ", str(v)).strip()


def parse_export_raw(wb):
    """ফাইলের প্রতিটা শিট যেমন আছে তেমন (সব কলাম, সব সারি) — শুধু দেখার জন্য।"""
    from openpyxl.utils import get_column_letter
    out = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(min_row=1, max_row=250, max_col=26, values_only=True))
        grid = [[_cell_txt(v) for v in r] for r in rows]
        grid = [r for r in grid if any(r)]
        if not grid:
            continue
        width = max(len(r) for r in grid)
        grid = [r + [""] * (width - len(r)) for r in grid]
        keep = [i for i in range(width) if any(r[i] for r in grid)]
        out[(ws.title or "").strip()] = pd.DataFrame([[r[i] for i in keep] for r in grid],
                                                     columns=[get_column_letter(i + 1) for i in keep])
    return out


def status_group(s):
    t = s.lower().strip() if isinstance(s, str) else ""
    for key, name in (("complete", "Completed"), ("hold", "Hold"), ("need to take", "Batch need to take"),
                      ("running", "Packing running"), ("blister", "Ready for Blister"),
                      ("bilister", "Ready for Blister"), ("compress", "Ready for Compression"),
                      ("granul", "Ready for Granulation"), ("coating", "Ready for Coating"),
                      ("packing", "Ready for Packing")):
        if key in t:
            return name
    return "Other" if t else "Unknown"


@st.cache_data(show_spinner="ফাইল পড়া হচ্ছে...")
def load_files(files):
    prod, tgt, bat, exp, expraw, errors = [], [], [], [], [], []
    for name, data in files:
        try:
            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            if "Production" in wb.sheetnames:  # এই সাইট থেকে ডাউনলোড করা ফাইল
                x = pd.ExcelFile(io.BytesIO(data))
                prod.append(x.parse("Production").reindex(columns=PROD_COLS))
                if "Daily Target" in x.sheet_names:
                    tgt.append(x.parse("Daily Target").reindex(columns=TARGET_COLS))
                if "Batch Release" in x.sheet_names:
                    bat.append(x.parse("Batch Release").reindex(columns=BATCH_COLS))
                if "Export Orders" in x.sheet_names:
                    exp.append(x.parse("Export Orders").reindex(columns=EXPORT_COLS))
                continue
            if any(n.strip().lower().startswith("current export order") for n in wb.sheetnames):
                exp.append(pd.DataFrame(parse_export(wb), columns=EXPORT_COLS))  # Export Order Status ফাইল
                expraw.append(parse_export_raw(wb))
                continue
            t = parse_targets(wb)
            if t:
                tgt.append(pd.DataFrame(t, columns=TARGET_COLS))
                continue
            hdr, names = find_batch_header(wb.worksheets[0])
            if hdr is not None:
                bat.append(pd.DataFrame(parse_batches(wb, hdr, names), columns=BATCH_COLS))
            else:
                prod.append(pd.DataFrame(parse_production(wb), columns=PROD_COLS))
        except Exception as e:  # noqa: BLE001
            errors.append(f"{name}: {e}")

    def cat(frames, cols):
        frames = [f for f in frames if len(f)]
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=cols)

    P, T, B = cat(prod, PROD_COLS).drop_duplicates(), cat(tgt, TARGET_COLS), cat(bat, BATCH_COLS)
    T = T.drop_duplicates("Date", keep="last") if len(T) else T
    B = B.drop_duplicates(["Date", "Batch No", "Product Name"]) if len(B) else B
    # একাধিক Export ফাইল থাকলে সবচেয়ে শেষেরটা (নতুন আপলোড) নেওয়া হয়, যাতে পুরনো/নতুন ভার্সন দ্বিগুণ না হয়
    X = exp[-1] if exp else pd.DataFrame(columns=EXPORT_COLS)
    return {"prod": P, "tgt": T, "bat": B, "exp": X, "expraw": (expraw[-1] if expraw else {})}, errors


# ---------- হিসাব ----------
def prep_prod(df):
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()]
    for c in ("Output", "Hours", "Std Capacity", "Actual Capacity", "Manpower", "Batch Size"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Output"] = df["Output"].fillna(0)
    df["Hours"] = df["Hours"].fillna(0)
    for c in ("Stage", "Line", "Shift", "Product Name", "Batch No", "Run Start", "Run End"):
        df[c] = df[c].fillna("-").astype(str)
    df["Month"] = df["Date"].dt.strftime("%Y-%m")
    df["Week"] = df["Date"].dt.strftime("%G-W%V")
    return df


def prep_tgt(df):
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()]
    for c in ("Daily Target", "Daily Output"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Daily Target"] = df["Daily Target"].where(df["Daily Target"] > 0)
    df["Month"] = df["Date"].dt.strftime("%Y-%m")
    return df.sort_values("Date")


def prep_bat(df):
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()]
    df["Month"] = df["Date"].dt.strftime("%Y-%m")
    return df


def achievement(df):
    m = df["Actual Capacity"] > 0
    if not m.any():
        return None
    return df.loc[m, "Output"].sum() / df.loc[m, "Actual Capacity"].sum() * 100


def summarize(df, by):
    d = df.assign(_oc=df["Output"].where(df["Actual Capacity"] > 0),
                  _cap=df["Actual Capacity"].where(df["Actual Capacity"] > 0))
    out = d.groupby(by).agg(Output=("Output", "sum"), Hours=("Hours", "sum"),
                            oc=("_oc", "sum"), cap=("_cap", "sum"), Entries=("Output", "size"))
    out["Achievement %"] = (out["oc"] / out["cap"] * 100).where(out["cap"] > 0).round(1)
    out["Output/Hr"] = (out["Output"] / out["Hours"]).where(out["Hours"] > 0).round(0)
    out["Output"] = out["Output"].round(0).astype("int64")
    return out.drop(columns=["oc", "cap"]).round({"Hours": 1})



RUN_COLS = ["Date", "Stage", "Line", "Shift", "Product Name", "Batch No", "Run Start", "Run End", "Hours", "Output"]


def nb(series):
    """আলাদা Batch No কয়টা (ফাঁকা '-' বাদ)।"""
    return series[series != "-"].nunique()


def join_set(series):
    return ", ".join(sorted({str(x) for x in series if str(x) != "-"})) or "-"


def fmt_day(d):
    return d if isinstance(d, str) else d.strftime("%d %b %Y, %a")


def run_table(df):
    t = df.sort_values(["Date", "Stage", "Line", "Run Start"])[RUN_COLS].copy()
    t["Date"] = t["Date"].dt.strftime("%d %b %Y")
    t["Output"] = t["Output"].round(0).astype("int64")
    t["Hours"] = t["Hours"].round(2)
    return t


def csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


# =====================================================================
# স্টাইল
# =====================================================================
PALETTE = ["#3B82F6", "#FFB020", "#22D3EE", "#A78BFA", "#34D399", "#F472B6", "#FB7185", "#60A5FA"]
STAGE_COL = {"Primary": "#3B82F6", "Secondary": "#FFB020"}
ACCENTS = ["#3B82F6", "#FFB020", "#22D3EE", "#A78BFA"]
ACH_CFG = {"Achievement %": st.column_config.ProgressColumn(
    "Achievement %", format="%.1f%%", min_value=0, max_value=100)}

# ---------- সব টেবিল: কলাম অনুযায়ী রঙিন ব্যাকগ্রাউন্ড, গাঢ় পরিষ্কার লেখা, বোল্ড হেডার ----------
BODY_BG = ["#e0f2fe", "#fef9c3", "#dcfce7", "#fce7f3", "#ede9fe", "#ffedd5", "#cffafe", "#f1f5f9"]
HEAD_BG = ["#7dd3fc", "#fde047", "#86efac", "#f9a8d4", "#c4b5fd", "#fdba74", "#67e8f9", "#cbd5e1"]
SOFT_BG = [("#bfdbfe", "#1e3a8a"), ("#fde68a", "#78350f"), ("#bbf7d0", "#14532d"), ("#fbcfe8", "#831843"),
           ("#ddd6fe", "#4c1d95"), ("#a5f3fc", "#155e75")]
LAKH_COLS = {"Output", "Daily Output", "Total Output", "Daily Target", "Target", "Monthly Target",
             "Daily Target Sum", "Output/Hr", "Output / Hr"}
HASH_COLS = ["Country", "Line", "Machine", "Shift", "Type", "Month", "Week", "Day"]
TBL_CSS = (
    "<style>"
    ".tbl-wrap{overflow:auto;border:1px solid #94a3b8;border-radius:10px;background:#fff}"
    ".tbl-wrap table{border-collapse:collapse;width:100%;font-size:.84rem;color:#0f172a}"
    ".tbl-wrap thead th{position:sticky;top:0;z-index:2;text-align:left;padding:.5rem .7rem;font-weight:800;"
    "font-size:.88rem;border-bottom:2px solid #475569;white-space:nowrap;color:#0f172a}"
    ".tbl-wrap tbody th{padding:.4rem .7rem;font-weight:800;background:#e2e8f0;color:#0f172a;text-align:left}"
    ".tbl-wrap td{padding:.38rem .7rem;border-bottom:1px solid rgba(15,23,42,.12);white-space:nowrap;"
    "color:#0f172a;font-weight:500}"
    ".tbl-wrap tbody tr:nth-child(even) td{filter:brightness(.96)}"
    "</style>")


def _smap(sty, fn, cols=None):
    f = getattr(sty, "map", None) or sty.applymap
    return f(fn, subset=cols) if cols else f(fn)


def _hash_i(v):
    return sum(ord(ch) for ch in str(v))


def _soft_css(v):
    if str(v) in ("", "-", "nan", "None"):
        return ""
    bg, fg = SOFT_BG[_hash_i(v) % len(SOFT_BG)]
    return f"background-color:{bg};color:{fg};font-weight:700"


def _stage_css(v):
    return ("background-color:#93c5fd;color:#1e3a8a;font-weight:800" if v == "Primary" else
            "background-color:#fcd34d;color:#78350f;font-weight:800" if v == "Secondary" else "")


def _ach_css(v):
    if pd.isna(v):
        return ""
    p = min(max(float(v), 0), 100)
    return f"background:linear-gradient(90deg,#4ade80 {p}%,#f1f5f9 {p}%);color:#0f172a;font-weight:800"


def html_table(df, extra=None, hide_index=True, height=560):
    """রঙিন HTML টেবিল: কলাম অনুযায়ী ব্যাকগ্রাউন্ড, বোল্ড হেডার, গাঢ় লেখা।"""
    sty = df.style
    if hide_index:
        try:
            sty = sty.hide(axis="index")
        except Exception:  # noqa: BLE001
            sty = sty.hide_index()
    from html import escape as _esc

    def f_txt(v):
        return "-" if (v is None or pd.isna(v)) else _esc(str(v))

    def f_pct(v):
        return "-" if pd.isna(v) else f"{v:.1f}%"

    def f_hrs(v):
        return "-" if pd.isna(v) else f"{v:.2f}"
    fm = {}
    for c in df.columns:
        num_ = pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_bool_dtype(df[c])
        fm[c] = (fmt_full if num_ and c in LAKH_COLS else f_pct if num_ and c == "Achievement %"
                 else f_hrs if num_ and c == "Hours" else f_txt)
    sty = sty.format(fm)
    head = []
    for i, c in enumerate(df.columns):
        sty = sty.set_properties(subset=[c], **{"background-color": BODY_BG[i % 8], "color": "#0f172a"})
        head.append({"selector": f"th.col_heading.col{i}",
                     "props": [("background-color", HEAD_BG[i % 8]), ("font-weight", "800")]})
    sty = sty.set_table_styles(head, overwrite=False)
    hc = [c for c in HASH_COLS if c in df.columns]
    if hc:
        sty = _smap(sty, _soft_css, hc)
    if "Stage" in df:
        sty = _smap(sty, _stage_css, ["Stage"])
    if "Achievement %" in df:
        sty = _smap(sty, _ach_css, ["Achievement %"])
    if {"Daily Output", "Daily Target"} <= set(df.columns):
        def met(col):
            return ["background-color:#86efac;color:#14532d;font-weight:800" if pd.notna(t) and pd.notna(o) and o >= t
                    else "background-color:#fcd34d;color:#78350f;font-weight:800" if pd.notna(t) and pd.notna(o)
                    else "" for o, t in zip(col, df["Daily Target"])]
        sty = sty.apply(met, subset=["Daily Output"])
    for kind, fn, cols in (extra or []):
        if kind == "map":
            sty = _smap(sty, fn, [c for c in cols if c in df.columns]) if any(c in df.columns for c in cols) else sty
        elif kind == "rows":
            sty = sty.apply(fn, axis=1)
    html = "".join(line.strip() for line in sty.to_html().splitlines())
    return f'{TBL_CSS}<div class="tbl-wrap" style="max-height:{int(height)}px">{html}</div>'


def cdf(data, extra=None, **kw):
    """st.dataframe-এর বদলে: রঙিন টেবিল (বন্ধ করলে বা বড় টেবিল হলে সাধারণ ইন্টারঅ্যাকটিভ টেবিল)।"""
    if (isinstance(data, pd.DataFrame) and not data.empty and data.size <= 5000
            and st.session_state.get("color_tbl", True)):
        try:
            st.markdown(html_table(data, extra=extra, hide_index=kw.get("hide_index", False),
                                   height=kw.get("height") or 560), unsafe_allow_html=True)
            return
        except Exception:  # noqa: BLE001
            pass
    st.dataframe(data, **kw)


st.markdown("""
<style>
.stApp {background: radial-gradient(1200px 600px at 10% -10%, #17356B 0%, #0B1730 55%, #081226 100%);
        color:#E6EEFB;}
header[data-testid="stHeader"] {background: transparent;}
.stApp, .stApp p, .stApp label, .stApp span, .stApp li, .stApp h1, .stApp h2, .stApp h3, .stApp h4 {color:#E6EEFB;}
/* ইনপুট ঘর: সব সময় হালকা ব্যাকগ্রাউন্ড ও গাঢ় লেখা (যে থিমেই হোক পড়া যায়) */
.stApp div[data-baseweb="select"] > div, .stApp div[data-baseweb="input"], .stApp div[data-baseweb="base-input"],
.stApp input, .stApp textarea {background:#EAF1FF !important;}
.stApp [data-baseweb="select"] *, .stApp input, .stApp textarea {color:#10224A !important;}
div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li {background:#EAF1FF !important;}
div[data-baseweb="popover"] li, div[data-baseweb="popover"] li * {color:#10224A !important;}
div[data-baseweb="popover"] li:hover {background:#CFE0FF !important;}
div[data-testid="stFileUploaderDropzone"] {background: rgba(255,255,255,.08); border: 1px dashed rgba(96,165,250,.5);}
section[data-testid="stSidebar"] {background: #0F2145;}
.block-container {padding-top: 1rem; max-width: 1250px;}
footer {visibility: hidden;}
/* ব্যানার */
.hero {position: relative; overflow: hidden; padding: 26px 28px; border-radius: 20px; margin-bottom: 18px;
       background: linear-gradient(120deg,#12306A 0%,#0E2250 60%,#0B1B3F 100%);
       border: 1px solid rgba(90,140,230,.35); box-shadow: 0 10px 28px rgba(0,0,0,.35);}
.hero:before {content:""; position:absolute; inset:0; background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='900' height='220' viewBox='0 0 900 220'><path d='M0 170 C150 60 300 210 450 120 S720 40 900 120' stroke='%233B82F6' stroke-width='2' fill='none' opacity='.55'/><path d='M0 190 C170 90 320 220 470 150 S740 70 900 150' stroke='%23FFB020' stroke-width='1.6' fill='none' opacity='.45'/><path d='M0 150 C140 40 310 190 460 100 S700 20 900 90' stroke='%2322D3EE' stroke-width='1.2' fill='none' opacity='.35'/></svg>");
       background-repeat:no-repeat; background-position: right bottom; background-size: cover; opacity:.9;}
.hero-t {position:relative; font-size: 1.9rem; font-weight: 800; letter-spacing:.3px; color:#fff !important;}
.hero-t:after {content:""; display:block; width:64px; height:4px; border-radius:4px; background:#FFB020; margin-top:8px;}
.hero-s {position:relative; color:#A9BEE6 !important; font-size:.96rem; margin-top:10px;}
.chips {position:relative; margin-top:12px;}
.chip {display:inline-block; background: rgba(59,130,246,.18); border:1px solid rgba(96,165,250,.45);
       color:#CFE0FF !important; padding: 4px 12px; border-radius: 999px; font-size:.78rem; margin: 0 6px 6px 0; font-weight:600;}
/* KPI কার্ড */
.kpi {position:relative; overflow:hidden; padding: 16px 18px; border-radius: 16px; margin-bottom: 12px; min-height: 112px;
      background: linear-gradient(160deg,#14305F,#0F2447); border: 1px solid rgba(90,140,230,.28);
      border-left: 5px solid var(--ac); box-shadow: 0 6px 18px rgba(0,0,0,.30);}
.kpi:after {content:""; position:absolute; right:-24px; top:-24px; width:96px; height:96px; border-radius:50%;
      background: var(--ac); opacity:.14;}
.kpi-l {font-size:.74rem; font-weight:700; text-transform:uppercase; letter-spacing:.8px; color: var(--ac) !important;}
.kpi-v {font-size: 1.8rem; font-weight: 800; margin-top: 4px; line-height: 1.15; color:#fff !important;}
.kpi-s {font-size:.8rem; margin-top:3px; color:#9FB4DD !important;}
/* সেকশন হেডার */
.sec {font-size: 1.02rem; font-weight: 700; margin: 24px 0 10px 0; padding: 8px 14px; letter-spacing:.4px;
      border-left: 5px solid #FFB020; border-radius: 8px; color:#fff !important;
      background: linear-gradient(90deg, rgba(59,130,246,.25), rgba(59,130,246,0));}
/* চার্ট ও টেবিল কার্ড */
div[data-testid="stPlotlyChart"] {background: linear-gradient(160deg,#122A55,#0E2147); border-radius:16px;
      padding:8px 6px; border: 1px solid rgba(90,140,230,.28); box-shadow: 0 6px 18px rgba(0,0,0,.30); margin-bottom: 10px;}
div[data-testid="stDataFrame"] {border-radius: 12px; border: 1px solid rgba(90,140,230,.28); overflow:hidden;
      box-shadow: 0 4px 14px rgba(0,0,0,.30);}
/* ট্যাব */
div[data-baseweb="tab-list"] {gap: 8px; flex-wrap: wrap;}
div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] {display:none;}
button[data-baseweb="tab"] {font-weight:700; padding: 9px 16px; border-radius: 12px;
      background: rgba(255,255,255,.06); border: 1px solid rgba(96,165,250,.30);}
button[data-baseweb="tab"][aria-selected="true"] {background:#FFB020; border-color:#FFB020;
      box-shadow: 0 4px 14px rgba(255,176,32,.35);}
button[data-baseweb="tab"][aria-selected="true"] p {color:#0B1730 !important;}
/* বাটন */
.stButton > button, .stDownloadButton > button, div[data-testid="stFormSubmitButton"] > button {
      border-radius: 12px; font-weight: 800; border: 0; background:#FFB020;}
.stButton > button p, .stDownloadButton > button p, div[data-testid="stFormSubmitButton"] > button p {color:#0B1730 !important;}
div[data-testid="stForm"] {border: 1px solid rgba(90,140,230,.30); border-radius: 16px; background: rgba(15,36,71,.6);}
</style>
""", unsafe_allow_html=True)


def style(fig, h=370):
    fig.update_layout(height=h, margin=dict(l=10, r=10, t=48, b=10), paper_bgcolor="rgba(0,0,0,0)",
                      plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#D6E2F7", size=12), colorway=PALETTE,
                      legend=dict(orientation="h", y=-0.2, x=0), title_font=dict(size=16, color="#FFFFFF"))
    fig.update_xaxes(gridcolor="rgba(120,160,230,.16)", zeroline=False, linecolor="rgba(120,160,230,.35)",
                     automargin=True)
    fig.update_yaxes(gridcolor="rgba(120,160,230,.16)", zeroline=False, linecolor="rgba(120,160,230,.35)",
                     automargin=True)
    fig.update_layout(bargap=0.3)
    for tr in fig.data:  # "2026-11" বা "2026-W45" লেখাকে তারিখ ভেবে ভুল করা ঠেকায়
        x = getattr(tr, "x", None)
        if x is not None and len(x) and isinstance(x[0], str) and re.match(r"^\d{4}-(W?\d{2})$", x[0]):
            fig.update_xaxes(type="category")
            break
    return fig


def show(fig, h=370):
    edit = st.session_state.get("chart_edit", False)
    if edit:  # জুম/মুভ/ডাউনলোড চালু
        cfg = {"displayModeBar": True, "scrollZoom": False, "displaylogo": False, "doubleClick": "reset"}
    else:  # লক: স্ক্রল করলে চার্ট নড়বে না
        cfg = {"staticPlot": True}
    st.plotly_chart(style(fig, h), theme=None, config=cfg)


def fmt(n):
    if n is None or pd.isna(n):
        return "-"
    n = float(n)
    if abs(n) >= 1e6:
        return f"{n / 1e6:.2f}M"
    if abs(n) >= 1e4:
        return f"{n / 1e3:.1f}K"
    return f"{n:,.0f}"


def fmt_full(n):
    """পুরো সংখ্যা, বাংলাদেশি/ভারতীয় কমা-সহ: 1520000 → 15,20,000; 15200000 → 1,52,00,000।"""
    if n is None or pd.isna(n):
        return "-"
    n = int(round(float(n)))
    s = str(abs(n))
    if len(s) > 3:
        s = re.sub(r"(\d)(?=(\d\d)+$)", r"\1,", s[:-3]) + "," + s[-3:]
    return ("-" if n < 0 else "") + s


def kpi(col, label, value, sub="", g=0):
    col.markdown(f'<div class="kpi" style="--ac:{ACCENTS[g % 4]}"><div class="kpi-l">{label}</div>'
                 f'<div class="kpi-v">{value}</div><div class="kpi-s">{sub}</div></div>',
                 unsafe_allow_html=True)


def sec_title(t):
    st.markdown(f'<div class="sec">{t}</div>', unsafe_allow_html=True)


def hero(sub, chips=()):
    ch = "".join(f'<span class="chip">{c}</span>' for c in chips)
    st.markdown(f'<div class="hero"><div class="hero-t">🏭 BSU Production Tracker</div>'
                f'<div class="hero-s">{sub}</div><div class="chips">{ch}</div></div>',
                unsafe_allow_html=True)


# =====================================================================
# Login (ID / Password): Streamlit Secrets থেকে নেয়
# =====================================================================
def get_secret(section):
    try:
        return dict(st.secrets[section])
    except Exception:  # noqa: BLE001
        return None


ADMIN, VIEWER, GH = get_secret("admin"), get_secret("viewer"), get_secret("github")
LOGIN_ON = bool(ADMIN or VIEWER)


def same(a, b):
    return hmac.compare_digest(str(a).encode(), str(b).encode())


def try_login(u, p):
    for role, acc in (("admin", ADMIN), ("viewer", VIEWER)):
        if acc and acc.get("password") and same(u.strip(), acc.get("username", "")) and same(p, acc["password"]):
            return role
    return None


if not LOGIN_ON:
    st.session_state.role = "admin"  # Secrets না দেওয়া পর্যন্ত সবার জন্য খোলা
if "role" not in st.session_state:
    hero("Primary &amp; Secondary Packaging · Daily Target · Batch Release")
    mid = st.columns([1, 2, 1])[1]
    with mid:
        with st.form("login"):
            st.markdown("#### 🔐 Login")
            u = st.text_input("Username")
            p = st.text_input("Password", type="password")
            if st.form_submit_button("Login"):
                r = try_login(u, p)
                if r:
                    st.session_state.role = r
                    st.rerun()
                else:
                    st.error("Username বা Password ভুল।")
    st.stop()

is_admin = st.session_state.role == "admin"


# =====================================================================
# GitHub-এ স্থায়ীভাবে সেভ
# =====================================================================
def gh_ready():
    return bool(GH and GH.get("token") and GH.get("repo"))


def gh_put(path, data, message):
    url = f"https://api.github.com/repos/{GH['repo']}/contents/{quote(path)}"
    branch = GH.get("branch", "main")
    hdr = {"Authorization": f"Bearer {GH['token']}", "Accept": "application/vnd.github+json"}
    old = requests.get(url, headers=hdr, params={"ref": branch}, timeout=30)
    body = {"message": message, "content": base64.b64encode(data).decode(), "branch": branch}
    if old.status_code == 200:
        body["sha"] = old.json()["sha"]
    r = requests.put(url, headers=hdr, json=body, timeout=90)
    if r.status_code >= 300:
        raise RuntimeError(f"GitHub {r.status_code}: {r.json().get('message', '')}")


# =====================================================================
# সাইডবার ও ডেটা লোড
# =====================================================================
MANUAL_FILE = DATA_DIR / "manual_entries.xlsx"


def autosave_on():
    return gh_ready() and st.session_state.get("autosave", True)


def manual_bytes(df):
    mb, mx = io.BytesIO(), df.copy()
    mx["Date"] = pd.to_datetime(mx["Date"], errors="coerce").dt.date
    mx.to_excel(mb, sheet_name="Production", index=False)
    return mb.getvalue()


def save_manual_now():
    """হাতে দেওয়া এন্ট্রি বদলালেই নিজে থেকে GitHub-এ সেভ (অটো-সেভ চালু থাকলে)।"""
    if autosave_on():
        try:
            gh_put(f"data/{MANUAL_FILE.name}", manual_bytes(st.session_state.manual), "Update manual entries")
            st.session_state.flash = ("ok", "✅ এন্ট্রি নিজে থেকে সেভ হয়েছে। ১-২ মিনিটে সাইট আপডেট হবে।")
        except Exception as e:  # noqa: BLE001
            st.session_state.flash = ("err", f"সেভ হয়নি: {e}")


def save_mt():
    """মাসিক Target সেভ: GitHub থাকলে সেখানে, না থাকলে ডাউনলোডের ব্যবস্থা।"""
    data = json.dumps(st.session_state.mt_raw, ensure_ascii=False, indent=1).encode("utf-8")
    if autosave_on():
        try:
            gh_put(f"data/{MT_FILE.name}", data, "Update monthly targets")
            st.session_state.flash = ("ok", "✅ Target সেভ হয়েছে।")
        except Exception as e:  # noqa: BLE001
            st.session_state.flash = ("err", f"Target সেভ হয়নি: {e}")
    else:
        st.session_state.flash = ("ok", "Target বসানো হয়েছে (এই সেশনে)। স্থায়ী করতে 'Data' ট্যাব থেকে monthly_targets.json "
                                        "নামিয়ে GitHub-এর data ফোল্ডারে দিন।")
uploads = []
with st.sidebar:
    st.markdown("### 🏭 BSU Tracker")
    if LOGIN_ON:
        st.caption("Login: " + ("Admin" if is_admin else "Viewer"))
        if st.button("Logout"):
            st.session_state.pop("role", None)
            st.rerun()
    else:
        st.warning("Login চালু নেই। Streamlit Secrets-এ admin/viewer দিলে চালু হবে।")
    if is_admin:
        st.markdown("#### 📁 ডেটা ফাইল")
        if gh_ready():
            st.checkbox("🔄 নিজে থেকে সেভ (GitHub)", value=True, key="autosave")
        uploads = st.file_uploader("Excel ফাইল (একাধিক দেওয়া যায়)", type=["xlsx"],
                                   accept_multiple_files=True)
        if uploads and autosave_on():
            saved = st.session_state.setdefault("saved_uploads", {})
            for up in uploads:
                if (up.name, up.size) not in saved:
                    try:
                        gh_put(f"data/{up.name}", up.getvalue(), f"Add {up.name}")
                        saved[(up.name, up.size)] = True
                    except Exception as e:  # noqa: BLE001
                        st.error(f"{up.name}: {e}")
            if saved:
                st.caption("✅ সেভ হয়েছে: " + ", ".join(n for n, _ in saved))
        elif uploads:
            if gh_ready():
                if st.button("💾 স্থায়ীভাবে সেভ করুন"):
                    done, bad = [], []
                    for up in uploads:
                        try:
                            gh_put(f"data/{up.name}", up.getvalue(), f"Add {up.name}")
                            done.append(up.name)
                        except Exception as e:  # noqa: BLE001
                            bad.append(f"{up.name}: {e}")
                    if done:
                        st.success("সেভ হয়েছে। ১-২ মিনিটে সাইট নিজে আপডেট হবে।")
                    for b in bad:
                        st.error(b)
            else:
                st.caption("স্থায়ী করতে ফাইলগুলো GitHub-এর data ফোল্ডারে আপলোড করুন "
                           "(বা Secrets-এ github token দিন)।")

repo_paths = [p for p in sorted(DATA_DIR.iterdir())  # .xlsx আর .XLSX দুটোই পড়বে
              if p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
              and p.name != MANUAL_FILE.name] if DATA_DIR.exists() else []
if DEFAULT_FILE.exists():
    repo_paths.append(DEFAULT_FILE)
files = tuple([(p.name, p.read_bytes()) for p in repo_paths] + [(u.name, u.getvalue()) for u in uploads])

key = tuple((n, len(d)) for n, d in files)
if st.session_state.get("files_key") != key:
    st.session_state.data, st.session_state.errs = load_files(files)
    st.session_state.files_key = key
for e in st.session_state.get("errs", []):
    st.sidebar.warning(e)

if "manual" not in st.session_state:
    m = pd.DataFrame(columns=PROD_COLS)
    if MANUAL_FILE.exists():
        try:
            m = pd.read_excel(MANUAL_FILE, sheet_name="Production").reindex(columns=PROD_COLS)
        except Exception:  # noqa: BLE001
            pass
    st.session_state.manual = m

D = st.session_state.data
frames = [f for f in (D["prod"], st.session_state.manual) if len(f)]
prod = prep_prod(pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=PROD_COLS))
tgt, bat = prep_tgt(D["tgt"]), prep_bat(D["bat"])

if "mt_raw" not in st.session_state:
    raw = dict(DEFAULT_MONTHLY_TARGETS)
    try:
        if MT_FILE.exists():
            raw.update({str(k): int(v) for k, v in json.loads(MT_FILE.read_text("utf-8")).items()})
    except Exception:  # noqa: BLE001
        pass
    # ভুল করে বসানো পুরনো মান (যেমন নভেম্বর'26-এ 17,83,00,000) নিজে থেকে ঠিক মানে বদলে যাবে
    for _m, (_bad, _good) in TARGET_FIXES.items():
        if raw.get(_m) == _bad:
            raw[_m] = _good
    st.session_state.mt_raw = raw
MT = {k: v for k, v in st.session_state.mt_raw.items() if v and v > 0}  # মাস → মাসিক Target

def _span(df, label):
    if df.empty:
        return None
    return f"{label}: {df['Date'].min():%d %b %Y} – {df['Date'].max():%d %b %Y}"


hero("Primary &amp; Secondary Packaging · Daily Target vs Output · Batch Release",
     [c for c in (_span(prod, "📊 Production"), _span(tgt, "🎯 Target"), _span(bat, "📦 Batch")) if c])
fl = st.session_state.pop("flash", None)
if fl:
    (st.success if fl[0] == "ok" else st.error)(fl[1])
st.toggle("🖱️ চার্ট জুম/মুভ চালু করুন (স্ক্রল করার সময় বন্ধ রাখুন)", key="chart_edit")
st.toggle("🎨 রঙিন টেবিল (বোল্ড হেডার)", value=True, key="color_tbl")
labels = ["📊 Production", "📋 Daily Report", "🎯 Target vs Output", "📦 Batch Release", "🚢 Export Order"]
if is_admin:
    labels += ["➕ নতুন এন্ট্রি", "🗂️ ডেটা"]
tabs = st.tabs(labels)
t1, t_dr, t2, t3, t_ex = tabs[0], tabs[1], tabs[2], tabs[3], tabs[4]

# =====================================================================
# Production
# =====================================================================
with t1:
    if prod.empty:
        st.info("এখনো কোনো ডেটা নেই। Admin হিসেবে Primary/Secondary Excel ফাইল আপলোড করুন।")
    else:
        sel = st.radio("ধাপ", ["সব", "Primary", "Secondary"], horizontal=True)
        base = prod if sel == "সব" else prod[prod["Stage"] == sel]
        if sel == "সব":
            st.caption("⚠️ 'সব' দিলে Primary + Secondary Output যোগ হয়: একই প্রোডাক্ট দুই ধাপে গোনা হয়ে মোট অনেক "
                       "বেশি দেখাতে পারে। কারখানার প্রকৃত মোট Output দেখতে একটা ধাপ বেছে নিন।")
        months = sorted(base["Month"].unique())
        pick = st.selectbox("মাস", ["সব মাস"] + months, index=len(months))
        view = base if pick == "সব মাস" else base[base["Month"] == pick]

        ach = achievement(view)
        hrs = view["Hours"].sum()
        k = st.columns(4)
        kpi(k[0], "মোট Output", fmt_full(view["Output"].sum()), f"{len(view)} এন্ট্রি", 0)
        kpi(k[1], "চলার সময়", f"{hrs:,.0f} ঘণ্টা", f"{view['Date'].nunique()} দিন", 1)
        kpi(k[2], "Achievement", f"{ach:.1f}%" if ach is not None else "-", f"{sel} · {pick}", 2)
        kpi(k[3], "Output / ঘণ্টা", fmt_full(view["Output"].sum() / hrs) if hrs > 0 else "-", "গড়", 3)

        sec_title("মাসিক Output ও চলার সময়")
        mo = base.groupby(["Month", "Stage"], as_index=False).agg(Output=("Output", "sum"), Hours=("Hours", "sum"))
        c1, c2 = st.columns(2)
        with c1:
            show(px.bar(mo, x="Month", y="Output", color="Stage", barmode="group",
                        color_discrete_map=STAGE_COL, title="মাসিক Output"))
        with c2:
            show(px.bar(mo, x="Month", y="Hours", color="Stage", barmode="group",
                        color_discrete_map=STAGE_COL, title="মাসিক চলার সময় (ঘণ্টা)"))

        sec_title("দৈনিক Output")
        dly = view.groupby(["Date", "Stage"], as_index=False)["Output"].sum()
        show(px.line(dly, x="Date", y="Output", color="Stage", markers=True,
                     color_discrete_map=STAGE_COL, title="দৈনিক Output"))

        sec_title("সাপ্তাহিক Achievement")
        wk = summarize(view, ["Week", "Stage"]).reset_index()
        c1, c2 = st.columns(2)
        with c1:
            show(px.bar(wk, x="Week", y="Output", color="Stage", barmode="group",
                        color_discrete_map=STAGE_COL, title="সাপ্তাহিক Output"))
        with c2:
            show(px.line(wk, x="Week", y="Achievement %", color="Stage", markers=True,
                         color_discrete_map=STAGE_COL, title="সাপ্তাহিক Achievement %"))
        cdf(wk, column_config=ACH_CFG, hide_index=True)

        sec_title("Line / Machine ও Shift")
        ln = summarize(view, ["Stage", "Line"]).reset_index().sort_values("Output")
        c1, c2 = st.columns([3, 2])
        with c1:
            show(px.bar(ln, x="Output", y="Line", color="Stage", orientation="h",
                        color_discrete_map=STAGE_COL, title="Line অনুযায়ী Output"), 420)
        with c2:
            sh = view.groupby("Shift", as_index=False)["Output"].sum()
            show(px.pie(sh, names="Shift", values="Output", hole=0.55,
                        color_discrete_sequence=PALETTE, title="Shift অনুযায়ী Output"), 420)
        cdf(ln.sort_values("Output", ascending=False), column_config=ACH_CFG, hide_index=True)

        sec_title("Product অনুযায়ী")
        pr = summarize(view, "Product Name").sort_values("Output", ascending=False)
        top = pr.head(10).reset_index().sort_values("Output")
        fig = px.bar(top, x="Output", y="Product Name", orientation="h", title="Top 10 Product",
                     color="Output", color_continuous_scale=["#9CCBF5", "#3B82F6", "#0B3C6F"])
        fig.update_coloraxes(showscale=False)
        show(fig, 420)
        cdf(pr, column_config=ACH_CFG)

# =====================================================================
# Daily Report: Batch, Start/End, Product-Batch, Line ফিল্টার
# =====================================================================
with t_dr:
    if prod.empty:
        st.info("এখনো কোনো Production ডেটা নেই।")
    else:
        c1, c2 = st.columns(2)
        f_stage = c1.radio("ধাপ", ["সব", "Primary", "Secondary"], horizontal=True, key="dr_stage")
        pb = prod if f_stage == "সব" else prod[prod["Stage"] == f_stage]
        mlist = sorted(pb["Month"].unique())
        f_month = c2.selectbox("মাস", ["সব মাস"] + mlist, index=len(mlist), key="dr_month")
        pm = pb if f_month == "সব মাস" else pb[pb["Month"] == f_month]
        c3, c4 = st.columns(2)
        f_day = c3.selectbox("দিন", ["সব দিন"] + sorted(pm["Date"].dt.date.unique()), key="dr_day", format_func=fmt_day)
        f_line = c4.selectbox("Line / Machine", ["সব Line"] + sorted(pm["Line"].unique()), key="dr_line")
        q = st.text_input("🔍 Product বা Batch No খুঁজুন", key="dr_q")

        v = pm
        if f_day != "সব দিন":
            v = v[v["Date"].dt.date == f_day]
        if f_line != "সব Line":
            v = v[v["Line"] == f_line]
        if q.strip():
            ql = q.strip().lower()
            v = v[v["Product Name"].str.lower().str.contains(ql, regex=False)
                  | v["Batch No"].str.lower().str.contains(ql, regex=False)]

        if v.empty:
            st.warning("এই ফিল্টারে কিছু পাওয়া যায়নি।")
        else:
            k = st.columns(4)
            kpi(k[0], "রান / এন্ট্রি", f"{len(v):,}", f"{v['Date'].nunique()} দিন", 0)
            kpi(k[1], "Batch সংখ্যা", f"{nb(v['Batch No'])}", "আলাদা Batch No", 1)
            kpi(k[2], "Product", f"{v['Product Name'].nunique()}", "আলাদা Product", 2)
            kpi(k[3], "মোট Output", fmt_full(v["Output"].sum()), f"{v['Hours'].sum():,.1f} ঘণ্টা", 3)

            sec_title("দৈনিক রিপোর্ট (Batch, Start ও End সময়সহ)")
            rt = run_table(v)
            cdf(rt, hide_index=True)
            st.download_button("⬇️ এই রিপোর্ট CSV", csv_bytes(rt), "daily_report.csv", mime="text/csv")

            sec_title("Product অনুযায়ী: কয়টা Batch, কোন কোন Batch No")
            ps = v.groupby("Product Name").agg(
                Batches=("Batch No", nb), Runs=("Batch No", "size"), Output=("Output", "sum"),
                Hours=("Hours", "sum"), Lines=("Line", join_set), Batch_Numbers=("Batch No", join_set)
            ).reset_index().rename(columns={"Batch_Numbers": "Batch Numbers"})
            ps["Output"] = ps["Output"].round(0).astype("int64")
            ps["Hours"] = ps["Hours"].round(1)
            cdf(ps.sort_values("Output", ascending=False), hide_index=True)

            sec_title("একটা Product বেছে বিস্তারিত দেখুন")
            psel = st.selectbox("Product", ["—"] + sorted(v["Product Name"].unique()), key="dr_prod")
            if psel != "—":
                pv = v[v["Product Name"] == psel]
                bt = pv.groupby("Batch No").agg(
                    Dates=("Date", lambda s: ", ".join(sorted({d.strftime("%d %b") for d in s}))),
                    Lines=("Line", join_set), Runs=("Output", "size"),
                    Output=("Output", "sum"), Hours=("Hours", "sum")).reset_index()
                bt["Output"] = bt["Output"].round(0).astype("int64")
                bt["Hours"] = bt["Hours"].round(1)
                st.caption(f"{psel}: মোট {nb(pv['Batch No'])}টা Batch")
                cdf(bt, hide_index=True)

            sec_title("Line অনুযায়ী কোন কোন Product হয়েছে")
            lp = v.groupby(["Stage", "Line", "Product Name"]).agg(
                Batches=("Batch No", nb), Output=("Output", "sum"), Hours=("Hours", "sum")).reset_index()
            lp["Output"] = lp["Output"].round(0).astype("int64")
            lp["Hours"] = lp["Hours"].round(1)
            cdf(lp.sort_values(["Stage", "Line", "Output"], ascending=[True, True, False]), hide_index=True)

# =====================================================================
# Target vs Output
# =====================================================================
with t2:
    if tgt.empty:
        st.info("Daily output reporting ফাইলটা (Daily Target Output Achievement শিটসহ) আপলোড করুন।")
    else:
        valid = tgt[tgt["Daily Target"].notna() & tgt["Daily Output"].notna()]

        # ---------- মাস বেছে নিন: নিচের সব KPI এই মাসের ----------
        all_m = sorted(set(tgt["Month"]) | set(MT))
        last_m = sorted(tgt["Month"].unique())[-1]
        tm = st.selectbox("মাস", all_m, index=all_m.index(last_m), key="tm")
        dd = tgt[tgt["Month"] == tm]
        mv = valid[valid["Month"] == tm]
        out_m = float(dd["Daily Output"].sum())          # এই মাসের মোট Output (সব দিন, Target না থাকা দিনসহ)
        mt_val = MT.get(tm)                              # মাসিক Target (সেট করা)
        day_t = float(mv["Daily Target"].sum())          # দৈনিক Target-এর যোগফল (বিকল্প)
        used_t = mt_val or day_t or None
        ach_m = out_m / used_t * 100 if used_t else None

        sec_title(f"{tm} · মাসিক Target বনাম Output")
        k = st.columns(4)
        kpi(k[0], "মোট Output", fmt_full(out_m), f"{dd['Daily Output'].gt(0).sum()} দিনে", 0)
        if mt_val:
            kpi(k[1], "মাসিক Target", fmt_full(mt_val), "সেট করা Target", 1)
        else:
            kpi(k[1], "Target", fmt_full(day_t) if day_t else "-",
                "দৈনিক Target-এর যোগফল (মাসিক Target সেট নেই)" if day_t else "নিচে থেকে Target সেট করুন", 1)
        kpi(k[2], "Achievement", f"{ach_m:.1f}%" if ach_m is not None else "-", "Output ÷ Target", 2)
        if used_t:
            kpi(k[3], "বাকি" if out_m < used_t else "Target ছাড়িয়েছে",
                fmt_full(abs(used_t - out_m)), "Target পূরণে লাগবে" if out_m < used_t else "অতিরিক্ত Output", 3)
        else:
            kpi(k[3], "বাকি", "-", "", 3)
        if ach_m is not None:
            st.progress(min(max(ach_m / 100, 0.0), 1.0))

        worked = dd[dd["Daily Output"].fillna(0) > 0]
        if used_t and len(worked):
            n_work = len(worked)
            avg_d = out_m / n_work
            last_d = worked["Date"].max()
            m_end = pd.Period(tm, freq="M").end_time.normalize()
            rem_days = sum(1 for d in pd.date_range(last_d + pd.Timedelta(days=1), m_end) if d.weekday() != 4)
            need = max(used_t - out_m, 0)
            proj = out_m + avg_d * rem_days
            k2 = st.columns(4)
            kpi(k2[0], "দিনপ্রতি গড় Output", fmt_full(avg_d), f"{n_work} কর্মদিবসের গড়", 0)
            kpi(k2[1], "বাকি কর্মদিবস", f"{rem_days}", f"{last_d:%d %b}-এর পর, শুক্রবার বাদ", 1)
            kpi(k2[2], "দিনে লাগবে", fmt_full(need / rem_days) if rem_days and need > 0 else "-",
                "Target ধরতে বাকি দিনগুলোতে" if need > 0 else "Target পূরণ হয়েছে", 2)
            kpi(k2[3], "এই গতিতে মাস শেষে", fmt_full(proj), f"Target-এর {proj / used_t * 100:.1f}%", 3)

        if is_admin:
            with st.expander("🎯 মাসিক Target সেট / পরিবর্তন"):
                c1, c2 = st.columns(2)
                m_in = c1.text_input("মাস (YYYY-MM)", value=tm, key=f"mt_month_{tm}")
                v_in = c2.text_input("Target (কমা দিয়ে লিখলেও চলবে, যেমন 1,78,30,000)",
                                     value=fmt_full(mt_val) if mt_val else "", key=f"mt_val_{tm}")
                if st.button("💾 Target সেভ করুন"):
                    if not re.match(r"^\d{4}-(0[1-9]|1[0-2])$", m_in.strip()):
                        st.error("মাস লিখুন এভাবে: 2026-11")
                    else:
                        digits = re.sub(r"\D", "", v_in)
                        st.session_state.mt_raw[m_in.strip()] = int(digits) if digits else 0  # ফাঁকা = Target মুছে ফেলা
                        save_mt()
                        st.rerun()

        # ---------- সব মাসের তুলনা ----------
        mon = tgt.groupby("Month").agg(**{"Total Output": ("Daily Output", "sum")})
        vm = valid.groupby("Month").agg(**{"Daily Target Sum": ("Daily Target", "sum"),
                                           "Output on Target Days": ("Daily Output", "sum")})
        mon = mon.join(vm, how="left").reset_index()
        mon["Monthly Target"] = mon["Month"].map(MT)
        has_mt = mon["Monthly Target"].notna()
        mon["Target"] = mon["Monthly Target"].where(has_mt, mon["Daily Target Sum"])
        mon["Output"] = mon["Total Output"].where(has_mt, mon["Output on Target Days"])
        mon["Achievement %"] = (mon["Output"] / mon["Target"] * 100).where(mon["Target"] > 0).round(1)
        mon["Basis"] = has_mt.map({True: "মাসিক Target", False: "দৈনিক Target-এর যোগফল"})

        sec_title("মাসিক Target বনাম Output (সব মাস)")
        c1, c2 = st.columns(2)
        with c1:
            f1 = go.Figure()
            f1.add_bar(x=mon["Month"], y=mon["Target"], name="Target", marker_color="#6B7FA3")
            f1.add_bar(x=mon["Month"], y=mon["Output"], name="Output", marker_color="#3B82F6")
            f1.update_layout(barmode="group", title="Target vs Output")
            show(f1)
        with c2:
            f2 = px.bar(mon, x="Month", y="Achievement %", text="Achievement %", title="মাসিক Achievement %",
                        color="Achievement %", color_continuous_scale=["#FB7185", "#FFB020", "#34D399"],
                        range_color=[0, 100])
            f2.update_coloraxes(showscale=False)
            show(f2)
        mshow = mon[["Month", "Total Output", "Monthly Target", "Daily Target Sum", "Achievement %", "Basis"]].copy()
        for c_ in ("Total Output", "Monthly Target", "Daily Target Sum"):
            mshow[c_] = mshow[c_].map(fmt_full)
        cdf(mshow, column_config=ACH_CFG, hide_index=True)
        st.caption("মাসিক Target সেট থাকলে Achievement = মাসের মোট Output ÷ মাসিক Target (শুক্রবারের Output-সহ)। "
                   "সেট না থাকলে শুধু Target দেওয়া দিনগুলোর Output ÷ দৈনিক Target-এর যোগফল।")

        # ---------- দৈনিক বিশ্লেষণ ----------
        sec_title("দৈনিক বিশ্লেষণ")
        c1, c2 = st.columns([3, 2])
        with c1:
            f3 = go.Figure()
            f3.add_bar(x=dd["Date"], y=dd["Daily Output"], name="Output", marker_color="#3B82F6")
            f3.add_scatter(x=dd["Date"], y=dd["Daily Target"], name="Target", mode="lines",
                           connectgaps=True, line=dict(color="#FFB020", width=3, dash="dash"))
            f3.update_layout(title=f"{tm} · দৈনিক Output বনাম Target")
            show(f3, 400)
        with c2:
            a_g = ach_m or 0
            f4 = go.Figure(go.Indicator(
                mode="gauge+number", value=a_g, number={"suffix": "%", "valueformat": ".1f"},
                title={"text": f"{tm} Achievement"},
                gauge={"axis": {"range": [0, max(100, a_g)]}, "bar": {"color": "#3B82F6"},
                       "steps": [{"range": [0, 50], "color": "rgba(251,113,133,.28)"},
                                 {"range": [50, 80], "color": "rgba(255,176,32,.28)"},
                                 {"range": [80, 100], "color": "rgba(52,211,153,.28)"}]}))
            show(f4, 400)
        d2 = dd[["Date", "Daily Target", "Daily Output"]].copy()
        d2["Achievement %"] = (d2["Daily Output"] / d2["Daily Target"] * 100).round(1)
        cdf(d2, column_config=ACH_CFG, hide_index=True)

        sec_title("দিনের Batch ও সময় (Start / End)")
        tdays = sorted(dd["Date"].dt.date.unique())
        if tdays:
            tsel = st.selectbox("দিন বেছে নিন", tdays, index=len(tdays) - 1, key="tday", format_func=fmt_day)
            rd = prod[prod["Date"].dt.date == tsel]
            if rd.empty:
                st.info("এই দিনের Production লগ পাওয়া যায়নি।")
            else:
                cdf(run_table(rd), hide_index=True)
        else:
            st.info("এই মাসে এখনো কোনো Output ডেটা নেই।")

# =====================================================================
# Batch Release
# =====================================================================
with t3:
    if bat.empty:
        st.info("Batch Release List ফাইলটা আপলোড করুন।")
    else:
        c1, c2 = st.columns(2)
        bm = sorted(bat["Month"].unique())
        bpick = c1.selectbox("মাস", ["সব মাস"] + bm, index=len(bm), key="bm")
        bview = bat if bpick == "সব মাস" else bat[bat["Month"] == bpick]
        bday = c2.selectbox("দিন", ["সব দিন"] + sorted(bview["Date"].dt.date.unique()), key="bday", format_func=fmt_day)
        btype = st.multiselect("Type (ফাঁকা থাকলে সব)", sorted(bview["Type"].unique()), key="btype")
        bq = st.text_input("🔍 Product বা Batch No খুঁজুন (কবে Release হয়েছে দেখুন)", key="bq")

        r = bview
        if bday != "সব দিন":
            r = r[r["Date"].dt.date == bday]
        if btype:
            r = r[r["Type"].isin(btype)]
        if bq.strip():
            ql = bq.strip().lower()
            r = r[r["Product Name"].str.lower().str.contains(ql, regex=False)
                  | r["Batch No"].str.lower().str.contains(ql, regex=False)]

        k = st.columns(3)
        kpi(k[0], "মোট Batch Release", f"{len(r):,}", "ফিল্টার অনুযায়ী", 0)
        kpi(k[1], "আলাদা Product", f"{r['Product Name'].nunique()}", "", 1)
        kpi(k[2], "দিন", f"{r['Date'].nunique()}", f"গড় {len(r) / max(r['Date'].nunique(), 1):.1f} batch/দিন", 2)

        if r.empty:
            st.warning("কিছু পাওয়া যায়নি।")
        else:
            sec_title("Batch Release তালিকা (কবে Release হয়েছে)")
            tb = r.sort_values(["Date", "Product Name"], ascending=[False, True])[["Date", "Product Name", "Batch No", "Type"]].copy()
            tb["Date"] = tb["Date"].dt.strftime("%d %b %Y")
            cdf(tb, hide_index=True)
            st.download_button("⬇️ এই তালিকা CSV", csv_bytes(tb), "batch_release.csv", mime="text/csv")

            sec_title("Product অনুযায়ী: কয়টা Batch, কোন কোন Batch No")
            pr = r.groupby("Product Name").agg(
                Batches=("Batch No", nb), First=("Date", "min"), Last=("Date", "max"),
                Batch_Numbers=("Batch No", join_set)).reset_index().rename(columns={"Batch_Numbers": "Batch Numbers"})
            pr["First"] = pr["First"].dt.strftime("%d %b %Y")
            pr["Last"] = pr["Last"].dt.strftime("%d %b %Y")
            cdf(pr.sort_values("Batches", ascending=False), hide_index=True)

            if bq.strip():
                prodrun = prod[prod["Batch No"].isin(set(r["Batch No"]))]
                if not prodrun.empty:
                    sec_title("এই Batch-এর Production (Start / End সময়সহ)")
                    cdf(run_table(prodrun), hide_index=True)

        with st.expander("📈 চার্ট দেখুন"):
            mt = bat.groupby(["Month", "Type"], as_index=False).size().rename(columns={"size": "Batches"})
            show(px.bar(mt, x="Month", y="Batches", color="Type", title="Type অনুযায়ী মাসিক Batch",
                        color_discrete_sequence=PALETTE))
            if not r.empty:
                c1, c2 = st.columns([3, 2])
                with c1:
                    dl = r.groupby("Date", as_index=False).size().rename(columns={"size": "Batches"})
                    show(px.bar(dl, x="Date", y="Batches", title="দৈনিক Batch Release",
                                color_discrete_sequence=["#22D3EE"]))
                with c2:
                    ty = r.groupby("Type", as_index=False).size().rename(columns={"size": "Batches"})
                    show(px.pie(ty, names="Type", values="Batches", hole=0.55, title="Type অনুযায়ী",
                                color_discrete_sequence=PALETTE))

# =====================================================================
# Export Order
# =====================================================================
STATUS_CSS = {
    "Completed": ("#bbf7d0", "#14532d"), "Packing running": ("#bfdbfe", "#1e3a8a"),
    "Ready for Blister": ("#fde68a", "#78350f"), "Ready for Compression": ("#fed7aa", "#7c2d12"),
    "Ready for Granulation": ("#ddd6fe", "#4c1d95"), "Ready for Coating": ("#fbcfe8", "#831843"),
    "Ready for Packing": ("#a5f3fc", "#164e63"), "Batch need to take": ("#fecaca", "#7f1d1d"),
    "Hold": ("#e5e7eb", "#374151"), "Other": ("#e0e7ff", "#3730a3"), "Unknown": ("#f3f4f6", "#6b7280"),
}
COUNTRY_COLORS = ["#1d4ed8", "#b45309", "#047857", "#be123c", "#6d28d9", "#0e7490", "#a16207", "#c2410c",
                  "#4d7c0f", "#9d174d", "#0f766e", "#7c3aed"]


def _status_css(v):
    bg, fg = STATUS_CSS.get(status_group(v if isinstance(v, str) and v != "-" else None), STATUS_CSS["Other"])
    return f"background-color:{bg};color:{fg};font-weight:600"


def _mat_css(v):
    s = str(v).lower()
    if s in ("-", "", "nan", "none"):
        return ""
    if "unavail" in s or "short" in s or "not " in s:
        return "background-color:#fecaca;color:#7f1d1d;font-weight:600"
    if "avail" in s:
        return "background-color:#bbf7d0;color:#14532d;font-weight:600"
    return "background-color:#fef3c7;color:#92400e;font-weight:600"


def _days_css(v):
    if pd.isna(v):
        return ""
    if v < 0:
        return "background-color:#fecaca;color:#b91c1c;font-weight:700"
    if v <= 7:
        return "background-color:#fde68a;color:#92400e;font-weight:700"
    return "background-color:#bbf7d0;color:#166534;font-weight:700"


def _late_css(v):
    if pd.isna(v):
        return ""
    return ("background-color:#fecaca;color:#b91c1c;font-weight:700" if v > 0
            else "background-color:#bbf7d0;color:#166534;font-weight:700")


def style_ex(df, cmap):
    """Export টেবিলের বাড়তি রং (ওভারলে): দেশ, Status, Delivery বাকি দিন, RM/PM ইত্যাদি।"""
    ov = [("map", lambda v: "font-weight:700", ["Product Name", "Required Pcs", "Required Box"]),
          ("map", _status_css, ["Status"]),
          ("map", _mat_css, ["RM Available", "PM Available", "PLSD"]),
          ("map", _days_css, ["Days to Delivery"]),
          ("map", _late_css, ["Late (দিন)"])]
    return ov


def raw_style(df):
    """মূল শিট দেখানোর রং: দেশের সারি, Total সারি, হেডার সারি, Status-এর শব্দ।"""
    def row_css(row):
        vals = [v for v in row if v != ""]
        low = " ".join(vals).lower()
        if len(vals) == 1 and not vals[0].replace(".", "").isdigit():
            css = "background-color:#fcd34d;color:#78350f;font-weight:800"
        elif "total" in low:
            css = "background-color:#86efac;color:#14532d;font-weight:800"
        elif ("product" in low or "country" in low) and ("qty" in low or "order" in low or "date" in low):
            css = "background-color:#93c5fd;color:#1e3a8a;font-weight:800"
        else:
            css = ""
        return [css] * len(row)

    def cell_css(v):
        s = str(v).lower()
        if "need to take" in s or "unavail" in s or "short" in s:
            return "background-color:#fecaca;color:#7f1d1d;font-weight:700"
        if "complete" in s:
            return "background-color:#bbf7d0;color:#14532d;font-weight:700"
        if "running" in s:
            return "background-color:#bfdbfe;color:#1e3a8a;font-weight:700"
        if "ready" in s:
            return "background-color:#fde68a;color:#78350f;font-weight:700"
        return ""
    return [("map", cell_css, list(df.columns)), ("rows", row_css, None)]


with t_ex:
    E = D.get("exp")
    RAW = D.get("expraw") or {}
    if E is None or E.empty:
        st.info("Export Order Status ফাইলটা (Current Export Order Status শিটসহ) আপলোড করুন।")
    else:
        E = E.reindex(columns=EXPORT_COLS).copy()
        for c_ in ("Order Date", "Completed Date", "Delivery Date"):
            E[c_] = pd.to_datetime(E[c_], errors="coerce")
        for c_ in ("Pcs per Box", "Required Box", "Required Pcs", "Batch Size", "Batches Req"):
            E[c_] = pd.to_numeric(E[c_], errors="coerce")
        E["Country"] = E["Country"].fillna("-").astype(str)
        E["Group"] = E["Status"].map(status_group)
        today = pd.Timestamp(date.today())
        cmap = {c_: COUNTRY_COLORS[i % len(COUNTRY_COLORS)] for i, c_ in enumerate(sorted(E["Country"].unique()))}

        def dfmt(s):
            return s.dt.strftime("%d %b %Y").fillna("-")

        def prep_ex(d, cols):
            t = d[cols].copy()
            for c in ("Order Date", "Completed Date", "Delivery Date"):
                if c in t:
                    t[c] = dfmt(t[c])
            for c in ("Required Pcs", "Required Box", "Batch Size", "Pcs per Box"):
                if c in t:
                    t[c] = t[c].map(fmt_full)
            if "Batches Req" in t:
                t["Batches Req"] = t["Batches Req"].round(2)
            for c in t.columns:
                if not pd.api.types.is_numeric_dtype(t[c]):
                    t[c] = t[c].fillna("-").astype(str)
            return t

        view_ex = st.radio("দেখুন", ["Open Orders", "Completed", "Demand (আসছে অর্ডার)", "সব শিট (মূল ফাইল)"],
                           horizontal=True, key="ex_view")

        if view_ex == "Open Orders":
            o_all = E[E["Source"] == "Open"].copy()
            o_all["Days to Delivery"] = (o_all["Delivery Date"] - today).dt.days
            o = o_all[o_all["Group"] != "Completed"]
            late = o[o["Days to Delivery"] < 0]
            k = st.columns(4)
            kpi(k[0], "Open Order লাইন", f"{len(o)}", f"{o['Country'].nunique()}টা দেশ", 0)
            kpi(k[1], "মোট Required (Pcs)", fmt_full(o["Required Pcs"].sum()), "Open অর্ডারের", 1)
            kpi(k[2], "Batch নিতে বাকি", f"{(o['Group'] == 'Batch need to take').sum()}", "Batch need to take", 2)
            kpi(k[3], "Delivery Date পার", f"{len(late)}", f"{fmt_full(late['Required Pcs'].sum())} Pcs", 3)

            sec_title("দেশ ও Status অনুযায়ী")
            c1, c2 = st.columns([3, 2])
            with c1:
                cy = o.groupby("Country", as_index=False)["Required Pcs"].sum().sort_values("Required Pcs")
                show(px.bar(cy, x="Required Pcs", y="Country", orientation="h", title="দেশ অনুযায়ী Required (Pcs)",
                            color="Country", color_discrete_map=cmap), 420)
            with c2:
                sg = o.groupby("Group", as_index=False).size().rename(columns={"size": "Lines"})
                show(px.pie(sg, names="Group", values="Lines", hole=0.55, title="Status অনুযায়ী লাইন",
                            color="Group", color_discrete_map={g: v[0] for g, v in STATUS_CSS.items()}), 420)

            sec_title("Open Order তালিকা (ফাইলের সব কলাম)")
            c1, c2 = st.columns(2)
            fc = c1.multiselect("দেশ (ফাঁকা = সব)", sorted(o_all["Country"].unique()), key="ex_c")
            fs = c2.multiselect("Status (ফাঁকা = সব)", sorted(o_all["Group"].unique()), key="ex_s")
            fq = st.text_input("🔍 Product / Order No / Batch No খুঁজুন", key="ex_q")
            show_done = st.checkbox("Completed স্ট্যাটাসের লাইনও দেখান", key="ex_done")
            r = o_all if show_done else o
            if fc:
                r = r[r["Country"].isin(fc)]
            if fs:
                r = r[r["Group"].isin(fs)]
            if fq.strip():
                ql = fq.strip().lower()
                r = r[r["Product Name"].fillna("").str.lower().str.contains(ql, regex=False)
                      | r["Order No"].fillna("").str.lower().str.contains(ql, regex=False)
                      | r["Batch Number"].fillna("").str.lower().str.contains(ql, regex=False)]
            cols = ["Country", "Product Name", "SAP Code", "Order No", "Machine", "Pack Size", "Pcs per Box",
                    "Required Box", "Required Pcs", "Batch Size", "Batches Req", "Batch Number", "Status",
                    "Order Date", "Delivery Date", "Days to Delivery", "PLSD", "RM Available", "PM Available",
                    "Completed Date", "Remarks"]
            t = prep_ex(r.sort_values("Delivery Date"), cols)
            cdf(t, extra=style_ex(t, cmap), hide_index=True)
            st.caption("Days to Delivery: আজ থেকে Delivery Date পর্যন্ত; মাইনাস/লাল মানে Delivery Date পার হয়ে গেছে। "
                       "Remarks = ফাইলের 'Status from Production', Completed Date = 'Packaging Completed Date'।")
            st.download_button("⬇️ এই তালিকা CSV", csv_bytes(t), "open_export_orders.csv", mime="text/csv")

        elif view_ex == "Completed":
            c = E[E["Source"] == "Completed"].copy()
            c["Late (দিন)"] = (c["Completed Date"] - c["Delivery Date"]).dt.days
            ok = c[c["Group"] == "Completed"]
            both = ok[ok["Completed Date"].notna() & ok["Delivery Date"].notna()]
            on_time = int((both["Completed Date"] <= both["Delivery Date"]).sum())
            k = st.columns(4)
            kpi(k[0], "Completed লাইন", f"{len(ok)}", c["Sheet"].iloc[0] if len(c) else "", 0)
            kpi(k[1], "মোট Completed (Pcs)", fmt_full(ok["Required Pcs"].sum()), "অর্ডার করা পরিমাণ", 1)
            kpi(k[2], "সময়মতো", f"{on_time}/{len(both)}", "Completed Date ≤ Delivery Date", 2)
            kpi(k[3], "দেশ", f"{ok['Country'].nunique()}", "", 3)
            if len(ok):
                cy = ok.groupby("Country", as_index=False)["Required Pcs"].sum().sort_values("Required Pcs")
                show(px.bar(cy, x="Required Pcs", y="Country", orientation="h", title="দেশ অনুযায়ী Completed (Pcs)",
                            color="Country", color_discrete_map=cmap), 420)
            sec_title("Completed তালিকা (ফাইলের সব কলাম)")
            cols = ["Country", "Product Name", "SAP Code", "Order No", "Machine", "Pack Size", "Pcs per Box",
                    "Required Box", "Required Pcs", "Batch Size", "Batches Req", "Batch Number", "Status",
                    "Order Date", "Completed Date", "Delivery Date", "Late (দিন)", "Remarks"]
            t = prep_ex(c.sort_values("Completed Date"), cols)
            cdf(t, extra=style_ex(t, cmap), hide_index=True)
            st.caption("Late (দিন): মাইনাস বা ০ = সময়মতো/আগে, প্লাস/লাল = Delivery Date-এর পরে হয়েছে। "
                       "Remarks-এ ফাইলের শেষের নোট/Available কলামগুলো আছে।")
            st.download_button("⬇️ এই তালিকা CSV", csv_bytes(t), "completed_export_orders.csv", mime="text/csv")

        elif view_ex.startswith("Demand"):
            dm = E[E["Source"] == "Demand"].copy()
            sheets = list(dict.fromkeys(dm["Sheet"]))
            if not sheets:
                st.info("Demand শিট পাওয়া যায়নি।")
            else:
                sh_pick = st.selectbox("Demand শিট", sheets, key="ex_dm")
                m_ = dm[dm["Sheet"] == sh_pick].copy()
                unit = m_["Pack Size"].iloc[0]
                col = "Required Pcs" if unit == "Pcs" else "Required Box"
                k = st.columns(3)
                kpi(k[0], "অর্ডার লাইন", f"{len(m_)}", f"{m_['Country'].nunique()}টা দেশ", 0)
                kpi(k[1], f"মোট ({unit})", fmt_full(m_[col].sum()), sh_pick, 1)
                kpi(k[2], "Batch নিতে বাকি", f"{(m_['Group'] == 'Batch need to take').sum()}", "", 2)
                cy = m_.groupby("Country", as_index=False)[col].sum().sort_values(col)
                show(px.bar(cy, x=col, y="Country", orientation="h", title=f"দেশ অনুযায়ী Demand ({unit})",
                            color="Country", color_discrete_map=cmap), 400)
                sec_title("Demand তালিকা (ফাইলের সব কলাম)")
                cols = ["Country", "Product Name", "SAP Code", "Order No", col, "Status", "Order Date",
                        "Delivery Date", "SO Aging", "Remarks"]
                t = prep_ex(m_.sort_values("Delivery Date"), cols)
                cdf(t, extra=style_ex(t, cmap), hide_index=True)

        else:
            if not RAW:
                st.info("মূল শিটের ডেটা পাওয়া যায়নি। Export ফাইলটা আবার আপলোড করুন।")
            else:
                sh = st.selectbox("শিট বেছে নিন", list(RAW), key="ex_raw")
                g = RAW[sh]
                st.caption(f"{sh}: {len(g)} সারি · {len(g.columns)} কলাম — ফাইলে যেমন আছে তেমন (A, B, C… = Excel কলাম)। "
                           "হলুদ = দেশ, সবুজ = Total, নীল = হেডার।")
                cdf(g, extra=raw_style(g), hide_index=True, height=560)
                st.download_button("⬇️ এই শিট CSV", csv_bytes(g), f"{sh}.csv", mime="text/csv")

# =====================================================================
# শুধু Admin: নতুন এন্ট্রি ও ডেটা
# =====================================================================
if is_admin:
    with tabs[5]:
        with st.form("entry", clear_on_submit=True):
            a, b, c = st.columns(3)
            stage = a.selectbox("Stage", ["Secondary", "Primary"])
            d = b.date_input("Date")
            shift = c.text_input("Shift (A/D)")
            owner = a.text_input("Shift Owner")
            line = b.text_input("Line / Machine")
            code = c.text_input("Code")
            batch = a.text_input("Batch No")
            product = b.text_input("Product Name")
            bsize = c.number_input("Batch Size", min_value=0, step=1000)
            psize = a.text_input("Pack Size")
            t_a = b.time_input("Run Start", value=time(6, 0))
            t_b = c.time_input("Run End", value=time(14, 0))
            manp = a.number_input("Manpower (Secondary-র জন্য)", min_value=0, step=1)
            qty = b.number_input("Output (No.)", min_value=0, step=100)
            std = c.number_input("Standard Capacity (প্রতি ঘণ্টা)", min_value=0, step=1000)
            down = a.text_input("Down Time")
            remarks = st.text_input("Remarks")
            if st.form_submit_button("➕ যোগ করুন"):
                mins = (t_b.hour * 60 + t_b.minute) - (t_a.hour * 60 + t_a.minute)
                if mins < 0:
                    mins += 24 * 60  # রাতের শিফট
                h_ = round(mins / 60, 2)
                actual = std * h_ * manp if (stage == "Secondary" and manp) else None
                row = {"Stage": stage, "Date": pd.Timestamp(d), "Day": pd.Timestamp(d).strftime("%a"),
                       "Shift": shift, "Shift Owner": owner, "Line": line, "Code": code,
                       "Batch No": batch, "Product Name": product, "Batch Size": bsize,
                       "Pack Size": psize, "Run Start": t_a.strftime("%H:%M"),
                       "Run End": t_b.strftime("%H:%M"), "Manpower": manp, "Output": qty,
                       "Hours": h_, "Std Capacity": std, "Actual Capacity": actual,
                       "Down Time": down, "Remarks": remarks}
                st.session_state.manual = pd.concat(
                    [st.session_state.manual, pd.DataFrame([row])], ignore_index=True) \
                    if len(st.session_state.manual) else pd.DataFrame([row]).reindex(columns=PROD_COLS)
                save_manual_now()
                st.rerun()

        sec_title("হাতে দেওয়া এন্ট্রি")
        man = st.session_state.manual
        cdf(man, hide_index=True)
        mb = io.BytesIO()
        mx = man.copy()
        mx["Date"] = pd.to_datetime(mx["Date"], errors="coerce").dt.date
        mx.to_excel(mb, sheet_name="Production", index=False)
        c1, c2, c3 = st.columns(3)
        if c1.button("↩️ শেষ এন্ট্রি মুছুন") and len(man):
            st.session_state.manual = man.iloc[:-1].reset_index(drop=True)
            save_manual_now()
            st.rerun()
        if gh_ready():
            if c2.button("💾 এন্ট্রিগুলো স্থায়ীভাবে সেভ"):
                try:
                    gh_put(f"data/{MANUAL_FILE.name}", mb.getvalue(), "Update manual entries")
                    st.success("সেভ হয়েছে। ১-২ মিনিটে সাইট নিজে আপডেট হবে।")
                except Exception as e:  # noqa: BLE001
                    st.error(str(e))
        else:
            c2.download_button("⬇️ manual_entries.xlsx", mb.getvalue(), MANUAL_FILE.name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.caption("স্থায়ী করতে এই ফাইলটা GitHub-এর data ফোল্ডারে আপলোড করুন।")

    with tabs[6]:
        sec_title("লোড হওয়া ফাইল")
        st.write(", ".join(n for n, _ in files) or "কোনো ফাইল নেই")
        k = st.columns(4)
        kpi(k[0], "Production সারি", f"{len(prod):,}", "Primary + Secondary", 0)
        kpi(k[1], "Target দিন", f"{len(tgt):,}", "", 1)
        kpi(k[2], "Batch সারি", f"{len(bat):,}", "", 2)
        kpi(k[3], "Export Order সারি", f"{len(D.get('exp', [])):,}", "", 3)
        st.caption("মাসিক Target: " + (", ".join(f"{m} = {fmt_full(v)}" for m, v in sorted(MT.items())) or "কিছু সেট করা নেই"))
        st.download_button("⬇️ monthly_targets.json", json.dumps(st.session_state.mt_raw, ensure_ascii=False, indent=1).encode("utf-8"),
                           MT_FILE.name, mime="application/json")
        with st.expander("Production ডেটা"):
            cdf(prod[PROD_COLS], hide_index=True)
        with st.expander("Daily Target ডেটা"):
            cdf(tgt[TARGET_COLS], hide_index=True)
        with st.expander("Batch Release ডেটা"):
            cdf(bat[BATCH_COLS], hide_index=True)
        with st.expander("Export Order ডেটা"):
            cdf(D.get("exp", pd.DataFrame(columns=EXPORT_COLS)), hide_index=True)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            for name, frame, cols in (("Production", prod, PROD_COLS), ("Daily Target", tgt, TARGET_COLS),
                                      ("Batch Release", bat, BATCH_COLS)):
                f = frame[cols].copy()
                f["Date"] = pd.to_datetime(f["Date"], errors="coerce").dt.date
                f.to_excel(xw, sheet_name=name, index=False)
            ex_ = D.get("exp", pd.DataFrame(columns=EXPORT_COLS)).reindex(columns=EXPORT_COLS)
            ex_.to_excel(xw, sheet_name="Export Orders", index=False)
        st.download_button("⬇️ সব ডেটা Excel (ব্যাকআপ)", buf.getvalue(), "all_data.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
