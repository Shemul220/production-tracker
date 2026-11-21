import io
import re
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="BSU Production Tracker", layout="wide")

# রিপোজিটরিতে এই ফোল্ডারে বা data.xlsx নামে ফাইল রাখলে সাইট নিজে থেকেই পড়ে নেবে।
DATA_DIR = Path("data")
DEFAULT_FILE = Path("data.xlsx")

PROD_COLS = ["Stage", "Date", "Day", "Shift", "Shift Owner", "Line", "Code", "Batch No",
             "Product Name", "Batch Size", "Pack Size", "Run Start", "Run End", "Manpower",
             "Output", "Hours", "Std Capacity", "Actual Capacity", "Down Time", "Remarks"]
TARGET_COLS = ["Date", "Daily Target", "Daily Output"]
BATCH_COLS = ["Date", "Product Name", "Batch No", "Type"]

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


@st.cache_data(show_spinner="ফাইল পড়া হচ্ছে...")
def load_files(files):
    prod, tgt, bat, errors = [], [], [], []
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
    return {"prod": P, "tgt": T, "bat": B}, errors


# ---------- হিসাব ----------
def prep_prod(df):
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()]
    for c in ("Output", "Hours", "Std Capacity", "Actual Capacity", "Manpower", "Batch Size"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Output"] = df["Output"].fillna(0)
    df["Hours"] = df["Hours"].fillna(0)
    for c in ("Stage", "Line", "Shift", "Product Name"):
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
    return out.drop(columns=["oc", "cap"]).round({"Hours": 1})


# ---------- সাইডবার: ডেটা লোড ----------
st.sidebar.header("ডেটা ফাইল")
st.sidebar.caption("Primary/Secondary লগবুক, Daily output, Batch release: সব ফাইল একসাথে দেওয়া যায়।")
uploads = st.sidebar.file_uploader("Excel ফাইল আপলোড", type=["xlsx"], accept_multiple_files=True)

repo_paths = sorted(p for p in DATA_DIR.glob("*.xlsx") if not p.name.startswith("~$")) if DATA_DIR.exists() else []
if DEFAULT_FILE.exists():
    repo_paths.append(DEFAULT_FILE)
files = tuple([(p.name, p.read_bytes()) for p in repo_paths] + [(u.name, u.getvalue()) for u in uploads])

key = tuple((n, len(d)) for n, d in files)
if st.session_state.get("files_key") != key:
    st.session_state.data, errs = load_files(files)
    st.session_state.errs = errs
    st.session_state.files_key = key
for e in st.session_state.get("errs", []):
    st.sidebar.warning(e)

D = st.session_state.data
prod, tgt, bat = prep_prod(D["prod"]), prep_tgt(D["tgt"]), prep_bat(D["bat"])

st.title("BSU Production Tracker")
t1, t2, t3, t4, t5 = st.tabs(["Production", "Target vs Output", "Batch Release", "নতুন এন্ট্রি", "সব ডেটা"])

# ---------- Production ----------
with t1:
    if prod.empty:
        st.info("এখনো কোনো ডেটা নেই। বাঁ পাশ থেকে Primary/Secondary Excel ফাইল আপলোড করুন।")
    else:
        sel = st.radio("ধাপ", ["সব", "Primary", "Secondary"], horizontal=True)
        base = prod if sel == "সব" else prod[prod["Stage"] == sel]
        months = sorted(base["Month"].unique())
        pick = st.selectbox("মাস", ["সব মাস"] + months, index=len(months))
        view = base if pick == "সব মাস" else base[base["Month"] == pick]

        ach = achievement(view)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("মোট Output", f"{view['Output'].sum():,.0f}")
        c2.metric("মোট চলার সময় (ঘণ্টা)", f"{view['Hours'].sum():,.1f}")
        c3.metric("Achievement", f"{ach:.1f}%" if ach is not None else "-")
        c4.metric("এন্ট্রি / দিন", f"{len(view)} / {view['Date'].nunique()}")

        left, right = st.columns(2)
        left.subheader("মাসিক Output")
        left.bar_chart(base.groupby(["Month", "Stage"])["Output"].sum().unstack(fill_value=0), stack=False)
        right.subheader("মাসিক চলার সময় (ঘণ্টা)")
        right.bar_chart(base.groupby(["Month", "Stage"])["Hours"].sum().unstack(fill_value=0), stack=False)

        st.subheader("সাপ্তাহিক Achievement")
        st.dataframe(summarize(view, ["Stage", "Week"] if sel == "সব" else "Week"), use_container_width=True)

        a, b = st.columns(2)
        a.subheader("Line / Machine অনুযায়ী")
        a.dataframe(summarize(view, ["Stage", "Line"] if sel == "সব" else "Line"), use_container_width=True)
        b.subheader("Shift অনুযায়ী")
        b.dataframe(summarize(view, ["Stage", "Shift"] if sel == "সব" else "Shift"), use_container_width=True)

        st.subheader("Product অনুযায়ী")
        st.dataframe(summarize(view, "Product Name").sort_values("Output", ascending=False),
                     use_container_width=True)

# ---------- Target vs Output ----------
with t2:
    if tgt.empty:
        st.info("Daily output reporting ফাইলটা (Daily Target Output Achievement শিটসহ) আপলোড করুন।")
    else:
        valid = tgt[tgt["Daily Target"].notna() & tgt["Daily Output"].notna()]
        mon = tgt.groupby("Month").agg(Total_Output=("Daily Output", "sum"))
        vm = valid.groupby("Month").agg(Output_on_target_days=("Daily Output", "sum"),
                                        Target=("Daily Target", "sum"), Days=("Date", "count"))
        mon = mon.join(vm, how="left")
        mon["Achievement %"] = (mon["Output_on_target_days"] / mon["Target"] * 100).round(1)

        c1, c2, c3 = st.columns(3)
        c1.metric("মোট Output", f"{tgt['Daily Output'].sum():,.0f}")
        tt, oo = valid["Daily Target"].sum(), valid["Daily Output"].sum()
        c2.metric("মোট Target", f"{tt:,.0f}")
        c3.metric("Achievement", f"{oo / tt * 100:.1f}%" if tt else "-")

        st.subheader("মাসিক Target বনাম Output")
        st.bar_chart(mon[["Target", "Output_on_target_days"]].rename(
            columns={"Output_on_target_days": "Output"}), stack=False)
        st.dataframe(mon, use_container_width=True)
        st.caption("Achievement শুধু Target দেওয়া দিনগুলোর হিসাবে (যেমন শুক্রবার বাদ)।")

        tm = st.selectbox("দৈনিক দেখুন, মাস", sorted(tgt["Month"].unique()),
                          index=len(tgt["Month"].unique()) - 1, key="tm")
        dd = tgt[tgt["Month"] == tm].set_index("Date")[["Daily Output", "Daily Target"]]
        st.line_chart(dd)
        dd2 = dd.copy()
        dd2["Achievement %"] = (dd2["Daily Output"] / dd2["Daily Target"] * 100).round(1)
        st.dataframe(dd2, use_container_width=True)

# ---------- Batch Release ----------
with t3:
    if bat.empty:
        st.info("Batch Release List ফাইলটা আপলোড করুন।")
    else:
        bm = sorted(bat["Month"].unique())
        bpick = st.selectbox("মাস", ["সব মাস"] + bm, index=len(bm), key="bm")
        bview = bat if bpick == "সব মাস" else bat[bat["Month"] == bpick]
        c1, c2, c3 = st.columns(3)
        c1.metric("মোট Batch Release", f"{len(bview):,}")
        c2.metric("আলাদা Product", bview["Product Name"].nunique())
        c3.metric("কত দিনে", bview["Date"].nunique())

        st.subheader("মাসিক Batch Release (Type অনুযায়ী)")
        st.bar_chart(bat.groupby(["Month", "Type"]).size().unstack(fill_value=0))

        if bpick != "সব মাস":
            st.subheader("দৈনিক Batch Release")
            st.bar_chart(bview.groupby("Date").size().rename("Batches"))

        a, b = st.columns(2)
        a.subheader("Type অনুযায়ী")
        a.dataframe(bview.groupby("Type").size().rename("Batches").sort_values(ascending=False),
                    use_container_width=True)
        b.subheader("সবচেয়ে বেশি Release হওয়া Product")
        b.dataframe(bview.groupby("Product Name").size().rename("Batches")
                    .sort_values(ascending=False).head(20), use_container_width=True)

# ---------- নতুন এন্ট্রি (Production) ----------
with t4:
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
        if st.form_submit_button("যোগ করুন"):
            mins = (t_b.hour * 60 + t_b.minute) - (t_a.hour * 60 + t_a.minute)
            if mins < 0:
                mins += 24 * 60  # রাতের শিফট
            hrs = round(mins / 60, 2)
            actual = std * hrs * manp if (stage == "Secondary" and manp) else None
            row = {"Stage": stage, "Date": pd.Timestamp(d), "Day": pd.Timestamp(d).strftime("%a"),
                   "Shift": shift, "Shift Owner": owner, "Line": line, "Code": code,
                   "Batch No": batch, "Product Name": product, "Batch Size": bsize,
                   "Pack Size": psize, "Run Start": t_a.strftime("%H:%M"),
                   "Run End": t_b.strftime("%H:%M"), "Manpower": manp, "Output": qty,
                   "Hours": hrs, "Std Capacity": std, "Actual Capacity": actual,
                   "Down Time": down, "Remarks": remarks}
            st.session_state.data["prod"] = pd.concat(
                [st.session_state.data["prod"], pd.DataFrame([row])], ignore_index=True)
            st.rerun()

# ---------- সব ডেটা ----------
with t5:
    with st.expander("Production ডেটা", expanded=False):
        e_prod = st.data_editor(prod[PROD_COLS], num_rows="dynamic", use_container_width=True, key="e1")
    with st.expander("Daily Target ডেটা", expanded=False):
        e_tgt = st.data_editor(tgt[TARGET_COLS], num_rows="dynamic", use_container_width=True, key="e2")
    with st.expander("Batch Release ডেটা", expanded=False):
        e_bat = st.data_editor(bat[BATCH_COLS], num_rows="dynamic", use_container_width=True, key="e3")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, frame in (("Production", e_prod), ("Daily Target", e_tgt), ("Batch Release", e_bat)):
            f = frame.copy()
            f["Date"] = pd.to_datetime(f["Date"], errors="coerce").dt.date
            f.to_excel(xw, sheet_name=name, index=False)
    st.download_button("সব ডেটা Excel হিসেবে ডাউনলোড", buf.getvalue(), "data.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("সাইট রিস্টার্ট হলে সাইটে দেওয়া এন্ট্রি মুছে যেতে পারে। ডাউনলোড করে GitHub-এর data.xlsx বদলে দিন।")
