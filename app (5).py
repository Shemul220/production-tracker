import base64
import hmac
import io
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
    out["Output"] = out["Output"].round(0).astype("int64")
    return out.drop(columns=["oc", "cap"]).round({"Hours": 1})



# =====================================================================
# স্টাইল
# =====================================================================
PALETTE = ["#1F77D0", "#FF8C42", "#19B5A5", "#8E5BD9", "#EF476F", "#FFC43D", "#2DA44E", "#00A6D6"]
STAGE_COL = {"Primary": "#1F77D0", "Secondary": "#FF8C42"}
GRADS = ["linear-gradient(135deg,#1565C0,#2AA7E8)", "linear-gradient(135deg,#F57C00,#FFC046)",
         "linear-gradient(135deg,#0E9F8E,#43E0A8)", "linear-gradient(135deg,#7B3FE4,#C08BFF)"]
ACH_CFG = {"Achievement %": st.column_config.ProgressColumn(
    "Achievement %", format="%.1f%%", min_value=0, max_value=100)}

st.markdown("""
<style>
.block-container {padding-top: 1rem; max-width: 1250px;}
footer {visibility: hidden;}
/* ব্যানার */
.hero {position: relative; overflow: hidden; color:#fff; padding: 24px 28px; border-radius: 20px;
       margin-bottom: 16px; background: linear-gradient(120deg,#0B3C6F 0%,#1F77D0 50%,#19B5A5 100%);
       box-shadow: 0 8px 24px rgba(31,119,208,.35);}
.hero:before {content:""; position:absolute; right:-40px; top:-60px; width:230px; height:230px;
       border-radius:50%; background: rgba(255,255,255,.12);}
.hero:after {content:""; position:absolute; right:110px; bottom:-90px; width:200px; height:200px;
       border-radius:50%; background: rgba(255,140,66,.30);}
.hero-t {position:relative; font-size: 1.85rem; font-weight: 800; letter-spacing:.3px;}
.hero-s {position:relative; opacity:.93; font-size:.96rem; margin-top:4px;}
.chips {position:relative; margin-top:12px;}
.chip {display:inline-block; background: rgba(255,255,255,.20); border:1px solid rgba(255,255,255,.35);
       padding: 4px 12px; border-radius: 999px; font-size:.78rem; margin: 0 6px 6px 0; font-weight:600;}
/* KPI কার্ড */
.kpi {position:relative; overflow:hidden; color:#fff; padding: 16px 18px; border-radius: 18px;
      box-shadow: 0 6px 16px rgba(20,40,80,.22); margin-bottom: 12px; min-height: 112px;}
.kpi:after {content:""; position:absolute; right:-22px; top:-22px; width:90px; height:90px;
      border-radius:50%; background: rgba(255,255,255,.18);}
.kpi-l {font-size:.76rem; opacity:.96; font-weight:700; text-transform:uppercase; letter-spacing:.7px;}
.kpi-v {font-size: 1.75rem; font-weight: 800; margin-top: 4px; line-height: 1.15;}
.kpi-s {font-size:.8rem; opacity:.93; margin-top:3px;}
/* সেকশন হেডার */
.sec {font-size: 1.08rem; font-weight: 700; margin: 22px 0 10px 0; padding: 8px 14px;
      border-left: 6px solid #FF8C42; border-radius: 8px;
      background: linear-gradient(90deg, rgba(31,119,208,.20), rgba(31,119,208,0));}
/* চার্ট কার্ড */
div[data-testid="stPlotlyChart"] {background:#FFFFFF; border-radius:18px; padding:8px 6px;
      box-shadow: 0 4px 18px rgba(20,40,80,.14); border-top: 4px solid #1F77D0; margin-bottom: 10px;}
div[data-testid="stDataFrame"] {border-radius: 12px; box-shadow: 0 3px 12px rgba(20,40,80,.12); overflow:hidden;}
/* ট্যাব */
div[data-baseweb="tab-list"] {gap: 8px; flex-wrap: wrap;}
div[data-baseweb="tab-highlight"], div[data-baseweb="tab-border"] {display:none;}
button[data-baseweb="tab"] {font-weight:700; padding: 9px 16px; border-radius: 12px;
      background: rgba(31,119,208,.10); border: 1px solid rgba(31,119,208,.25);}
button[data-baseweb="tab"][aria-selected="true"] {color:#fff !important;
      background: linear-gradient(120deg,#1F77D0,#19B5A5); border-color: transparent;
      box-shadow: 0 4px 12px rgba(31,119,208,.35);}
button[data-baseweb="tab"][aria-selected="true"] p {color:#fff !important;}
/* বাটন */
.stButton > button, .stDownloadButton > button, div[data-testid="stFormSubmitButton"] > button {
      border-radius: 12px; font-weight: 700; border: 0; color:#fff;
      background: linear-gradient(120deg,#FF8C42,#EF476F);}
</style>
""", unsafe_allow_html=True)


def style(fig, h=370):
    fig.update_layout(height=h, margin=dict(l=10, r=10, t=48, b=10), paper_bgcolor="#FFFFFF",
                      plot_bgcolor="#FFFFFF", font=dict(color="#1B2A41", size=12), colorway=PALETTE,
                      legend=dict(orientation="h", y=-0.2, x=0), title_font=dict(size=16, color="#0B3C6F"))
    fig.update_xaxes(gridcolor="#E6ECF5", zeroline=False, linecolor="#C9D4E5")
    fig.update_yaxes(gridcolor="#E6ECF5", zeroline=False, linecolor="#C9D4E5")
    return fig


def show(fig, h=370):
    st.plotly_chart(style(fig, h), theme=None)


def fmt(n):
    if n is None or pd.isna(n):
        return "-"
    n = float(n)
    if abs(n) >= 1e6:
        return f"{n / 1e6:.2f}M"
    if abs(n) >= 1e4:
        return f"{n / 1e3:.1f}K"
    return f"{n:,.0f}"


def kpi(col, label, value, sub="", g=0):
    col.markdown(f'<div class="kpi" style="background:{GRADS[g % 4]}"><div class="kpi-l">{label}</div>'
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
        uploads = st.file_uploader("Excel ফাইল (একাধিক দেওয়া যায়)", type=["xlsx"],
                                   accept_multiple_files=True)
        if uploads:
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

repo_paths = [p for p in sorted(DATA_DIR.glob("*.xlsx"))
              if not p.name.startswith("~$") and p.name != MANUAL_FILE.name] if DATA_DIR.exists() else []
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

def _span(df, label):
    if df.empty:
        return None
    return f"{label}: {df['Date'].min():%d %b %Y} – {df['Date'].max():%d %b %Y}"


hero("Primary &amp; Secondary Packaging · Daily Target vs Output · Batch Release",
     [c for c in (_span(prod, "📊 Production"), _span(tgt, "🎯 Target"), _span(bat, "📦 Batch")) if c])
labels = ["📊 Production", "🎯 Target vs Output", "📦 Batch Release"]
if is_admin:
    labels += ["➕ নতুন এন্ট্রি", "🗂️ ডেটা"]
tabs = st.tabs(labels)
t1, t2, t3 = tabs[0], tabs[1], tabs[2]

# =====================================================================
# Production
# =====================================================================
with t1:
    if prod.empty:
        st.info("এখনো কোনো ডেটা নেই। Admin হিসেবে Primary/Secondary Excel ফাইল আপলোড করুন।")
    else:
        sel = st.radio("ধাপ", ["সব", "Primary", "Secondary"], horizontal=True)
        base = prod if sel == "সব" else prod[prod["Stage"] == sel]
        months = sorted(base["Month"].unique())
        pick = st.selectbox("মাস", ["সব মাস"] + months, index=len(months))
        view = base if pick == "সব মাস" else base[base["Month"] == pick]

        ach = achievement(view)
        hrs = view["Hours"].sum()
        k = st.columns(4)
        kpi(k[0], "মোট Output", fmt(view["Output"].sum()), f"{len(view)} এন্ট্রি", 0)
        kpi(k[1], "চলার সময়", f"{hrs:,.0f} ঘণ্টা", f"{view['Date'].nunique()} দিন", 1)
        kpi(k[2], "Achievement", f"{ach:.1f}%" if ach is not None else "-", f"{sel} · {pick}", 2)
        kpi(k[3], "Output / ঘণ্টা", fmt(view["Output"].sum() / hrs) if hrs > 0 else "-", "গড়", 3)

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
        st.dataframe(wk, column_config=ACH_CFG, hide_index=True)

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
        st.dataframe(ln.sort_values("Output", ascending=False), column_config=ACH_CFG, hide_index=True)

        sec_title("Product অনুযায়ী")
        pr = summarize(view, "Product Name").sort_values("Output", ascending=False)
        top = pr.head(10).reset_index().sort_values("Output")
        fig = px.bar(top, x="Output", y="Product Name", orientation="h", title="Top 10 Product",
                     color="Output", color_continuous_scale=["#9CCBF5", "#1F77D0", "#0B3C6F"])
        fig.update_coloraxes(showscale=False)
        show(fig, 420)
        st.dataframe(pr, column_config=ACH_CFG)

# =====================================================================
# Target vs Output
# =====================================================================
with t2:
    if tgt.empty:
        st.info("Daily output reporting ফাইলটা (Daily Target Output Achievement শিটসহ) আপলোড করুন।")
    else:
        valid = tgt[tgt["Daily Target"].notna() & tgt["Daily Output"].notna()]
        tt, oo = valid["Daily Target"].sum(), valid["Daily Output"].sum()
        k = st.columns(4)
        kpi(k[0], "মোট Output", fmt(tgt["Daily Output"].sum()), "সব দিন", 0)
        kpi(k[1], "মোট Target", fmt(tt), "Target-যুক্ত দিন", 1)
        kpi(k[2], "Achievement", f"{oo / tt * 100:.1f}%" if tt else "-", "Output ÷ Target", 2)
        kpi(k[3], "দিন", f"{len(valid)}", "Target-যুক্ত দিন", 3)

        mon = tgt.groupby("Month").agg(Total_Output=("Daily Output", "sum"))
        vm = valid.groupby("Month").agg(Output=("Daily Output", "sum"), Target=("Daily Target", "sum"),
                                        Days=("Date", "count"))
        mon = mon.join(vm, how="left").reset_index()
        mon["Achievement %"] = (mon["Output"] / mon["Target"] * 100).round(1)

        sec_title("মাসিক Target বনাম Output")
        c1, c2 = st.columns(2)
        with c1:
            f1 = go.Figure()
            f1.add_bar(x=mon["Month"], y=mon["Target"], name="Target", marker_color="#B8C4D6")
            f1.add_bar(x=mon["Month"], y=mon["Output"], name="Output", marker_color="#1F77D0")
            f1.update_layout(barmode="group", title="Target vs Output")
            show(f1)
        with c2:
            f2 = px.bar(mon, x="Month", y="Achievement %", text="Achievement %", title="মাসিক Achievement %",
                        color="Achievement %", color_continuous_scale=["#EF476F", "#FFC43D", "#2DA44E"],
                        range_color=[0, 100])
            f2.update_coloraxes(showscale=False)
            show(f2)
        st.dataframe(mon, column_config=ACH_CFG, hide_index=True)
        st.caption("Achievement শুধু Target দেওয়া দিনগুলোর হিসাবে (শুক্রবার ইত্যাদি বাদ)।")

        sec_title("দৈনিক বিশ্লেষণ")
        all_m = sorted(tgt["Month"].unique())
        tm = st.selectbox("মাস", all_m, index=len(all_m) - 1, key="tm")
        dd = tgt[tgt["Month"] == tm]
        mv = valid[valid["Month"] == tm]
        a_m = mv["Daily Output"].sum() / mv["Daily Target"].sum() * 100 if mv["Daily Target"].sum() > 0 else 0
        c1, c2 = st.columns([3, 2])
        with c1:
            f3 = go.Figure()
            f3.add_bar(x=dd["Date"], y=dd["Daily Output"], name="Output", marker_color="#1F77D0")
            f3.add_scatter(x=dd["Date"], y=dd["Daily Target"], name="Target", mode="lines",
                           connectgaps=True, line=dict(color="#FF8C42", width=3, dash="dash"))
            f3.update_layout(title=f"{tm} · দৈনিক Output বনাম Target")
            show(f3, 400)
        with c2:
            f4 = go.Figure(go.Indicator(
                mode="gauge+number", value=a_m, number={"suffix": "%", "valueformat": ".1f"},
                title={"text": f"{tm} Achievement"},
                gauge={"axis": {"range": [0, 100]}, "bar": {"color": "#1F77D0"},
                       "steps": [{"range": [0, 50], "color": "#FADADD"},
                                 {"range": [50, 80], "color": "#FFF1C9"},
                                 {"range": [80, 100], "color": "#D4F3E0"}]}))
            show(f4, 400)
        d2 = dd[["Date", "Daily Target", "Daily Output"]].copy()
        d2["Achievement %"] = (d2["Daily Output"] / d2["Daily Target"] * 100).round(1)
        st.dataframe(d2, column_config=ACH_CFG, hide_index=True)

# =====================================================================
# Batch Release
# =====================================================================
with t3:
    if bat.empty:
        st.info("Batch Release List ফাইলটা আপলোড করুন।")
    else:
        bm = sorted(bat["Month"].unique())
        bpick = st.selectbox("মাস", ["সব মাস"] + bm, index=len(bm), key="bm")
        bview = bat if bpick == "সব মাস" else bat[bat["Month"] == bpick]
        k = st.columns(3)
        kpi(k[0], "মোট Batch Release", f"{len(bview):,}", bpick, 0)
        kpi(k[1], "আলাদা Product", f"{bview['Product Name'].nunique()}", "", 1)
        kpi(k[2], "দিন", f"{bview['Date'].nunique()}", f"গড় {len(bview) / max(bview['Date'].nunique(), 1):.1f} batch/দিন", 2)

        sec_title("মাসিক Batch Release")
        mt = bat.groupby(["Month", "Type"], as_index=False).size().rename(columns={"size": "Batches"})
        show(px.bar(mt, x="Month", y="Batches", color="Type", title="Type অনুযায়ী মাসিক Batch",
                    color_discrete_sequence=PALETTE))

        c1, c2 = st.columns([3, 2])
        with c1:
            if bpick != "সব মাস":
                dl = bview.groupby("Date", as_index=False).size().rename(columns={"size": "Batches"})
                show(px.bar(dl, x="Date", y="Batches", title="দৈনিক Batch Release",
                            color_discrete_sequence=["#19B5A5"]))
            else:
                tot = bat.groupby("Month", as_index=False).size().rename(columns={"size": "Batches"})
                show(px.line(tot, x="Month", y="Batches", markers=True, title="মাসিক মোট Batch",
                             color_discrete_sequence=["#8E5BD9"]))
        with c2:
            ty = bview.groupby("Type", as_index=False).size().rename(columns={"size": "Batches"})
            show(px.pie(ty, names="Type", values="Batches", hole=0.55, title="Type অনুযায়ী",
                        color_discrete_sequence=PALETTE))

        sec_title("সবচেয়ে বেশি Release হওয়া Product")
        tp = bview.groupby("Product Name", as_index=False).size().rename(columns={"size": "Batches"})
        tp = tp.sort_values("Batches", ascending=False).head(15).sort_values("Batches")
        show(px.bar(tp, x="Batches", y="Product Name", orientation="h", color="Batches",
                    color_continuous_scale=["#B7E4DD", "#19B5A5", "#0B6E63"]).update_coloraxes(showscale=False), 480)

# =====================================================================
# শুধু Admin: নতুন এন্ট্রি ও ডেটা
# =====================================================================
if is_admin:
    with tabs[3]:
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
                st.rerun()

        sec_title("হাতে দেওয়া এন্ট্রি")
        man = st.session_state.manual
        st.dataframe(man, hide_index=True)
        mb = io.BytesIO()
        mx = man.copy()
        mx["Date"] = pd.to_datetime(mx["Date"], errors="coerce").dt.date
        mx.to_excel(mb, sheet_name="Production", index=False)
        c1, c2, c3 = st.columns(3)
        if c1.button("↩️ শেষ এন্ট্রি মুছুন") and len(man):
            st.session_state.manual = man.iloc[:-1].reset_index(drop=True)
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

    with tabs[4]:
        sec_title("লোড হওয়া ফাইল")
        st.write(", ".join(n for n, _ in files) or "কোনো ফাইল নেই")
        k = st.columns(3)
        kpi(k[0], "Production সারি", f"{len(prod):,}", "Primary + Secondary", 0)
        kpi(k[1], "Target দিন", f"{len(tgt):,}", "", 1)
        kpi(k[2], "Batch সারি", f"{len(bat):,}", "", 2)
        with st.expander("Production ডেটা"):
            st.dataframe(prod[PROD_COLS], hide_index=True)
        with st.expander("Daily Target ডেটা"):
            st.dataframe(tgt[TARGET_COLS], hide_index=True)
        with st.expander("Batch Release ডেটা"):
            st.dataframe(bat[BATCH_COLS], hide_index=True)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            for name, frame, cols in (("Production", prod, PROD_COLS), ("Daily Target", tgt, TARGET_COLS),
                                      ("Batch Release", bat, BATCH_COLS)):
                f = frame[cols].copy()
                f["Date"] = pd.to_datetime(f["Date"], errors="coerce").dt.date
                f.to_excel(xw, sheet_name=name, index=False)
        st.download_button("⬇️ সব ডেটা Excel (ব্যাকআপ)", buf.getvalue(), "all_data.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
