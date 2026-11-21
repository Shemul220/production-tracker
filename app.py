import io
from datetime import time
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Secondary Packaging Tracker", layout="wide")

DEFAULT_FILE = Path("data.xlsx")

# মূল লগবুক ফাইলের কলামের অবস্থান (০ থেকে গোনা) -> সাইটে ব্যবহৃত নাম।
# নতুন কলাম লাগলে এখানে যোগ করুন।
LOGBOOK_MAP = {
    0: "Date", 1: "Day", 2: "Shift", 3: "Shift Owner", 4: "Line", 5: "Code",
    6: "Batch No", 7: "Product Name", 8: "Batch Size", 9: "Pack Size",
    13: "Run Start", 14: "Run End", 17: "Manpower", 18: "Output",
    25: "Hours", 27: "Std Capacity", 30: "Down Time", 31: "Remarks",
}
COLUMNS = list(LOGBOOK_MAP.values())


def to_hhmm(x):
    return x.strftime("%H:%M") if hasattr(x, "strftime") else x


def read_any(src):
    """মূল লগবুক ফাইল বা এই সাইট থেকে ডাউনলোড করা ফাইল, দুটোই পড়তে পারে।"""
    if hasattr(src, "seek"):
        src.seek(0)
    if str(getattr(src, "name", src)).lower().endswith(".csv"):
        return pd.read_csv(src)[COLUMNS]
    peek = pd.read_excel(src, sheet_name=0, header=0, nrows=3)
    if hasattr(src, "seek"):
        src.seek(0)
    if "Product Name" in peek.columns:  # সাইট থেকে ডাউনলোড করা ফরম্যাট
        return pd.read_excel(src, sheet_name=0)[COLUMNS]
    raw = pd.read_excel(src, sheet_name=0, header=1, nrows=3000)  # মূল লগবুক
    df = raw.iloc[:, list(LOGBOOK_MAP)].copy()
    df.columns = COLUMNS
    df = df[df["Date"].notna()]
    df = df[df["Product Name"].notna() & ~df["Product Name"].astype(str).str.startswith("#")]
    for c in ("Run Start", "Run End"):
        df[c] = df[c].apply(to_hhmm)
    return df


@st.cache_data(show_spinner=False)
def load_files(files):
    frames = [read_any(io.BytesIO(data)) for _, data in files]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=COLUMNS)


def prepare(df):
    df = df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()]
    for c in ("Output", "Hours", "Std Capacity", "Manpower", "Batch Size"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Output"] = df["Output"].fillna(0)
    df["Hours"] = df["Hours"].fillna(0)
    df["Actual Capacity"] = df["Std Capacity"] * df["Hours"] * df["Manpower"]
    df["Output/Hr"] = (df["Output"] / df["Hours"]).where(df["Hours"] > 0)
    df["Month"] = df["Date"].dt.strftime("%Y-%m")
    df["Week"] = df["Date"].dt.strftime("%G-W%V")
    return df


def summarize(g):
    out = g.agg(Output=("Output", "sum"), Hours=("Hours", "sum"),
                Actual_Cap=("Actual Capacity", "sum"))
    out["Achievement %"] = (out["Output"] / out["Actual_Cap"] * 100).where(out["Actual_Cap"] > 0).round(1)
    out["Output/Hr"] = (out["Output"] / out["Hours"]).where(out["Hours"] > 0).round(0)
    return out.drop(columns="Actual_Cap").round({"Hours": 1})


# ---------- ডেটা লোড ----------
st.sidebar.header("ডেটা ফাইল")
uploads = st.sidebar.file_uploader("মাসিক Excel ফাইল (একাধিক দেওয়া যায়)",
                                   type=["xlsx", "csv"], accept_multiple_files=True)
files = tuple((u.name, u.getvalue()) for u in uploads)
if not files and DEFAULT_FILE.exists():
    files = ((DEFAULT_FILE.name, DEFAULT_FILE.read_bytes()),)

if st.session_state.get("files_key") != tuple(n for n, _ in files):
    st.session_state.raw = load_files(files)
    st.session_state.files_key = tuple(n for n, _ in files)

df = prepare(st.session_state.raw)

st.title("Secondary Packaging Tracker")
tab1, tab2, tab3 = st.tabs(["Dashboard", "নতুন এন্ট্রি", "সব ডেটা"])

# ---------- Dashboard ----------
with tab1:
    if df.empty:
        st.info("এখনো কোনো ডেটা নেই। বাঁ পাশ থেকে Excel আপলোড করুন বা 'নতুন এন্ট্রি' ট্যাবে ডেটা দিন।")
    else:
        months = sorted(df["Month"].unique())
        pick = st.selectbox("মাস", ["সব মাস"] + months, index=len(months))
        view = df if pick == "সব মাস" else df[df["Month"] == pick]

        cap = view["Actual Capacity"].sum()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("মোট Output", f"{view['Output'].sum():,.0f}")
        c2.metric("মোট চলার সময় (ঘণ্টা)", f"{view['Hours'].sum():,.1f}")
        c3.metric("Achievement", f"{view['Output'].sum() / cap * 100:.1f}%" if cap > 0 else "-")
        c4.metric("এন্ট্রি / দিন", f"{len(view)} / {view['Date'].nunique()}")

        monthly = summarize(df.groupby("Month"))
        left, right = st.columns(2)
        left.subheader("মাসিক Output")
        left.bar_chart(monthly["Output"])
        right.subheader("মাসিক চলার সময় (ঘণ্টা)")
        right.bar_chart(monthly["Hours"])

        st.subheader("সাপ্তাহিক Achievement")
        st.dataframe(summarize(view.groupby("Week")), use_container_width=True)

        a, b = st.columns(2)
        a.subheader("Line অনুযায়ী")
        a.dataframe(summarize(view.groupby("Line")), use_container_width=True)
        b.subheader("Shift অনুযায়ী")
        b.dataframe(summarize(view.groupby("Shift")), use_container_width=True)

        st.subheader("Product অনুযায়ী")
        st.dataframe(summarize(view.groupby("Product Name")).sort_values("Output", ascending=False),
                     use_container_width=True)

# ---------- নতুন এন্ট্রি ----------
with tab2:
    with st.form("entry", clear_on_submit=True):
        a, b, c = st.columns(3)
        d = a.date_input("Date")
        shift = b.text_input("Shift (A/D)")
        owner = c.text_input("Shift Owner")
        line = a.text_input("Line")
        code = b.text_input("Code")
        batch = c.text_input("Batch No")
        product = a.text_input("Product Name")
        bsize = b.number_input("Batch Size", min_value=0, step=1000)
        psize = c.text_input("Pack Size")
        t1 = a.time_input("Run Start", value=time(6, 0))
        t2 = b.time_input("Run End", value=time(14, 0))
        manp = c.number_input("Manpower", min_value=0, step=1)
        qty = a.number_input("Output (No.)", min_value=0, step=100)
        std = b.number_input("Standard Capacity (প্রতি ঘণ্টা)", min_value=0, step=1000)
        down = c.text_input("Down Time")
        remarks = st.text_input("Remarks")
        if st.form_submit_button("যোগ করুন"):
            mins = (t2.hour * 60 + t2.minute) - (t1.hour * 60 + t1.minute)
            if mins < 0:
                mins += 24 * 60  # রাতের শিফট
            row = {"Date": pd.Timestamp(d), "Day": pd.Timestamp(d).strftime("%a"), "Shift": shift,
                   "Shift Owner": owner, "Line": line, "Code": code, "Batch No": batch,
                   "Product Name": product, "Batch Size": bsize, "Pack Size": psize,
                   "Run Start": t1.strftime("%H:%M"), "Run End": t2.strftime("%H:%M"),
                   "Manpower": manp, "Output": qty, "Hours": round(mins / 60, 2),
                   "Std Capacity": std, "Down Time": down, "Remarks": remarks}
            st.session_state.raw = pd.concat(
                [st.session_state.raw, pd.DataFrame([row])], ignore_index=True)
            st.rerun()

# ---------- সব ডেটা ----------
with tab3:
    edited = st.data_editor(df[COLUMNS], num_rows="dynamic", use_container_width=True)
    out = edited.copy()
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce").dt.date
    buf = io.BytesIO()
    out.to_excel(buf, index=False)
    st.download_button("আপডেট করা Excel ডাউনলোড", buf.getvalue(), "data.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("সাইটে দেওয়া নতুন এন্ট্রি রিস্টার্টে মুছে যেতে পারে। ডাউনলোড করে GitHub-এর data.xlsx বদলে দিন।")
