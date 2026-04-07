import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px

os.environ["TZ"] = "Asia/Kuala_Lumpur"
st.set_page_config(layout="wide")

# =========================
# 默认工时
# =========================
DEFAULT_LEAD_TIME = {
    "Laser Cut": 4, "Laser Tube": 5, "Punching": 3,
    "Bending": 6, "Welding": 8, "Painting": 10, "Assembly": 12
}

# =========================
# 自动识别 header
# =========================
def detect_header(file):
    raw = pd.read_excel(file, header=None, engine="openpyxl")
    for i, row in raw.iterrows():
        row_str = row.astype(str)
        if row_str.str.contains("Main Part", case=False).any():
            return i
    return 5

def load_excel(files):
    dfs = []
    for f in files:
        h = detect_header(f)
        df = pd.read_excel(f, header=h, engine="openpyxl")
        df.columns = df.columns.astype(str).str.strip()
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

# =========================
# 自动列 mapping
# =========================
def find_col(df, keywords):
    for col in df.columns:
        for k in keywords:
            if k.lower() in col.lower():
                return col
    return None

def map_columns(df):
    return {
        "job": find_col(df, ["job"]),
        "part": find_col(df, ["subpart"]),
        "main": find_col(df, ["main part"]),
        "category": find_col(df, ["category"]),
        "exwork": find_col(df, ["exwork","due"]),
        "order": find_col(df, ["order date"]),
        "current": find_col(df, ["current operation"]),
        "nest": find_col(df, ["nesting"])
    }

# =========================
# Step解析（含merge）
# =========================
def extract_steps(row):
    steps = []
    for col in row.index:
        if pd.isna(row[col]):
            continue
        if "step" in str(col).lower() or "unnamed" in str(col).lower():
            steps.append(row[col])
    return steps

def get_current_step(row, colmap):
    if colmap["current"] and pd.notna(row[colmap["current"]]):
        return row[colmap["current"]]
    steps = extract_steps(row)
    return steps[0] if steps else None

def next_step(row, colmap):
    steps = extract_steps(row)
    cur = get_current_step(row, colmap)
    if cur in steps:
        i = steps.index(cur)
        if i+1 < len(steps):
            return steps[i+1]
    return None

def calc_eta(row, cal):
    total = 0
    for s in extract_steps(row):
        total += cal.get(s, DEFAULT_LEAD_TIME.get(s, 5))
    return datetime.now() + timedelta(hours=total)

def progress(row, colmap):
    steps = extract_steps(row)
    cur = get_current_step(row, colmap)
    if not steps:
        return 0
    if cur not in steps:
        return 100
    return int(100 * steps.index(cur) / len(steps))

# =========================
# Customer识别
# =========================
def get_customer(row, colmap):
    if colmap["main"] and pd.notna(row[colmap["main"]]):
        return str(row[colmap["main"]]).split("-")[0]
    return None

# =========================
# 🔥 Base Part（核心逻辑）
# =========================
def get_base_part(row, colmap):
    if colmap["part"] and pd.notna(row[colmap["part"]]):
        val = str(row[colmap["part"]])
    elif colmap["main"] and pd.notna(row[colmap["main"]]):
        val = str(row[colmap["main"]])
    else:
        return None

    parts = val.split("-")

    if parts[-1].isdigit():
        return "-".join(parts[:-1])

    return val

# =========================
# 🔥 Engineering判断（最终正确逻辑）
# =========================
def is_engineering_required(row, df_all, colmap):
    base = row["Base Part"]
    if not base:
        return False

    has_job = df_all[
        (df_all["Base Part"] == base) &
        (df_all[colmap["job"]].notna())
    ]

    return has_job.empty

# =========================
# 登录
# =========================
pwd = st.sidebar.text_input("Password", type="password")
if pwd != os.getenv("APP_PASSWORD", "admin123"):
    st.stop()

# =========================
# Session
# =========================
if "data" not in st.session_state:
    st.session_state.data = None
if "cal" not in st.session_state:
    st.session_state.cal = {}

# =========================
# 上传
# =========================
files = st.sidebar.file_uploader("Upload Excel", type=["xlsx"], accept_multiple_files=True)

if files:
    st.session_state.data = load_excel(files)

if st.session_state.data is None:
    st.warning("请上传Excel")
    st.stop()

df_raw = st.session_state.data.copy()
colmap = map_columns(df_raw)

st.sidebar.write("Mapping:", colmap)

# =========================
# Filter
# =========================
if colmap["category"]:
    cats = df_raw[colmap["category"]].dropna().unique()
    selected = st.sidebar.multiselect("Category", cats, default=cats[:2])
    df = df_raw[df_raw[colmap["category"]].isin(selected)] if selected else df_raw
else:
    df = df_raw

# =========================
# 🔥 计算字段
# =========================
df_raw["Base Part"] = df_raw.apply(lambda r: get_base_part(r, colmap), axis=1)
df["Base Part"] = df.apply(lambda r: get_base_part(r, colmap), axis=1)

df["Current Step"] = df.apply(lambda r: get_current_step(r, colmap), axis=1)
df["ETA"] = df.apply(lambda r: calc_eta(r, st.session_state.cal), axis=1)
df["Status"] = df["ETA"].apply(lambda x: "⚠️ Delayed" if x < datetime.now() else "On Track")
df["Progress"] = df.apply(lambda r: progress(r, colmap), axis=1)
df["Customer"] = df.apply(lambda r: get_customer(r, colmap), axis=1)

# =========================
# Tabs
# =========================
tabs = st.tabs([
"📋 All Items","🏭 Department","📅 Gantt","⚠️ Delayed",
"📊 Customer","🛠️ Engineering"
])

# =========================
# All Items（去重）
# =========================
with tabs[0]:
    st.write("Filtered:", len(df), "| Raw:", len(df_raw))

    df_unique = df.sort_values("ETA").drop_duplicates("Base Part")

    st.dataframe(df_unique)

# =========================
# Department
# =========================
with tabs[1]:
    dept = st.selectbox("Dept", df["Current Step"].dropna().unique())
    ddf = df[df["Current Step"]==dept]

    for i,r in ddf.iterrows():
        st.write(r.get(colmap["job"]), r.get(colmap["part"]))

        if st.button(f"Complete {i}"):
            nxt = next_step(r,colmap)
            st.session_state.data.at[i,colmap["current"]] = nxt
            st.rerun()

# =========================
# Gantt
# =========================
with tabs[2]:
    if colmap["job"]:
        job = st.selectbox("Job", df[colmap["job"]].dropna().unique())
        gdf = df[df[colmap["job"]]==job]

        if not gdf.empty:
            gdf["Start"]=datetime.now()
            fig = px.timeline(gdf,x_start="Start",x_end="ETA",y=colmap["part"],color="Current Step")
            st.plotly_chart(fig)

# =========================
# Delayed
# =========================
with tabs[3]:
    d = df[df["Status"]=="⚠️ Delayed"]
    st.dataframe(d)

# =========================
# Customer
# =========================
with tabs[4]:
    if colmap["exwork"]:
        df["Month"] = pd.to_datetime(df[colmap["exwork"]]).dt.to_period("M")
        st.dataframe(df.groupby(["Customer","Month"]).size())

# =========================
# Engineering（最终正确）
# =========================
with tabs[5]:
    st.title("Engineering WB Required")

    eng_df = df_raw[
        df_raw.apply(lambda r: is_engineering_required(r, df_raw, colmap), axis=1)
    ]

    eng_df = eng_df.drop_duplicates("Base Part")

    st.write("Total:", len(eng_df))
    st.dataframe(eng_df)
