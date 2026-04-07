import os
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import json
import plotly.express as px

os.environ["TZ"] = "Asia/Kuala_Lumpur"
st.set_page_config(layout="wide")

# =========================
# 默认工时
# =========================
DEFAULT_LEAD_TIME = {
    "Laser Cut": 4,
    "Laser Tube": 5,
    "Punching": 3,
    "Bending": 6,
    "Welding": 8,
    "Painting": 10,
    "Assembly": 12
}

# =========================
# 工具函数（👉 就是这里）
# =========================

def load_excel(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f, header=5, engine="openpyxl")
        df.columns = df.columns.astype(str).str.strip()
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def get_col(df, names):
    for n in names:
        if n in df.columns:
            return n
    return None

def map_columns(df):
    return {
        "job": get_col(df, ["JobNum/Asm"]),
        "part": get_col(df, ["Subpart Part Num"]),
        "main": get_col(df, ["Main Part Num"]),
        "category": get_col(df, ["Order Category"]),
        "exwork": get_col(df, ["Exwork Date"]),
        "order": get_col(df, ["Order Date"]),
        "current": get_col(df, ["Current Operation"]),
        "nest": get_col(df, ["Nesting Num"])
    }

# 🔥 Step解析（修复 Step17 merge）
def extract_steps(row):
    steps = []
    for col in row.index:
        if pd.isna(row[col]):
            continue
        if str(col).startswith("Step") or "Unnamed" in str(col):
            steps.append(row[col])
    return steps

def get_current_step(row, colmap):
    cur_col = colmap["current"]
    if cur_col and pd.notna(row[cur_col]):
        return row[cur_col]
    steps = extract_steps(row)
    return steps[0] if steps else None

def next_step(row, colmap):
    steps = extract_steps(row)
    cur = get_current_step(row, colmap)
    if cur in steps:
        i = steps.index(cur)
        if i + 1 < len(steps):
            return steps[i + 1]
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

# 🔥 新增：Customer 自动识别（你刚刚要的）
def get_customer_from_main(row, colmap):
    main_col = colmap.get("main")
    if not main_col or pd.isna(row.get(main_col)):
        return None
    val = str(row[main_col])
    return val.split("-")[0]

# =========================
# 登录
# =========================
if "login" not in st.session_state:
    st.session_state.login = False

pwd = st.sidebar.text_input("Password", type="password")
if pwd == os.getenv("APP_PASSWORD", "admin123"):
    st.session_state.login = True

if not st.session_state.login:
    st.stop()

# =========================
# Session
# =========================
if "data" not in st.session_state:
    st.session_state.data = None
if "cal" not in st.session_state:
    st.session_state.cal = {}
if "log" not in st.session_state:
    st.session_state.log = []
if "complete_time" not in st.session_state:
    st.session_state.complete_time = {}

# =========================
# 上传
# =========================
files = st.sidebar.file_uploader("Upload Excel", type=["xlsx"], accept_multiple_files=True)

if files:
    st.session_state.data = load_excel(files)

if st.session_state.data is None:
    st.stop()

df_raw = st.session_state.data.copy()
colmap = map_columns(df_raw)

# =========================
# Category filter
# =========================
if colmap["category"]:
    cats = df_raw[colmap["category"]].dropna().unique()
    selected = st.sidebar.multiselect(
        "Order Category",
        cats,
        default=[c for c in cats if c in ["New Awarded","New Revision"]]
    )
    df = df_raw[df_raw[colmap["category"]].isin(selected)]
else:
    df = df_raw

# =========================
# 计算字段
# =========================
df["Current Step"] = df.apply(lambda r: get_current_step(r, colmap), axis=1)
df["ETA"] = df.apply(lambda r: calc_eta(r, st.session_state.cal), axis=1)
df["Status"] = df["ETA"].apply(lambda x: "⚠️ Delayed" if x < datetime.now() else "On Track")
df["Progress"] = df.apply(lambda r: progress(r, colmap), axis=1)

# 🔥 Customer 自动生成
df["Customer"] = df.apply(lambda r: get_customer_from_main(r, colmap), axis=1)

# =========================
# Tabs
# =========================
tabs = st.tabs([
"All Items","Department","Capacity","Sales","Gantt",
"Delayed","Job Board","Stuck","Customer","Programmer","Engineering"
])

# =========================
# Customer Summary（完整版）
# =========================
with tabs[8]:
    st.title("Customer Summary")

    if colmap["exwork"]:
        main_parts = df[df[colmap["part"]].astype(str).str.endswith("-0")]

        main_parts["Month"] = pd.to_datetime(main_parts[colmap["exwork"]]).dt.to_period("M")

        summary = main_parts.groupby(["Customer","Month"]).size().reset_index(name="Count")
        st.dataframe(summary)

        cust = st.selectbox("Customer", summary["Customer"].unique())

        cust_df = main_parts[main_parts["Customer"] == cust]

        if colmap["order"]:
            cust_df["OrderDate"] = pd.to_datetime(cust_df[colmap["order"]])

            recent = cust_df[cust_df["OrderDate"] >= datetime.now() - timedelta(days=60)]
            trend = recent.groupby(cust_df["OrderDate"].dt.date).size()

            st.line_chart(trend)

            last7 = cust_df[cust_df["OrderDate"] >= datetime.now() - timedelta(days=7)]
            st.dataframe(last7)
