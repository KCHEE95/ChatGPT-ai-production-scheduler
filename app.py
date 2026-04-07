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
# 读取 Excel（修复 header）
# =========================
def load_excel(files):
    dfs = []
    for f in files:
        df = pd.read_excel(f, header=5, engine="openpyxl")
        df.columns = df.columns.astype(str).str.strip()
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

# =========================
# 自动列 mapping（适配你Excel）
# =========================
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

# =========================
# Step解析（修复Step17）
# =========================
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
# Sidebar
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
    selected = st.sidebar.multiselect("Category", cats, default=[c for c in cats if c in ["New Awarded","New Revision"]])
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

# =========================
# Tabs
# =========================
tabs = st.tabs([
"All Items","Department WB","Capacity","Sales Query","Gantt",
"Delayed","Job Board","Stuck","Customer","Programmer","Engineering"
])

# =========================
# All Items
# =========================
with tabs[0]:
    st.dataframe(df.sort_values("ETA"))
    for _,r in df.iterrows():
        with st.expander("🔍 Operation Chain"):
            st.write(extract_steps(r))

# =========================
# Department
# =========================
with tabs[1]:
    dept = st.selectbox("Dept", df["Current Step"].dropna().unique())
    ddf = df[df["Current Step"]==dept]

    for i,r in ddf.iterrows():
        st.markdown(f"### {r[colmap['job']]} - {r[colmap['part']]}")
        st.progress(r["Progress"]/100)

        c1,c2 = st.columns(2)

        if c1.button(f"Complete {i}"):
            nxt = next_step(r,colmap)
            st.session_state.data.at[i,colmap["current"]] = nxt
            st.session_state.complete_time[i]=datetime.now()
            st.session_state.log.append({"time":str(datetime.now()),"from":r["Current Step"],"to":nxt})
            st.rerun()

        actual = c2.number_input(f"hrs {i}",0.0)
        if c2.button(f"Cal {i}"):
            old = st.session_state.cal.get(dept,DEFAULT_LEAD_TIME.get(dept,5))
            st.session_state.cal[dept]=0.7*old+0.3*actual

    st.download_button("Download Excel", df.to_csv(index=False))

# =========================
# Capacity
# =========================
with tabs[2]:
    cap = df.groupby("Current Step").size().reset_index(name="Count")
    cap["Capacity"]=10
    cap["Load%"]=cap["Count"]/cap["Capacity"]*100
    st.dataframe(cap)

# =========================
# Sales Query
# =========================
with tabs[3]:
    q = st.text_input("Search")
    sdf = df[df.astype(str).apply(lambda x: x.str.contains(q, case=False)).any(axis=1)] if q else df
    st.dataframe(sdf)

# =========================
# Gantt
# =========================
with tabs[4]:
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
with tabs[5]:
    d = df[df["Status"]=="⚠️ Delayed"]
    st.dataframe(d)

# =========================
# Job Board
# =========================
with tabs[6]:
    if colmap["job"]:
        j = df.groupby(colmap["job"]).agg(total=(colmap["part"],"count"),
        delayed=("Status",lambda x:(x=="⚠️ Delayed").sum()))
        j["progress"]=100*(1-j["delayed"]/j["total"])
        st.dataframe(j)

# =========================
# Stuck
# =========================
with tabs[7]:
    th = st.number_input("Threshold",1,100,24)
    stuck=[]
    for i,t in st.session_state.complete_time.items():
        h=(datetime.now()-t).total_seconds()/3600
        if h>th:
            stuck.append((i,h))
    st.write(stuck)

# =========================
# Customer
# =========================
with tabs[8]:
    if colmap["exwork"]:
        df["Month"]=pd.to_datetime(df[colmap["exwork"]]).dt.to_period("M")
        st.dataframe(df.groupby(["Month"]).size())

# =========================
# Programmer
# =========================
with tabs[9]:
    if colmap["current"] and colmap["nest"]:
        p=df[(df[colmap["current"]].isin(["Laser Cut","Laser Tube","Punching"])) & (df[colmap["nest"]].isna())]
        st.dataframe(p)

# =========================
# Engineering
# =========================
with tabs[10]:
    def no_step(r):
        return all(pd.isna(r[c]) for c in r.index if "Step" in str(c) or "Unnamed" in str(c))
    e=df_raw[df_raw.apply(no_step,axis=1)]
    st.dataframe(e)
