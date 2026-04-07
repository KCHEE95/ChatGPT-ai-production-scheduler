import os
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import json
import plotly.express as px

# =========================
# 时区
# =========================
os.environ["TZ"] = "Asia/Kuala_Lumpur"

st.set_page_config(layout="wide")

# =========================
# 默认 Lead Time
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
# 工具函数
# =========================
def load_excel(files):
    df_list = []
    for f in files:
        df = pd.read_excel(f, header=5, engine="openpyxl")
        df_list.append(df)
    return pd.concat(df_list, ignore_index=True)

def extract_steps(row):
    return [row[f"Step {i}"] for i in range(1,21) if f"Step {i}" in row and pd.notna(row[f"Step {i}"])]

def get_current_step(row):
    if pd.isna(row.get("Current Operation")):
        steps = extract_steps(row)
        return steps[0] if steps else None
    return row["Current Operation"]

def next_step(row):
    steps = extract_steps(row)
    cur = get_current_step(row)
    if cur in steps:
        i = steps.index(cur)
        if i+1 < len(steps):
            return steps[i+1]
    return None

def calculate_eta(row, cal):
    total = 0
    for s in extract_steps(row):
        total += cal.get(s, DEFAULT_LEAD_TIME.get(s,5))
    return datetime.now() + timedelta(hours=total)

def get_progress(row):
    steps = extract_steps(row)
    cur = get_current_step(row)
    if not steps: return 0
    if cur not in steps: return 100
    return int((steps.index(cur)/len(steps))*100)

# =========================
# 登录
# =========================
if "login" not in st.session_state:
    st.session_state.login = False

PASSWORD = os.getenv("APP_PASSWORD", "admin123")
pwd = st.sidebar.text_input("🔐 Password", type="password")

if pwd == PASSWORD:
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
st.sidebar.title("📁 Upload Excel files exported from Epicor")

files = st.sidebar.file_uploader("", type=["xlsx"], accept_multiple_files=True)
if files:
    st.session_state.data = load_excel(files)

# Calibration
st.sidebar.subheader("Auto-Calibration")

if st.sidebar.button("🔄 Reset All Calibrations"):
    st.session_state.cal = {}

cal_file = st.sidebar.file_uploader("Load Calibration", type=["json"])
if cal_file:
    st.session_state.cal = json.load(cal_file)

st.sidebar.download_button("📥 Export Calibration (JSON)", json.dumps(st.session_state.cal), "cal.json")

# Order Category Filter
if st.session_state.data is not None:
    categories = st.session_state.data.get("Order Category", pd.Series()).dropna().unique()
    selected_cat = st.sidebar.multiselect("Order Category Filter", categories, default=[c for c in categories if c in ["New Awarded","New Revision"]])
else:
    selected_cat = []

# Change log
st.sidebar.download_button("📥 Export Change Log (JSON)", json.dumps(st.session_state.log), "log.json")
if st.sidebar.button("🗑️ Clear Change Log"):
    st.session_state.log = []

# =========================
# 数据准备
# =========================
if st.session_state.data is None:
    st.stop()

df_raw = st.session_state.data.copy()

df = df_raw.copy()
if selected_cat:
    df = df[df["Order Category"].isin(selected_cat)]

df["Current Step"] = df.apply(get_current_step, axis=1)
df["ETA"] = df.apply(lambda r: calculate_eta(r, st.session_state.cal), axis=1)
df["Status"] = df["ETA"].apply(lambda x: "⚠️ Delayed" if x < datetime.now() else "On Track")
df["Progress"] = df.apply(get_progress, axis=1)

# =========================
# Tabs
# =========================
tabs = st.tabs([
"📋 All Items","🏭 Department Workbench","📈 Capacity Dashboard",
"🔍 Sales Query","📅 Job Gantt Chart","⚠️ Delayed Alerts",
"📊 Job Progress Board","⏰ Stuck Alerts","📊 Customer Summary",
"🛠️ Programmer Board","🛠️ Engineering WB Required"
])

# =========================
# 1 All Items
# =========================
with tabs[0]:
    st.dataframe(df.sort_values("ETA"))

    for i,row in df.iterrows():
        with st.expander("🔍 View full operation chain"):
            st.write(extract_steps(row))

# =========================
# 2 Department WB
# =========================
with tabs[1]:
    dept = st.selectbox("Department", df["Current Step"].dropna().unique())
    ddf = df[df["Current Step"]==dept]

    for i,row in ddf.iterrows():
        st.markdown(f"### {row.get('JobNum')} - {row.get('PartNum')}")
        st.progress(row["Progress"]/100)
        st.write(row["Status"], row["ETA"])

        c1,c2 = st.columns(2)

        if c1.button(f"Complete & Next {i}"):
            nxt = next_step(row)
            st.session_state.data.at[i,"Current Operation"]=nxt
            st.session_state.complete_time[i]=datetime.now()

            st.session_state.log.append({
                "time":str(datetime.now()),
                "job":row.get("JobNum"),
                "from":row["Current Step"],
                "to":nxt
            })
            st.rerun()

        actual = c2.number_input(f"Actual hrs {i}",0.0)
        if c2.button(f"Calibrate {i}"):
            old = st.session_state.cal.get(dept,DEFAULT_LEAD_TIME.get(dept,5))
            st.session_state.cal[dept]=0.7*old+0.3*actual

    # 导出 Excel
    st.download_button("📥 Download updated Excel", df.to_csv(index=False), "updated.csv")

# =========================
# 3 Capacity
# =========================
with tabs[2]:
    cap = df.groupby("Current Step").size().reset_index(name="Count")
    cap["Capacity"]=10
    cap["Load%"]=cap["Count"]/cap["Capacity"]*100
    st.dataframe(cap)

# =========================
# 4 Sales Query
# =========================
with tabs[3]:
    q = st.text_input("Search Job / PO / Part")
    sdf = df[df.astype(str).apply(lambda x: x.str.contains(q, case=False)).any(axis=1)] if q else df

    st.write("Total:",len(sdf))
    st.write("Delayed:",len(sdf[sdf["Status"]=="⚠️ Delayed"]))
    st.dataframe(sdf)

# =========================
# 5 Gantt
# =========================
with tabs[4]:
    job = st.selectbox("Select Job", df["JobNum"].dropna().unique())
    gdf = df[df["JobNum"]==job]
    if not gdf.empty:
        gdf["Start"]=datetime.now()
        fig = px.timeline(gdf, x_start="Start", x_end="ETA", y="PartNum", color="Current Step")
        st.plotly_chart(fig)

# =========================
# 6 Delayed
# =========================
with tabs[5]:
    d = df[df["Status"]=="⚠️ Delayed"]
    st.bar_chart(d.groupby("Current Step").size())
    st.dataframe(d)

# =========================
# 7 Job Progress
# =========================
with tabs[6]:
    j = df.groupby("JobNum").agg(total=("PartNum","count"),
        delayed=("Status",lambda x:(x=="⚠️ Delayed").sum()))
    j["progress"]=100*(1-j["delayed"]/j["total"])
    st.dataframe(j)

# =========================
# 8 Stuck
# =========================
with tabs[7]:
    th = st.number_input("Threshold hrs",1,100,24)
    stuck=[]
    for i,t in st.session_state.complete_time.items():
        hours=(datetime.now()-t).total_seconds()/3600
        if hours>th:
            stuck.append((i,hours))
    st.write(stuck)

# =========================
# 9 Customer Summary
# =========================
with tabs[8]:
    if "Exwork Date" in df.columns:
        df["Month"]=pd.to_datetime(df["Exwork Date"]).dt.to_period("M")
        st.dataframe(df.groupby(["Customer","Month"]).size())

# =========================
# 10 Programmer
# =========================
with tabs[9]:
    p = df[df["Current Step"].isin(["Laser Cut","Laser Tube","Punching"])]
    p = p[p["Nesting Num"].isna()]
    st.dataframe(p)

# =========================
# 11 Engineering
# =========================
with tabs[10]:
    def no_step(r):
        return all(pd.isna(r.get(f"Step {i}")) for i in range(1,21))
    e = df_raw[df_raw.apply(no_step,axis=1)]
    st.dataframe(e)
