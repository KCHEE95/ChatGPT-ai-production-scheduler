import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import json

os.environ["TZ"] = "Asia/Kuala_Lumpur"
st.set_page_config(layout="wide")

DEFAULT_LEAD_TIME = {
    "Laser Cut": 4,"Laser Tube": 5,"Punching": 3,
    "Bending": 6,"Welding": 8,"Painting": 10,"Assembly": 12
}

# =========================
# Excel Auto Detect
# =========================
def detect_header(file):
    raw = pd.read_excel(file, header=None)
    for i,row in raw.iterrows():
        if row.astype(str).str.contains("Main Part",case=False).any():
            return i
    return 5

def load_excel(files):
    dfs=[]
    for f in files:
        h=detect_header(f)
        df=pd.read_excel(f,header=h)
        df.columns=df.columns.astype(str).str.strip()
        dfs.append(df)
    return pd.concat(dfs,ignore_index=True)

# =========================
# Mapping
# =========================
def find_col(df,keys):
    for c in df.columns:
        for k in keys:
            if k.lower() in c.lower():
                return c
    return None

def map_columns(df):
    return {
        "job": find_col(df,["job"]),
        "part": find_col(df,["subpart"]),
        "main": find_col(df,["main part"]),
        "category": find_col(df,["category"]),
        "exwork": find_col(df,["exwork"]),
        "order": find_col(df,["order date"]),
        "current": find_col(df,["current"]),
        "nest": find_col(df,["nest"])
    }

# =========================
# Steps
# =========================
def extract_steps(r):
    return [r[c] for c in r.index if ("step" in str(c).lower() or "unnamed" in str(c).lower()) and pd.notna(r[c])]

def get_current(r,colmap):
    if colmap["current"] and pd.notna(r[colmap["current"]]):
        return r[colmap["current"]]
    s=extract_steps(r)
    return s[0] if s else None

def next_step(r,colmap):
    s=extract_steps(r)
    cur=get_current(r,colmap)
    if cur in s:
        i=s.index(cur)
        if i+1<len(s):
            return s[i+1]
    return None

def calc_eta(r,cal):
    return datetime.now()+timedelta(hours=sum(cal.get(s,DEFAULT_LEAD_TIME.get(s,5)) for s in extract_steps(r)))

def progress(r,colmap):
    s=extract_steps(r)
    cur=get_current(r,colmap)
    return int(100*s.index(cur)/len(s)) if cur in s else 100 if s else 0

# =========================
# Base Part
# =========================
def get_base(r,colmap):
    val=None
    if colmap["part"] and pd.notna(r[colmap["part"]]):
        val=str(r[colmap["part"]])
    elif colmap["main"]:
        val=str(r[colmap["main"]])
    if not val:return None
    p=val.split("-")
    return "-".join(p[:-1]) if p[-1].isdigit() else val

# =========================
# Customer
# =========================
def get_customer(r,colmap):
    if colmap["main"] and pd.notna(r[colmap["main"]]):
        return str(r[colmap["main"]]).split("-")[0]

# =========================
# Engineering logic
# =========================
def is_eng(r,df,colmap):
    base=r["Base"]
    has=df[(df["Base"]==base)&(df[colmap["job"]].notna())]
    return has.empty

# =========================
# Login
# =========================
if st.sidebar.text_input("Password",type="password")!="admin123":
    st.stop()

# =========================
# Upload
# =========================
files=st.sidebar.file_uploader("Upload",type=["xlsx"],accept_multiple_files=True)
if not files: st.stop()

df_raw=load_excel(files)
colmap=map_columns(df_raw)

df_raw["Base"]=df_raw.apply(lambda r:get_base(r,colmap),axis=1)

df=df_raw.copy()

# =========================
# Calc
# =========================
df["Step"]=df.apply(lambda r:get_current(r,colmap),axis=1)
df["ETA"]=df.apply(lambda r:calc_eta(r,{}),axis=1)
df["Status"]=df["ETA"].apply(lambda x:"⚠️ Delayed" if x<datetime.now() else "On Track")
df["Progress"]=df.apply(lambda r:progress(r,colmap),axis=1)
df["Customer"]=df.apply(lambda r:get_customer(r,colmap),axis=1)

# =========================
# Tabs
# =========================
tabs=st.tabs([
"All","Dept","Capacity","Sales","Gantt",
"Delayed","Job","Stuck","Customer","Programmer","Engineering"
])

# All
with tabs[0]:
    st.dataframe(df.sort_values("ETA").drop_duplicates("Base"))

# Dept
with tabs[1]:
    d=st.selectbox("Dept",df["Step"].dropna().unique())
    ddf=df[df["Step"]==d]
    for i,r in ddf.iterrows():
        st.write(r[colmap["job"]],r[colmap["part"]])
        if st.button(f"Done{i}"):
            df_raw.at[i,colmap["current"]]=next_step(r,colmap)
            st.rerun()

# Capacity
with tabs[2]:
    c=df.groupby("Step").size().reset_index(name="Count")
    c["Cap"]=10
    c["Load%"]=c["Count"]/c["Cap"]*100
    st.dataframe(c)

# Sales
with tabs[3]:
    q=st.text_input("Search")
    st.dataframe(df[df.astype(str).apply(lambda x:x.str.contains(q,case=False)).any(axis=1)])

# Gantt
with tabs[4]:
    if colmap["job"]:
        j=st.selectbox("Job",df[colmap["job"]].dropna().unique())
        g=df[df[colmap["job"]]==j]
        g["Start"]=datetime.now()
        st.plotly_chart(px.timeline(g,x_start="Start",x_end="ETA",y=colmap["part"],color="Step"))

# Delayed
with tabs[5]:
    st.dataframe(df[df["Status"]=="⚠️ Delayed"])

# Job board
with tabs[6]:
    st.dataframe(df.groupby(colmap["job"]).agg(total=(colmap["part"],"count")))

# Stuck
with tabs[7]:
    st.write("TODO")

# Customer
with tabs[8]:
    df["Month"]=pd.to_datetime(df[colmap["exwork"]]).dt.to_period("M")
    st.dataframe(df.groupby(["Customer","Month"]).size())

# Programmer
with tabs[9]:
    st.dataframe(df[(df["Step"].isin(["Laser Cut","Laser Tube","Punching"]))&(df[colmap["nest"]].isna())])

# Engineering
with tabs[10]:
    eng=df_raw[df_raw.apply(lambda r:is_eng(r,df_raw,colmap),axis=1)]
    st.dataframe(eng.drop_duplicates("Base"))
