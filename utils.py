import pandas as pd
from datetime import datetime, timedelta

DEFAULT_LEAD_TIME = {
    "Laser Cut": 4,
    "Laser Tube": 5,
    "Punching": 3,
    "Bending": 6,
    "Welding": 8,
    "Painting": 10,
    "Assembly": 12
}

def load_excel(files):
    df_list = []
    for f in files:
        df = pd.read_excel(f, header=5)
        df_list.append(df)
    df = pd.concat(df_list, ignore_index=True)
    return df

def extract_steps(row):
    steps = []
    for i in range(1, 21):
        col = f"Step {i}"
        if col in row and pd.notna(row[col]):
            steps.append(row[col])
    return steps

def calculate_eta(row, calibration):
    steps = extract_steps(row)
    total_hours = 0
    for step in steps:
        total_hours += calibration.get(step, DEFAULT_LEAD_TIME.get(step, 5))
    return datetime.now() + timedelta(hours=total_hours)

def get_current_step(row):
    if pd.isna(row.get("Current Operation")):
        steps = extract_steps(row)
        return steps[0] if steps else None
    return row["Current Operation"]

def next_step(row):
    steps = extract_steps(row)
    current = get_current_step(row)
    if current in steps:
        idx = steps.index(current)
        if idx + 1 < len(steps):
            return steps[idx + 1]
    return None
