import io
import re
from datetime import time
from typing import Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule

# --- Student mapping ---
students = {
    1: "Ishmal",
    2: "Owais",
    3: "Neha",
    4: "Sarah",
    6: "Musfira",
    7: "Ayesha",
    8: "Junaid",
    9: "Abd-ur-Rehman",
    10: "Samra",
    13: "Shahida",
    15: "Seerat",
    16: "Usman",
    17: "Mahad",
    18: "Eman",
    19: "Izaz",
    20: "Kiran",
    21: "Hammad",
    24: "Faiza",
    26: "Laiba",
    29: "Ushna",
    30: "Ali",
    31: "Adnan",
    32: "Yaseen",
    33: "Abbas"
}

# --- Fixed check-in threshold (09:02) ---
CHECKIN_THRESHOLD = time(9, 2, 0)

# --- Streamlit setup ---
st.set_page_config(page_title="Attendance .dat → Excel", page_icon="📊", layout="centered")
st.title("📊 Attendance Processor (.dat → Excel)")

st.markdown(
    "Upload your ZKTeco `.dat` (or text) export. "
    "The app will extract **UserID**, match with **Name**, keep the **first (Check-In)** "
    "and **last (Check-Out) punch per day**, and mark arrivals after **09:02** as **Late** (red)."
)

# --- Controls ---
weekends_off = st.checkbox("Ignore weekends (Sat/Sun)", value=False)

uploaded = st.file_uploader("Upload .dat file", type=["dat", "txt", "csv", "log"])

# --- Helpers ---
CANDIDATE_USER_COLS = ["userid", "user_id", "pin", "enrollid", "empid", "id"]
CANDIDATE_TIME_COLS = ["timestamp", "time", "datetime", "logtime", "punch time", "punch_time"]

def _try_read_csv(buf: io.BytesIO) -> Optional[pd.DataFrame]:
    buf.seek(0)
    for sep in ["\t", ",", ";", "|"]:
        buf.seek(0)
        try:
            df = pd.read_csv(buf, sep=sep, engine="python")
            if df.shape[1] >= 2:
                return df
        except Exception:
            pass
    buf.seek(0)
    try:
        return pd.read_csv(buf, delim_whitespace=True, engine="python", header=None)
    except Exception:
        return None

def _coerce_encoding(file_bytes: bytes) -> io.BytesIO:
    try:
        text = file_bytes.decode("utf-8")
    except UnicodeDecodeError:
        text = file_bytes.decode("latin-1", errors="ignore")
    return io.BytesIO(text.encode("utf-8"))

def _detect_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    lower_cols = {c.lower(): c for c in df.columns.astype(str)}
    user_col, ts_col = None, None
    for k in CANDIDATE_USER_COLS:
        if k in lower_cols:
            user_col = lower_cols[k]; break
    for k in CANDIDATE_TIME_COLS:
        if k in lower_cols:
            ts_col = lower_cols[k]; break
    if user_col is None and ts_col is None and df.shape[1] >= 2:
        user_col, ts_col = df.columns[0], df.columns[1]
    return user_col, ts_col

def _extract_dataframe(file_bytes: bytes) -> pd.DataFrame:
    buf = _coerce_encoding(file_bytes)
    raw = _try_read_csv(buf)
    if raw is None or raw.empty:
        st.error("Could not parse the file. Check format or try another export.")
        st.stop()
    if any(not isinstance(c, str) for c in raw.columns):
        raw.columns = [f"col_{i+1}" for i in range(raw.shape[1])]
    user_col, ts_col = _detect_columns(raw)
    if user_col is None or ts_col is None:
        st.error(f"Could not detect UserID/Timestamp columns. Found: {list(raw.columns)}")
        st.stop()
    df = raw[[user_col, ts_col]].copy()
    df.columns = ["UserID", "Timestamp"]
    df["UserID"] = df["UserID"].astype(str).str.strip()

    def parse_dt(x):
        x = str(x).strip()
        x = re.sub(r"\s+", " ", x)
        for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S",
                    "%Y/%m/%d %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%Y-%m-%d %H:%M",
                    "%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M"):
            try:
                return pd.to_datetime(x, format=fmt)
            except Exception:
                pass
        return pd.to_datetime(x, errors="coerce")

    df["Timestamp"] = df["Timestamp"].apply(parse_dt)
    df = df.dropna(subset=["Timestamp"]).copy()
    df["Date"] = df["Timestamp"].dt.date
    if weekends_off:
        df = df[df["Timestamp"].dt.weekday < 5]
    return df

def _summarize_attendance(df: pd.DataFrame, late_t: time) -> pd.DataFrame:
    # First punch (Check-In)
    first = df.groupby(["UserID", "Date"], as_index=False)["Timestamp"].min().rename(columns={"Timestamp": "CheckIn"})
    # Last punch (Check-Out)
    last = df.groupby(["UserID", "Date"], as_index=False)["Timestamp"].max().rename(columns={"Timestamp": "CheckOut"})
    # Merge both
    merged = pd.merge(first, last, on=["UserID", "Date"], how="inner")

    # convert UserID to numeric for mapping + sorting
    merged["UserID_int"] = pd.to_numeric(merged["UserID"], errors="coerce")
    merged["Name"] = merged["UserID_int"].map(students)
    merged = merged.dropna(subset=["Name"]).copy()

    # Mark status based on CheckIn
    merged["Status"] = merged["CheckIn"].dt.time.apply(lambda t: "Late" if t > late_t else "On Time")
    merged["Minutes Late"] = [
        max(0, (t.hour - late_t.hour) * 60 + (t.minute - late_t.minute))
        for t in merged["CheckIn"]
    ]

    # Sort by ID, then Date
    merged = merged.sort_values(by=["UserID_int", "Date"])
    merged.drop(columns=["UserID_int"], inplace=True)
    return merged[["UserID", "Name", "Date", "CheckIn", "CheckOut", "Status", "Minutes Late"]]

def _to_styled_excel(df: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    df.to_excel(bio, index=False, sheet_name="Attendance")
    bio.seek(0)
    wb = load_workbook(bio)
    ws = wb.active
    header = [c.value for c in ws[1]]
    try:
        status_idx = header.index("Status") + 1
    except ValueError:
        status_idx = None
    red = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
    if status_idx:
        last_row = ws.max_row
        last_col_letter = ws.cell(row=1, column=ws.max_column).column_letter
        status_col_letter = ws.cell(row=1, column=status_idx).column_letter
        rule = FormulaRule(formula=[f'${status_col_letter}2="Late"'], fill=red)
        ws.conditional_formatting.add(f"A2:{last_col_letter}{last_row}", rule)
    for col in ws.columns:
        max_len, col_letter = 0, col[0].column_letter
        for cell in col:
            v = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(v))
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 40)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()

# --- Main flow ---
if uploaded:
    file_bytes = uploaded.read()
    df_raw = _extract_dataframe(file_bytes)
    st.success(f"Parsed {len(df_raw):,} log rows. Computing Check-In/Check-Out per user/day…")

    summary = _summarize_attendance(df_raw, CHECKIN_THRESHOLD)
    st.dataframe(summary.head(50))

    excel_bytes = _to_styled_excel(summary)
    st.download_button(
        "⬇️ Download Excel (Check-In + Check-Out, Late highlighted)",
        data=excel_bytes,
        file_name="attendance_marked.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
