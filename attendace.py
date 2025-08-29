import io
import re
from datetime import datetime, time
from typing import Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
from zk import ZK

# --- Employee mapping ---
employees = {
    1: "Ishmal", 2: "Owais", 3: "Neha", 4: "Sarah", 6: "Musfira", 7: "Ayesha",
    8: "Junaid", 9: "Abd-ur-Rehman", 10: "Samra", 13: "Shahida", 15: "Seerat",
    16: "Usman", 17: "Mahad", 18: "Eman", 19: "Izaz", 20: "Kiran", 21: "Hammad",
    24: "Faiza", 26: "Laiba", 29: "Ushna", 30: "Ali", 31: "Adnan", 32: "Yaseen",
    33: "Abbas"
}

CHECKIN_THRESHOLD = time(9, 2, 0)

# --- ZKTeco device config ---
DEVICE_IP = "192.168.18.200"   # change to your device IP
DEVICE_PORT = 4370             # default port

# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Attendance Dashboard", page_icon="📊", layout="wide")
st.title("📊 Attendance Processor")

st.sidebar.header("⚙️ Settings")
weekends_off = st.sidebar.checkbox("Ignore weekends (Sat/Sun)", value=False)

# Date pickers for ZK device
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date", datetime.today().replace(day=1).date())
with col2:
    end_date = st.date_input("End Date", datetime.today().date())

# --- Fetch logs from device ---
def fetch_logs(start_date, end_date):
    zk = ZK(DEVICE_IP, port=DEVICE_PORT, timeout=5)
    conn = None
    try:
        st.info("⏳ Connecting to device...")
        conn = zk.connect()
        st.success("✅ Connected to device!")

        conn.disable_device()
        attendances = conn.get_attendance()
        conn.enable_device()
        conn.disconnect()

        # Convert to DataFrame
        records = []
        for att in attendances:
            records.append({
                "UserID": att.user_id,
                "Timestamp": att.timestamp
            })
        df = pd.DataFrame(records)
        df["Timestamp"] = pd.to_datetime(df["Timestamp"])
        mask = (df["Timestamp"].dt.date >= start_date) & (df["Timestamp"].dt.date <= end_date)
        df_filtered = df.loc[mask]
        return df_filtered
    except Exception as e:
        if conn:
            conn.disconnect()
        st.error(f"❌ Error: {e}")
        return pd.DataFrame()

# --- Attendance processing helpers ---
def _summarize_attendance(df: pd.DataFrame, late_t: time) -> pd.DataFrame:
    df["Date"] = df["Timestamp"].dt.date
    # First check-in per user per day
    first = df.groupby(["UserID", "Date"], as_index=False)["Timestamp"].min().rename(columns={"Timestamp": "CheckIn"})
    # Last check-out per user per day
    last = df.groupby(["UserID", "Date"], as_index=False)["Timestamp"].max().rename(columns={"Timestamp": "CheckOut"})
    # Merge on UserID and Date
    merged = pd.merge(first, last, on=["UserID", "Date"], how="inner")
    
    # Map employee names
    merged["UserID_int"] = pd.to_numeric(merged["UserID"], errors="coerce")
    merged["Name"] = merged["UserID_int"].map(employees)
    merged = merged.dropna(subset=["Name"]).copy()
    
    # Extract time from timestamps
    merged["CheckIn"] = merged["CheckIn"].dt.time
    merged["CheckOut"] = merged["CheckOut"].dt.time
    
    # Status
    merged["Status"] = merged["CheckIn"].apply(lambda t: "Late" if t > late_t else "On Time")
    merged["Minutes Late"] = [(max(0, (t.hour - late_t.hour) * 60 + (t.minute - late_t.minute))) for t in merged["CheckIn"]]
    
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

# --- Fetch button ---
# --- Fetch button ---
if st.button("🔍 Fetch Attendance"):
    if start_date > end_date:
        st.error("⚠️ Start Date cannot be after End Date!")
    else:
        df_logs = fetch_logs(start_date, end_date)
        if not df_logs.empty:
            st.success(f"✅ Fetched {len(df_logs)} records from device")

            # --- Attendance summary ---
            summary = _summarize_attendance(df_logs, CHECKIN_THRESHOLD)

            # --- Metrics ---
            total_logs = len(summary)
            total_late = (summary["Status"] == "Late").sum()
            total_on_time = (summary["Status"] == "On Time").sum()

            col1, col2, col3 = st.columns(3)
            col1.metric("📄 Total Records", f"{total_logs}")
            col2.metric("⏰ Total Late", f"{total_late}")
            col3.metric("✅ Total On Time", f"{total_on_time}")

            st.markdown("---")
            with st.expander("🗂️ View Detailed Attendance Table"):
                st.dataframe(summary)  # Only this detailed summary table, not raw logs

            # --- Excel download ---
            excel_bytes = _to_styled_excel(summary)
            st.download_button(
                "⬇️ Download Excel",
                data=excel_bytes,
                file_name="attendance_marked.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # --- Late & Leave count summary ---
            late_counts = summary.groupby("Name")["Status"].apply(lambda x: (x == "Late").sum()).reset_index()
            late_counts.rename(columns={"Status": "Total Lates"}, inplace=True)

            dates_in_data = pd.date_range(start=start_date, end=end_date).date
            all_data = pd.MultiIndex.from_product([employees.values(), dates_in_data], names=["Name", "Date"])
            full_df = pd.DataFrame(index=all_data).reset_index()
            merged_summary = pd.merge(full_df, summary[["Name", "Date"]], on=["Name","Date"], how="left", indicator=True)
            leave_counts = merged_summary.groupby("Name")["_merge"].apply(lambda x: (x == "left_only").sum()).reset_index()
            leave_counts.rename(columns={"_merge": "Total Leaves"}, inplace=True)

            summary_counts = pd.DataFrame({"Name": list(employees.values())})
            summary_counts = summary_counts.merge(late_counts, on="Name", how="left").fillna(0)
            summary_counts = summary_counts.merge(leave_counts, on="Name", how="left").fillna(0)
            summary_counts["Total Lates"] = summary_counts["Total Lates"].astype(int)
            summary_counts["Total Leaves"] = summary_counts["Total Leaves"].astype(int)

            with st.expander("📋 Total Late & Leave Count of Employees"):
                def highlight_late_leave(row):
                    return [
                        "",
                        'background-color: #FFCDD2' if row["Total Lates"] > 0 else '',
                        'background-color: #BBDEFB' if row["Total Leaves"] > 0 else ''
                    ]
                st.dataframe(summary_counts.style.apply(highlight_late_leave, axis=1))
        else:
            st.warning("⚠️ No records found for selected dates.")
