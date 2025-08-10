import io
import os
from typing import List, Optional

import pandas as pd
import streamlit as st

# -------- Config --------
st.set_page_config(page_title="Unified Services Dashboard", layout="wide")
DEFAULT_PATH = r"C:\Users\LaxmanPareet\Downloads\dashboard.xlsx"

MODULE_LABELS = {
    "AUGGOLDINVESTMENT": "Digital Gold",
    "MobilePrepaidRecharge": "Mobile Recharge", 
    "HealthTestPayment": "Wellness Center",
}

# Fixed SUCCESS_TOKENS definition
SUCCESS_TOKENS = {"success", "successful", "succeeded", "completed", "ok", "done", "captured", "paid"}
TRUE_TOKENS = {"1", "true", "yes", "y"}

# -------- Helpers --------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names to lowercase with underscores"""
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(r"\s+", "_", regex=True)
        .str.replace("-", "_")
        .str.lower()
    )
    return df

def normalize_values(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize string values and uppercase service names"""
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].astype(str).str.strip()
    for svc in ["servicename", "service_name", "service"]:
        if svc in df.columns:
            df[svc] = df[svc].astype(str).str.strip().str.upper()
    return df

def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Find column by exact match or contains"""
    cols = list(df.columns)
    # Exact match first
    for name in candidates:
        if name in cols:
            return name
    # Contains match
    for name in candidates:
        for c in cols:
            if name in c:
                return c
    return None

def is_success_value(v) -> bool:
    """Check if value indicates success"""
    if pd.isna(v):
        return False
    s = str(v).strip().lower()
    return s in SUCCESS_TOKENS or s in TRUE_TOKENS

def like_filter(series: pd.Series, query: str) -> pd.Series:
    """Case-insensitive contains filter"""
    if series is None or query is None or str(query).strip() == "":
        return pd.Series([True] * len(series), index=series.index)
    s = series.astype(str).str.lower()
    q = str(query).strip().lower()
    return s.str.contains(q, na=False)

@st.cache_data(show_spinner=False)
def load_excel_all_sheets(path: str) -> pd.DataFrame:
    """Load all sheets from Excel with error handling"""
    try:
        xls = pd.ExcelFile(path)
        frames = []
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
                df = normalize_columns(df)
                df = normalize_values(df)
                df["__sheet__"] = sheet
                frames.append(df)
            except Exception as e:
                st.warning(f"Error reading sheet '{sheet}': {str(e)}")
                continue
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return pd.DataFrame()

def add_success_column(df: pd.DataFrame) -> pd.DataFrame:
    """Add success column based on status fields"""
    df = df.copy()
    status_col = find_col(
        df,
        ["status", "txn_status", "transaction_status", "payment_status", "invest_status", "investment_status"],
    )
    if status_col:
        df["__success__"] = df[status_col].map(is_success_value)
    else:
        df["__success__"] = False
    return df

def prepare_display(df: pd.DataFrame) -> pd.DataFrame:
    """Prepare dataframe for display with priority columns"""
    priority = [
        "servicename",
        "mobile",
        "mobile_no", 
        "mobile_number",
        "mobilenumber",
        "tvamcustid",
        "paymentrefno",
        "utr",
        "status",
        "payment_status",
        "txn_status",
        "transaction_status",
        "amount",
        "created_at",
        "createdon",
        "date",
        "__sheet__",
        "__success__",
    ]
    # Remove duplicates while preserving order
    seen = set()
    unique_priority = []
    for col in priority:
        if col in df.columns and col not in seen:
            unique_priority.append(col)
            seen.add(col)
    
    other_cols = [c for c in df.columns if c not in seen]
    cols = unique_priority + other_cols
    return df[cols]

def bytes_excel(df: pd.DataFrame) -> bytes:
    """Convert dataframe to Excel bytes"""
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return bio.getvalue()

# -------- UI --------
st.title("Unified Services Dashboard")

# File path input
path = st.text_input("Excel file path", value=DEFAULT_PATH)
if not path or not os.path.exists(path):
    st.error("Excel file not found. Please check the path.")
    st.stop()

# Load data
with st.spinner("Loading Excel file..."):
    raw = load_excel_all_sheets(path)

if raw.empty:
    st.warning("Excel appears to be empty or could not be loaded.")
    st.stop()

# Find columns
service_col = find_col(raw, ["servicename", "service_name", "service"])
mobile_col = find_col(raw, ["mobile", "mobile_no", "mobile_number", "mobilenumber", "msisdn", "phone"])
cust_col = find_col(raw, ["tvamcustid", "tvam_customer_id", "customer_id", "cust_id"])
payref_col = find_col(raw, ["paymentrefno", "payment_ref_no", "payment_reference", "ref_no"])
utr_col = find_col(raw, ["utr", "utrno", "utr_no"])
date_col = find_col(raw, ["created_at", "createdon", "created_date", "date", "txn_date", "transaction_date", "timestamp"])

if not service_col:
    st.error("Column 'servicename' (or similar) not found in the Excel file.")
    st.stop()

# Add success column
df = add_success_column(raw)

# Show module counts
with st.expander("Module counts", expanded=False):
    counts = (
        df[service_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .value_counts(dropna=False)
        .rename_axis("servicename")
        .to_frame("count")
    )
    st.dataframe(counts, use_container_width=True)
    st.write(f"**Total records: {len(df):,}**")

# Module dropdown
all_modules = (
    df[service_col]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.upper()
    .sort_values()
    .unique()
    .tolist()
)

friendly_map = {f"{m} ({MODULE_LABELS.get(m, 'Other')})": m for m in all_modules}
friendly_labels = list(friendly_map.keys())
module_choice = st.selectbox("Select module", options=["All modules"] + friendly_labels)

selected_module = None
if module_choice != "All modules":
    selected_module = friendly_map[module_choice]

# Date filter
if date_col:
    series_dates = pd.to_datetime(df[date_col], errors="coerce")
    if series_dates.notna().any():
        min_dt, max_dt = series_dates.min().date(), series_dates.max().date()
        c1, c2 = st.columns(2)
        start_dt = c1.date_input("From date", value=min_dt)
        end_dt = c2.date_input("To date", value=max_dt)
    else:
        start_dt = end_dt = None
else:
    start_dt = end_dt = None

st.subheader("Search")
c1, c2, c3, c4 = st.columns(4)
q_mobile = c1.text_input("Mobile Number (contains)")
q_cust = c2.text_input("TvamCustId (contains)")
q_payref = c3.text_input("Payment Ref No (contains)")
q_utr = c4.text_input("UTR (contains)")

# Apply filters
filtered = df.copy()

if selected_module:
    filtered = filtered[
        filtered[service_col].fillna("").astype(str).str.strip().str.upper() == selected_module.strip().upper()
    ]

if start_dt and end_dt and date_col:
    ser = pd.to_datetime(filtered[date_col], errors="coerce")
    mask = ser.isna() | ser.dt.date.between(start_dt, end_dt)
    filtered = filtered[mask]

if mobile_col and q_mobile.strip():
    filtered = filtered[like_filter(filtered[mobile_col], q_mobile)]
if cust_col and q_cust.strip():
    filtered = filtered[like_filter(filtered[cust_col], q_cust)]
if payref_col and q_payref.strip():
    filtered = filtered[like_filter(filtered[payref_col], q_payref)]
if utr_col and q_utr.strip():
    filtered = filtered[like_filter(filtered[utr_col], q_utr)]

# KPIs
total_rows = len(filtered)
success_rows = int(filtered["__success__"].sum()) if "__success__" in filtered.columns else 0
fail_rows = total_rows - success_rows

k1, k2, k3 = st.columns(3)
k1.metric("Total Records", f"{total_rows:,}")
k2.metric("Success", f"{success_rows:,}")
k3.metric("Failures", f"{fail_rows:,}")

# Status chart
status_col_hint = find_col(
    filtered,
    ["status", "txn_status", "transaction_status", "payment_status", "invest_status", "investment_status"],
)
if status_col_hint:
    st.bar_chart(filtered[status_col_hint].astype(str).value_counts())

# Results table
st.markdown("### Results")
display_df = prepare_display(filtered)
st.dataframe(display_df, use_container_width=True, hide_index=True)

# Download buttons
c_dl1, c_dl2 = st.columns(2)
c_dl1.download_button(
    "Download CSV",
    data=display_df.to_csv(index=False).encode("utf-8"),
    file_name="filtered_results.csv",
    mime="text/csv",
)
c_dl2.download_button(
    "Download Excel",
    data=bytes_excel(display_df),
    file_name="filtered_results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# Missing column warnings
missing = []
if not mobile_col:
    missing.append("Mobile Number column not found.")
if not cust_col:
    missing.append("TvamCustId column not found.")
if not payref_col:
    missing.append("Payment Ref No column not found.")
if not utr_col:
    missing.append("UTR column not found.")
if missing:
    st.info(" ".join(missing))

st.caption("Counts are normalized by trimming spaces and uppercasing 'servicename'. Rows with blank dates are included by default.")
