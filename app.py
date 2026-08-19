# Improved Streamlit app: more robust parsing, normalized columns, flexible inputs, better UI/UX, and safer caching.
import logging
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("crm_dashboard")

st.set_page_config(page_title="Deposit by RM Dashboard", layout="wide")

# Canonical column map: map many variants to a stable internal name
COLUMN_MAP = {
    # RM & center keys
    "RM_NAME": "rm_name",
    "Rm Name": "rm_name",
    "Rm_name": "rm_name",
    "rm_name": "rm_name",
    "PROCESSING_CENTER": "processing_center",
    "Processing Center": "processing_center",
    "processing_center": "processing_center",
    # Numeric columns
    "Baseline": "baseline",
    "BASELINE": "baseline",
    "baseline": "baseline",
    "Deposit Positional": "deposit_positional",
    "Deposit_Positional": "deposit_positional",
    "deposit_positional": "deposit_positional",
    "Incremental": "incremental",
    "INCREMENTAL": "incremental",
    "incremental": "incremental",
    "INCRIENTAL MOBILIZED": "incremental_mobilized",
    "INCRIMENTAL MOBILIZED": "incremental_mobilized",
    "INCRIENTAL_MOBILIZED": "incremental_mobilized",
    "INCRIMENTAL MOBILIZED": "incremental_mobilized",
    "INCRIMENTAL": "incremental",  # common typo
    # Percent / achievement
    "Achievment Percentage": "achievement_pct",
    "Achievement Percentage": "achievement_pct",
    "Achievment %": "achievement_pct",
    "ACHIEVMENT PERCENTAGE": "achievement_pct",
    "INCRIMENTAL PERCENTAGE": "incremental_pct",
    "INCREMENTAL PERCENTAGE": "incremental_pct",
    # Other detail fields
    "BRANCH_NAME": "branch_name",
    "Branch Name": "branch_name",
    "AC_DESC": "ac_desc",
    "AC Description": "ac_desc",
}


def canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip column names and map known variants to canonical names."""
    new_cols = {}
    for c in df.columns:
        if isinstance(c, str):
            s = c.strip()
            # Prefer exact match first
            if s in COLUMN_MAP:
                new_cols[c] = COLUMN_MAP[s]
            else:
                # Try case-insensitive match by upper-casing keys
                key_upper = s.upper()
                found = None
                for k, v in COLUMN_MAP.items():
                    if k.upper() == key_upper:
                        found = v
                        break
                if found:
                    new_cols[c] = found
                else:
                    # fallback: normalize spacing and lowercase
                    new_cols[c] = s.lower().replace(" ", "_")
        else:
            new_cols[c] = c
    return df.rename(columns=new_cols)


def find_section_start(sheet: pd.DataFrame, marker: str, search_col: int = 0) -> Optional[int]:
    """Return row index where marker occurs in the given search column; None if not found."""
    try:
        col = sheet.iloc[:, search_col].astype(str)
    except Exception:
        return None
    mask = col.str.contains(marker, na=False, case=False)
    matches = sheet[mask]
    if matches.empty:
        return None
    return int(matches.index[0])


def extract_table(sheet: pd.DataFrame, marker: str, header_offset: int = 1) -> pd.DataFrame:
    """
    Given sheet (read with header=None), find marker row, then header is marker+header_offset,
    table ends at next fully empty row (or end of sheet). Returns an empty DataFrame if marker not found.
    """
    start = find_section_start(sheet, marker)
    if start is None:
        logger.info("Marker '%s' not found in sheet.", marker)
        return pd.DataFrame()
    header_row = start + header_offset
    # find next empty row
    end = len(sheet)
    for i in range(header_row, len(sheet)):
        if sheet.iloc[i].isnull().all():
            end = i
            break
    # slice and set columns
    table = sheet.iloc[header_row:end].copy()
    if table.empty:
        return pd.DataFrame()
    # first row often contains column names; try to set them
    first_row = table.iloc[0].astype(str).str.strip()
    table.columns = first_row
    table = table.iloc[1:].reset_index(drop=True)
    table = canonicalize_columns(table)
    # Drop rows where main identifier is null (processing_center / rm_name)
    if "processing_center" in table.columns:
        table = table[table["processing_center"].notna()]
    elif "rm_name" in table.columns:
        table = table[table["rm_name"].notna()]
    return table


def coerce_numeric(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


@st.cache_data(show_spinner=False)
def load_data_from_bytes(content: bytes, filename: str) -> Dict:
    """
    Parse a single Excel file from bytes (for uploaded files) or path content.
    Returns a dict with date and dataframes.
    """
    # use BytesIO for uploaded files
    try:
        excel_bytes = BytesIO(content)
        # Read first sheet without header so we can detect sections by searching first column
        df_sheet1 = pd.read_excel(excel_bytes, sheet_name=0, header=None, engine="openpyxl")
    except Exception as e:
        logger.exception("Failed to read sheet1 from %s: %s", filename, e)
        return {"date": None, "processing": pd.DataFrame(), "crm": pd.DataFrame(), "detail": pd.DataFrame(), "transactions": pd.DataFrame()}

    # Extract date from filename if possible
    report_date = None
    try:
        m = pd.to_datetime(pd.Series([filename]).str.extract(r"(\d{4}-\d{2}-\d{2})")[0], errors="coerce").iloc[0]
        if pd.notna(m):
            report_date = m.date()
    except Exception:
        report_date = None

    # Extract sections
    processing = extract_table(df_sheet1, "Processing Center Summary")
    crm = extract_table(df_sheet1, "CRM Summary")
    detail = extract_table(df_sheet1, "Proc. Center & CRM Summary")

    # Try reading transactions: try "Sheet2" name first, then sheet index 1
    transactions = pd.DataFrame()
    try:
        excel_bytes.seek(0)
        transactions = pd.read_excel(excel_bytes, sheet_name="Sheet2", header=4, engine="openpyxl")
    except Exception:
        try:
            excel_bytes.seek(0)
            transactions = pd.read_excel(excel_bytes, sheet_name=1, header=4, engine="openpyxl")
        except Exception:
            # no sheet2 found - leave empty
            transactions = pd.DataFrame()

    # Normalize transaction columns if present
    if not transactions.empty:
        transactions = canonicalize_columns(transactions)
        transactions = coerce_numeric(transactions, ["baseline", "deposit_positional", "incremental_mobilized", "incremental", "achievement_pct", "incremental_pct"])

    # Coerce numeric columns in summary tables
    for df in (processing, crm, detail):
        if not df.empty:
            df = coerce_numeric(df, ["baseline", "deposit_positional", "incremental", "achievement_pct"])

    return {"date": report_date, "processing": processing, "crm": crm, "detail": detail, "transactions": transactions}


def load_files_from_glob(pattern: str) -> List[Tuple[bytes, str]]:
    """
    Read files matching a glob pattern in the current working directory.
    Returns list of (bytes, filename).
    """
    files = []
    for p in sorted(Path(".").glob(pattern)):
        if p.is_file() and p.suffix in {".xlsx", ".xls"}:
            try:
                files.append((p.read_bytes(), p.name))
            except Exception as e:
                logger.warning("Could not read file %s: %s", p, e)
    return files


# -------------------- UI: Inputs --------------------
st.title("💼 Deposit by RM - Executive Dashboard")
st.markdown("Analyzing daily deposit performance across processing centers and relationship managers.")

st.sidebar.header("Data source")

data_source = st.sidebar.radio("Choose how to provide data", ("Upload Excel files", "Use glob pattern (local)"))

uploaded_files: List[Tuple[bytes, str]] = []
if data_source == "Upload Excel files":
    uploaded = st.sidebar.file_uploader("Upload one or more Excel files (2026-08-*.xlsx)", type=["xlsx", "xls"], accept_multiple_files=True)
    if uploaded:
        for f in uploaded:
            try:
                content = f.read()
                uploaded_files.append((content, f.name))
            except Exception as e:
                st.sidebar.error(f"Failed to read uploaded file {f.name}: {e}")
else:
    pattern = st.sidebar.text_input("Glob pattern (relative to app directory)", value="2026-08-*.xlsx")
    if st.sidebar.button("Load from glob"):
        uploaded_files = load_files_from_glob(pattern)
        if not uploaded_files:
            st.sidebar.warning("No files found for pattern: " + pattern)

if not uploaded_files:
    st.info("No Excel files provided yet. Upload files or use a glob pattern in the sidebar.")
    st.stop()

# -------------------- Parse Files --------------------
data_dict = {}
for content, filename in uploaded_files:
    parsed = load_data_from_bytes(content, filename)
    if parsed and parsed["date"] is not None:
        data_dict[parsed["date"]] = parsed
    else:
        # If no date extracted, try to let user choose date later; still include under filename key
        # We attach with a unique key if date missing
        key = filename
        count = 1
        while key in data_dict:
            key = f"{filename}-{count}"
            count += 1
        data_dict[key] = parsed

if not data_dict:
    st.error("No valid data parsed from files.")
    st.stop()

# -------------------- Combine data --------------------
all_processing = []
all_crm = []
for date_key, data in data_dict.items():
    proc = data["processing"].copy() if not data["processing"].empty else pd.DataFrame()
    crm = data["crm"].copy() if not data["crm"].empty else pd.DataFrame()
    # if parsed date is a date object, add as Date; otherwise try to parse from date_key
    if isinstance(date_key, (str,)) and not isinstance(date_key, datetime):
        # key could be filename; try parse date inside parsed['date']
        date_value = data.get("date")
    else:
        date_value = date_key
    if not proc.empty:
        proc["Date"] = date_value
        all_processing.append(proc)
    if not crm.empty:
        crm["Date"] = date_value
        all_crm.append(crm)

if all_processing:
    df_proc_all = pd.concat(all_processing, ignore_index=True)
else:
    df_proc_all = pd.DataFrame()

if all_crm:
    df_crm_all = pd.concat(all_crm, ignore_index=True)
else:
    df_crm_all = pd.DataFrame()

# If no CRM summary rows at all, stop with helpful message
if df_crm_all.empty:
    st.error("No CRM summary rows found in uploaded files. Check file formatting.")
    st.stop()

# Ensure Date column is datetime.date or pd.Timestamp
if "Date" in df_crm_all.columns:
    df_crm_all["Date"] = pd.to_datetime(df_crm_all["Date"], errors="coerce").dt.date

if "Date" in df_proc_all.columns:
    df_proc_all["Date"] = pd.to_datetime(df_proc_all["Date"], errors="coerce").dt.date

# Sidebar filters: dynamic based on parsed content
st.sidebar.header("Filters")
dates_sorted = sorted([d for d in df_crm_all["Date"].unique() if pd.notna(d)])
# date selection: date_range or multiselect if limited
if dates_sorted:
    selected_dates = st.sidebar.multiselect("Select Dates", options=dates_sorted, default=dates_sorted)
else:
    selected_dates = []

centers = sorted(df_proc_all["processing_center"].dropna().unique()) if not df_proc_all.empty else []
selected_centers = st.sidebar.multiselect("Processing Center", options=centers, default=centers)

# Filtered frames
filtered_crm = df_crm_all.copy()
if selected_dates:
    filtered_crm = filtered_crm[filtered_crm["Date"].isin(selected_dates)]
filtered_proc = df_proc_all.copy()
if selected_dates:
    filtered_proc = filtered_proc[filtered_proc["Date"].isin(selected_dates)]
if selected_centers and not filtered_proc.empty:
    filtered_proc = filtered_proc[filtered_proc["processing_center"].isin(selected_centers)]

# Ensure numeric columns exist and are numeric
for col in ["baseline", "deposit_positional", "incremental", "achievement_pct"]:
    if col in filtered_crm.columns:
        filtered_crm[col] = pd.to_numeric(filtered_crm[col], errors="coerce")
    if col in filtered_proc.columns:
        filtered_proc[col] = pd.to_numeric(filtered_proc[col], errors="coerce")

# ---------- KPIs ----------
col1, col2, col3, col4 = st.columns(4)
total_deposits = filtered_crm["deposit_positional"].sum(min_count=1)
total_baseline = filtered_crm["baseline"].sum(min_count=1)
total_inc = filtered_crm["incremental"].sum(min_count=1)
avg_achievement = filtered_crm["achievement_pct"].mean()

# Compute delta between earliest and latest selected dates (absolute and percent)
if selected_dates and len(selected_dates) >= 2:
    first_date = min(selected_dates)
    last_date = max(selected_dates)
    first_dep = df_crm_all[df_crm_all["Date"] == first_date]["deposit_positional"].sum(min_count=1)
    last_dep = df_crm_all[df_crm_all["Date"] == last_date]["deposit_positional"].sum(min_count=1)
    daily_change_abs = (last_dep or 0) - (first_dep or 0)
    daily_change_pct = (daily_change_abs / (first_dep or 1)) if (first_dep and first_dep != 0) else None
else:
    daily_change_abs = None
    daily_change_pct = None

delta_str = None
if daily_change_abs is not None:
    if daily_change_pct is not None:
        delta_str = f"{daily_change_abs:+,.2f} ({daily_change_pct:+.1%})"
    else:
        delta_str = f"{daily_change_abs:+,.2f}"

col1.metric("Total Deposits", f"{total_deposits:,.2f}", delta=delta_str)
col2.metric("Baseline", f"{total_baseline:,.2f}")
col3.metric("Incremental", f"{total_inc:+,.2f}")
col4.metric("Avg Achievement %", f"{avg_achievement:.1%}" if pd.notna(avg_achievement) else "N/A")

# ---------- Time Series ----------
st.subheader("📈 Deposit Trend Over Time")
ts_data = filtered_crm.groupby("Date", dropna=False)["deposit_positional"].sum(min_count=1).reset_index()
if not ts_data.empty:
    fig_ts = px.line(ts_data.sort_values("Date"), x="Date", y="deposit_positional", title="Total Deposits by Date", markers=True)
    fig_ts.update_traces(hovertemplate="%{x}: %{y:,.0f}")
    fig_ts.update_layout(yaxis_tickformat=",")
    st.plotly_chart(fig_ts, use_container_width=True)
else:
    st.info("No time series data for selected dates/filters.")

# ---------- Processing Center Breakdown ----------
st.subheader("🏢 Processing Center Performance")
if not filtered_proc.empty:
    center_data = filtered_proc.groupby("processing_center", dropna=False).agg(
        baseline=pd.NamedAgg(column="baseline", aggfunc="sum"),
        deposit_positional=pd.NamedAgg(column="deposit_positional", aggfunc="sum"),
        incremental=pd.NamedAgg(column="incremental", aggfunc="sum"),
    ).reset_index()

    fig_center = go.Figure()
    fig_center.add_trace(go.Bar(x=center_data["processing_center"], y=center_data["deposit_positional"],
                                 name="Deposit Positional", marker_color="royalblue", hovertemplate="%{x}: %{y:,.0f}"))
    fig_center.add_trace(go.Bar(x=center_data["processing_center"], y=center_data["baseline"],
                                 name="Baseline", marker_color="lightgray", hovertemplate="%{x}: %{y:,.0f}"))
    fig_center.update_layout(barmode="group", yaxis_tickformat=",", title="Deposits by Processing Center")
    st.plotly_chart(fig_center, use_container_width=True)

    fig_inc = px.bar(center_data, x="processing_center", y="incremental", title="Incremental by Processing Center", color="incremental", color_continuous_scale="RdBu")
    fig_inc.update_layout(yaxis_tickformat=",")
    st.plotly_chart(fig_inc, use_container_width=True)
else:
    st.info("No processing center data available for selected filters.")

# ---------- RM Performance Table ----------
st.subheader("👤 Relationship Manager Performance")
display_cols = ["Date", "rm_name", "baseline", "deposit_positional", "incremental", "achievement_pct"]
rm_display = filtered_crm.copy()
# Format achievement column for display without mutating numeric columns used for charts
if "achievement_pct" in rm_display.columns:
    rm_display["Achievement %"] = rm_display["achievement_pct"].apply(lambda x: f"{x:.1%}" if pd.notna(x) else "N/A")
else:
    rm_display["Achievement %"] = "N/A"
rm_display = rm_display.rename(columns={"rm_name": "RM Name", "baseline": "Baseline", "deposit_positional": "Deposit Positional", "incremental": "Incremental"})
# Sort by incremental if present
if "Incremental" in rm_display.columns:
    rm_display = rm_display.sort_values("Incremental", ascending=False)

# Use st.dataframe; column_config requires Streamlit >= 1.18; fall back if not available
try:
    st.dataframe(rm_display[["Date", "RM Name", "Baseline", "Deposit Positional", "Incremental", "Achievement %"]], use_container_width=True, hide_index=True)
except Exception:
    st.write(rm_display[["Date", "RM Name", "Baseline", "Deposit Positional", "Incremental", "Achievement %"]])

# ---------- Scatter Plot ----------
st.subheader("📊 Incremental vs Achievement %")
if {"incremental", "achievement_pct"}.issubset(filtered_crm.columns):
    scatter_data = filtered_crm.dropna(subset=["incremental", "achievement_pct"])
    if not scatter_data.empty:
        fig_scatter = px.scatter(scatter_data, x="incremental", y="achievement_pct", color="Date", hover_data=["rm_name"], title="Incremental vs Achievement Percentage (by RM)", labels={"incremental": "Incremental (ETB)", "achievement_pct": "Achievement %"})
        fig_scatter.update_layout(yaxis_tickformat=".0%")
        st.plotly_chart(fig_scatter, use_container_width=True)
    else:
        st.info("No data available for scatter plot.")
else:
    st.info("Required columns for scatter plot are missing.")

# ---------- Drill-down: RM Detailed Transactions ----------
st.subheader("🔍 Drill-down: RM Detailed Transactions")
rm_list = sorted(filtered_crm["rm_name"].dropna().unique())
selected_rm = st.selectbox("Select RM to view detailed transactions", options=["-- select --"] + rm_list)
if selected_rm and selected_rm != "-- select --":
    trans_list = []
    for key, data in data_dict.items():
        trans = data.get("transactions", pd.DataFrame())
        if not trans.empty and "rm_name" in trans.columns:
            trans_rm = trans[trans["rm_name"] == selected_rm].copy()
            if not trans_rm.empty:
                # attach Date from parsed stream if available (data['date'])
                date_val = data.get("date")
                trans_rm["Date"] = date_val
                trans_list.append(trans_rm)
    if trans_list:
        trans_combined = pd.concat(trans_list, ignore_index=True)
        total_dep = trans_combined.get("deposit_positional", pd.Series(dtype=float)).sum(min_count=1)
        total_incr_mobilized = trans_combined.get("incremental_mobilized", pd.Series(dtype=float)).sum(min_count=1)
        st.write(f"**{selected_rm}** – Total Deposits: {total_dep:,.2f}, Incremental Mobilized: {total_incr_mobilized:+,.2f}")
        display_cols = ["Date", "processing_center", "branch_name", "ac_desc", "baseline", "deposit_positional", "incremental_mobilized"]
        available_cols = [c for c in display_cols if c in trans_combined.columns]
        st.dataframe(trans_combined[available_cols], use_container_width=True, hide_index=True)
    else:
        st.warning("No detailed transactions found for this RM.")

# ---------- Export ----------
st.sidebar.markdown("---")
st.sidebar.subheader("Export Data")
if st.sidebar.button("Prepare Filtered CSV"):
    csv = filtered_crm.to_csv(index=False).encode("utf-8")
    fn = f"deposit_summary_{datetime.now().strftime('%Y%m%d')}.csv"
    st.sidebar.download_button(label="Download CSV", data=csv, file_name=fn, mime="text/csv")

st.markdown("---")
st.caption("Dashboard built with Streamlit and Plotly. Data parsing is more robust and column names are normalized for consistent analysis.")
