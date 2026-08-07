import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import glob
import re

# Set page config
st.set_page_config(page_title="Deposit by RM Dashboard", layout="wide")

# -------------------- Data Loading and Parsing --------------------
@st.cache_data
def load_data(file_path):
    """
    Load and parse the Excel file.
    Returns: dict with dataframes: processing_summary, crm_summary, detailed, date
    """
    try:
        # Extract date from filename (e.g., 2026-08-05.xlsx)
        filename = file_path.split('/')[-1]
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
        report_date = datetime.strptime(date_match.group(1), '%Y-%m-%d').date() if date_match else None
        
        # Read Sheet1 (summary)
        df_sheet1 = pd.read_excel(file_path, sheet_name=0, header=None)
        
        # Find "Processing Center Summary"
        proc_start = df_sheet1[df_sheet1[0].astype(str).str.contains("Processing Center Summary", na=False)].index[0]
        proc_header = proc_start + 1
        proc_end = df_sheet1[proc_header+1:][df_sheet1[proc_header+1:][0].isnull().all(axis=1)].index[0]
        proc_df = df_sheet1.iloc[proc_header:proc_end].copy()
        proc_df.columns = proc_df.iloc[0]
        proc_df = proc_df[1:].reset_index(drop=True)
        proc_df = proc_df[['PROCESSING_CENTER', 'Baseline', 'Deposit Positional', 'Incremental', 'Achievment Percentage']]
        proc_df = proc_df[proc_df['PROCESSING_CENTER'].notna()]
        
        # Find "CRM Summary"
        crm_start = df_sheet1[df_sheet1[0].astype(str).str.contains("CRM Summary", na=False)].index[0]
        crm_header = crm_start + 1
        crm_end = df_sheet1[crm_header+1:][df_sheet1[crm_header+1:][0].isnull().all(axis=1)].index[0]
        crm_df = df_sheet1.iloc[crm_header:crm_end].copy()
        crm_df.columns = crm_df.iloc[0]
        crm_df = crm_df[1:].reset_index(drop=True)
        crm_df = crm_df[['RM_NAME', 'Baseline', 'Deposit Positional', 'Incremental', 'Achievment Percentage']]
        crm_df = crm_df[crm_df['RM_NAME'].notna()]
        
        # Find "Proc. Center & CRM Summary" - detailed breakdown
        detail_start = df_sheet1[df_sheet1[0].astype(str).str.contains("Proc. Center & CRM Summary", na=False)].index[0]
        detail_header = detail_start + 1
        detail_end = df_sheet1[detail_header+1:][df_sheet1[detail_header+1:][0].isnull().all(axis=1)].index[0]
        detail_df = df_sheet1.iloc[detail_header:detail_end].copy()
        detail_df.columns = detail_df.iloc[0]
        detail_df = detail_df[1:].reset_index(drop=True)
        detail_df = detail_df[['PROCESSING_CENTER', 'RM_NAME', 'Baseline', 'Deposit Positional', 'Incremental']]
        detail_df = detail_df[detail_df['PROCESSING_CENTER'].notna()]
        
        # Sheet2: Detailed transactions
        df_sheet2 = pd.read_excel(file_path, sheet_name=1, header=4)
        df_sheet2.columns = df_sheet2.iloc[0]
        df_sheet2 = df_sheet2[1:].reset_index(drop=True)
        df_sheet2 = df_sheet2[df_sheet2['RM_NAME'].notna()]
        
        # Convert numeric columns
        for col in ['Baseline', 'Deposit Positional', 'INCRIENTAL MOBILIZED', 'INCRIMENTAL PERCENTAGE']:
            if col in df_sheet2.columns:
                df_sheet2[col] = pd.to_numeric(df_sheet2[col], errors='coerce')
        
        return {
            'date': report_date,
            'processing': proc_df,
            'crm': crm_df,
            'detail': detail_df,
            'transactions': df_sheet2
        }
    except Exception as e:
        st.warning(f"Could not load file {file_path}: {e}")
        return None

# Load all Excel files in the current directory matching pattern "2026-08-*.xlsx"
file_pattern = "2026-08-*.xlsx"
file_list = glob.glob(file_pattern)
if not file_list:
    st.error("No Excel files found matching pattern. Please ensure files are in the current directory.")
    st.stop()

data_dict = {}
for file_path in file_list:
    data = load_data(file_path)
    if data and data['date']:
        data_dict[data['date']] = data

if not data_dict:
    st.error("No valid data loaded. Please check file format.")
    st.stop()

# Combine data across dates into a single dataframe for time series
all_processing = []
all_crm = []
for date, data in data_dict.items():
    proc = data['processing'].copy()
    proc['Date'] = date
    all_processing.append(proc)
    
    crm = data['crm'].copy()
    crm['Date'] = date
    all_crm.append(crm)

df_proc_all = pd.concat(all_processing, ignore_index=True)
df_crm_all = pd.concat(all_crm, ignore_index=True)

# Convert numeric columns
for col in ['Baseline', 'Deposit Positional', 'Incremental', 'Achievment Percentage']:
    if col in df_proc_all.columns:
        df_proc_all[col] = pd.to_numeric(df_proc_all[col], errors='coerce')
    if col in df_crm_all.columns:
        df_crm_all[col] = pd.to_numeric(df_crm_all[col], errors='coerce')

# -------------------- Dashboard UI --------------------
st.title("💼 Deposit by RM - Executive Dashboard")
st.markdown("Analyzing daily deposit performance across processing centers and relationship managers.")

# Sidebar filters
st.sidebar.header("Filters")
selected_dates = st.sidebar.multiselect(
    "Select Dates",
    options=sorted(df_crm_all['Date'].unique()),
    default=sorted(df_crm_all['Date'].unique())
)

selected_centers = st.sidebar.multiselect(
    "Processing Center",
    options=sorted(df_proc_all['PROCESSING_CENTER'].unique()),
    default=sorted(df_proc_all['PROCESSING_CENTER'].unique())
)

# Filter data
filtered_crm = df_crm_all[df_crm_all['Date'].isin(selected_dates)]
filtered_proc = df_proc_all[df_proc_all['Date'].isin(selected_dates)]
filtered_proc = filtered_proc[filtered_proc['PROCESSING_CENTER'].isin(selected_centers)]

# ---------- KPIs ----------
col1, col2, col3, col4 = st.columns(4)

total_deposits = filtered_crm['Deposit Positional'].sum()
total_baseline = filtered_crm['Baseline'].sum()
total_inc = filtered_crm['Incremental'].sum()
avg_achievement = filtered_crm['Achievment Percentage'].mean()

# For daily change, compare first and last selected dates
if len(selected_dates) >= 2:
    first_date = min(selected_dates)
    last_date = max(selected_dates)
    first_dep = df_crm_all[df_crm_all['Date']==first_date]['Deposit Positional'].sum()
    last_dep = df_crm_all[df_crm_all['Date']==last_date]['Deposit Positional'].sum()
    daily_change = last_dep - first_dep
else:
    daily_change = 0

col1.metric("Total Deposits", f"{total_deposits:,.2f}", delta=f"{daily_change:+,.2f}")
col2.metric("Baseline", f"{total_baseline:,.2f}")
col3.metric("Incremental", f"{total_inc:+,.2f}")
col4.metric("Avg Achievement %", f"{avg_achievement:.2%}" if pd.notna(avg_achievement) else "N/A")

# ---------- Time Series ----------
st.subheader("📈 Deposit Trend Over Time")
ts_data = filtered_crm.groupby('Date')['Deposit Positional'].sum().reset_index()
fig_ts = px.line(ts_data, x='Date', y='Deposit Positional', 
                 title='Total Deposits by Date', markers=True)
fig_ts.update_layout(yaxis_tickformat=',.0f')
st.plotly_chart(fig_ts, use_container_width=True)

# ---------- Processing Center Breakdown ----------
st.subheader("🏢 Processing Center Performance")
center_data = filtered_proc.groupby('PROCESSING_CENTER').agg({
    'Baseline': 'sum',
    'Deposit Positional': 'sum',
    'Incremental': 'sum'
}).reset_index()

fig_center = go.Figure()
fig_center.add_trace(go.Bar(
    x=center_data['PROCESSING_CENTER'],
    y=center_data['Deposit Positional'],
    name='Deposit Positional',
    marker_color='royalblue'
))
fig_center.add_trace(go.Bar(
    x=center_data['PROCESSING_CENTER'],
    y=center_data['Baseline'],
    name='Baseline',
    marker_color='lightgray'
))
fig_center.update_layout(barmode='group', yaxis_tickformat=',.0f',
                         title='Deposits by Processing Center')
st.plotly_chart(fig_center, use_container_width=True)

fig_inc = px.bar(center_data, x='PROCESSING_CENTER', y='Incremental',
                 title='Incremental by Processing Center',
                 color='Incremental', color_continuous_scale='RdBu')
fig_inc.update_layout(yaxis_tickformat=',.0f')
st.plotly_chart(fig_inc, use_container_width=True)

# ---------- RM Performance Table ----------
st.subheader("👤 Relationship Manager Performance")
rm_display = filtered_crm[['Date', 'RM_NAME', 'Baseline', 'Deposit Positional', 'Incremental', 'Achievment Percentage']]
rm_display['Achievement %'] = rm_display['Achievment Percentage'].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "N/A")
rm_display = rm_display.sort_values('Incremental', ascending=False)

st.dataframe(
    rm_display,
    column_config={
        "Date": st.column_config.DateColumn("Date"),
        "RM_NAME": "RM Name",
        "Baseline": st.column_config.NumberColumn("Baseline", format="%.2f"),
        "Deposit Positional": st.column_config.NumberColumn("Deposit Positional", format="%.2f"),
        "Incremental": st.column_config.NumberColumn("Incremental", format="%.2f"),
        "Achievement %": st.column_config.TextColumn("Achievement %"),
    },
    use_container_width=True,
    hide_index=True
)

# ---------- Scatter Plot ----------
st.subheader("📊 Incremental vs Achievement %")
scatter_data = filtered_crm.dropna(subset=['Incremental', 'Achievment Percentage'])
if not scatter_data.empty:
    fig_scatter = px.scatter(
        scatter_data, x='Incremental', y='Achievment Percentage',
        hover_data=['RM_NAME', 'Date'],
        color='Date',
        title='Incremental vs Achievement Percentage (by RM)',
        labels={'Incremental': 'Incremental (ETB)', 'Achievment Percentage': 'Achievement %'}
    )
    fig_scatter.update_layout(yaxis_tickformat='.0%')
    st.plotly_chart(fig_scatter, use_container_width=True)
else:
    st.info("No data available for scatter plot.")

# ---------- Drill-down ----------
st.subheader("🔍 Drill-down: RM Detailed Transactions")
rm_list = sorted(df_crm_all['RM_NAME'].unique())
selected_rm = st.selectbox("Select RM to view detailed transactions", rm_list)

if selected_rm:
    trans_list = []
    for date, data in data_dict.items():
        trans = data['transactions']
        trans_rm = trans[trans['RM_NAME'] == selected_rm].copy()
        trans_rm['Date'] = date
        trans_list.append(trans_rm)
    
    if trans_list:
        trans_combined = pd.concat(trans_list, ignore_index=True)
        st.write(f"**{selected_rm}** - Total Deposits: {trans_combined['Deposit Positional'].sum():,.2f}, Incremental: {trans_combined['INCRIENTAL MOBILIZED'].sum():+,.2f}")
        
        st.dataframe(
            trans_combined[['Date', 'PROCESSING_CENTER', 'BRANCH_NAME', 'AC_DESC', 'Baseline', 'Deposit Positional', 'INCRIENTAL MOBILIZED']],
            column_config={
                "Date": st.column_config.DateColumn("Date"),
                "PROCESSING_CENTER": "Processing Center",
                "BRANCH_NAME": "Branch",
                "AC_DESC": "Account Description",
                "Baseline": st.column_config.NumberColumn("Baseline", format="%.2f"),
                "Deposit Positional": st.column_config.NumberColumn("Deposit Positional", format="%.2f"),
                "INCRIENTAL MOBILIZED": st.column_config.NumberColumn("Incremental", format="%.2f"),
            },
            use_container_width=True,
            hide_index=True
        )
    else:
        st.warning("No detailed transactions found for this RM.")

# ---------- Export ----------
st.sidebar.markdown("---")
st.sidebar.subheader("Export Data")
if st.sidebar.button("Download Filtered Data (CSV)"):
    csv = filtered_crm.to_csv(index=False).encode('utf-8')
    st.sidebar.download_button(
        label="Download CSV",
        data=csv,
        file_name=f"deposit_summary_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

st.markdown("---")
st.caption("Dashboard built with Streamlit and Plotly. Data updated daily.")
