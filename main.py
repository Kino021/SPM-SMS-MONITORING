import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    layout="wide",
    page_title="SMS SUMMARY AGGREGATOR",
    page_icon="📈",
    initial_sidebar_state="expanded"
)

# Custom styling
st.markdown(
    """
    <style>
    .reportview-container {
        background: #2E2E2E;
        color: white;
        max-width: 100%;
        padding: 1rem;
    }
    .sidebar .sidebar-content {
        background: #2E2E2E;
        width: 300px;
    }
    h1, h2, h3 {
        color: #87CEEB !important;
        font-weight: bold !important;
    }
    .stDataFrame {
        width: 100% !important;
        font-size: 14px;
    }
    .stSelectbox, .stDateInput {
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Title
st.title('SMS SUMMARY AGGREGATOR')

# Data loading function
@st.cache_data
def load_summaries(uploaded_files):
    dfs = []
    for file in uploaded_files:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

# Number formatting function
def format_with_commas(df, numeric_cols):
    df_copy = df.copy()
    for col in numeric_cols:
        if col in df_copy.columns:
            df_copy[col] = df_copy[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) else x)
    return df_copy

# Excel export function
def to_excel_single(df, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_copy = df.copy()
        if 'DATE' in df_copy.columns:
            df_copy['DATE'] = df_copy['DATE'].dt.strftime('%d-%m-%Y')
        
        df_copy.to_excel(writer, index=False, sheet_name=sheet_name)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#87CEEB',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1, 'num_format': '#,##0'})
        text_format = workbook.add_format({'border': 1})
        
        for col_num, value in enumerate(df_copy.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        numeric_cols = ['SMS SENDING', 'DELIVERED', 'FAILED']
        for row_num in range(1, len(df_copy) + 1):
            for col_num, col_name in enumerate(df_copy.columns):
                value = df_copy.iloc[row_num-1, col_num]
                format_to_use = cell_format if col_name in numeric_cols else text_format
                worksheet.write(row_num, col_num, value, format_to_use)
        
        for i, col in enumerate(df_copy.columns):
            max_length = max(df_copy[col].astype(str).map(len).max(), len(str(col)))
            worksheet.set_column(i, i, max_length + 2)
    
    return output.getvalue()

# Summary creation function
def create_monthly_summary(df):
    df['MONTH_YEAR'] = df['DATE'].dt.strftime('%B %Y')
    
    valid_dates = df['DATE'].dropna()
    date_range_str = (f"{valid_dates.min().strftime('%B %d, %Y')} - "
                     f"{valid_dates.max().strftime('%B %d, %Y')}" if len(valid_dates) > 0 
                     else "Invalid Date Range")
    
    monthly_summary = df.groupby(['MONTH_YEAR', 'ENVIRONMENT', 'CLIENT']).agg({
        'SMS SENDING': 'sum',
        'DELIVERED': 'sum',
        'FAILED': 'sum'
    }).reset_index()
    
    monthly_summary.insert(0, 'DATE_RANGE', date_range_str)
    return monthly_summary

# Sidebar
with st.sidebar:
    st.subheader("Upload & Filter Options")
    
    uploaded_files = st.file_uploader(
        "Upload Daily Summary Files",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Select multiple Daily_SMS_Summary files"
    )
    
    if uploaded_files:
        df = load_summaries(uploaded_files)
        
        # Filter options
        st.subheader("Filters")
        environments = ['All'] + sorted(df['ENVIRONMENT'].unique().tolist())
        selected_env = st.selectbox("Select Environment", environments)
        
        clients = ['All'] + sorted(df['CLIENT'].unique().tolist())
        selected_client = st.selectbox("Select Client", clients)
        
        min_date = df['DATE'].min().date()
        max_date = df['DATE'].max().date()
        date_range = st.date_input(
            "Select Date Range",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )

# Main content
if uploaded_files:
    df = load_summaries(uploaded_files)
    
    # Apply filters
    filtered_df = df.copy()
    if selected_env != 'All':
        filtered_df = filtered_df[filtered_df['ENVIRONMENT'] == selected_env]
    if selected_client != 'All':
        filtered_df = filtered_df[filtered_df['CLIENT'] == selected_client]
    if len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['DATE'].dt.date >= start_date) & 
            (filtered_df['DATE'].dt.date <= end_date)
        ]
    
    # Create summary
    monthly_summary_df = create_monthly_summary(filtered_df)
    
    # Display
    numeric_cols = ['SMS SENDING', 'DELIVERED', 'FAILED']
    display_df = format_with_commas(monthly_summary_df, numeric_cols)
    
    st.subheader("Monthly SMS Summary")
    st.dataframe(display_df, use_container_width=True)
    
    # Download button
    excel_data = to_excel_single(monthly_summary_df, "Monthly_SMS_Summary")
    st.download_button(
        label="Download Monthly Summary",
        data=excel_data,
        file_name=f"monthly_sms_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # Show some basic stats
    with st.expander("Quick Statistics"):
        total_sms = monthly_summary_df['SMS SENDING'].sum()
        total_delivered = monthly_summary_df['DELIVERED'].sum()
        total_failed = monthly_summary_df['FAILED'].sum()
        delivery_rate = (total_delivered / total_sms * 100) if total_sms > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total SMS", f"{total_sms:,}")
        col2.metric("Delivery Rate", f"{delivery_rate:.1f}%")
        col3.metric("Failed SMS", f"{total_failed:,}")
else:
    st.info("Please upload Daily_SMS_Summary Excel files to begin.")
