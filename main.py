import streamlit as st
import pandas as pd
from io import BytesIO

# Page configuration
st.set_page_config(
    layout="wide",
    page_title="DAILY SMS SUMMARY CONSOLIDATOR",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# Apply dark mode and custom styling
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
    </style>
    """,
    unsafe_allow_html=True
)

# Title
st.title('DAILY SMS SUMMARY CONSOLIDATOR')

# Function to load and cache multiple Excel files
@st.cache_data
def load_daily_summaries(uploaded_files):
    dfs = []
    for file in uploaded_files:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()  # Clean column names
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

# Function to format numbers with commas for display
def format_with_commas(df, numeric_cols):
    df_copy = df.copy()
    for col in numeric_cols:
        if col in df_copy.columns:
            df_copy[col] = df_copy[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) else x)
    return df_copy

# Function to create Excel file
def to_excel_single(df, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_copy = df.copy()
        # Find and format date column
        date_col = next((col for col in df_copy.columns if 'date' in col.lower()), None)
        if date_col and date_col in df_copy.columns:
            if pd.api.types.is_datetime64_any_dtype(df_copy[date_col]):
                df_copy[date_col] = df_copy[date_col].dt.strftime('%d-%m-%Y')
            elif pd.api.types.is_object_dtype(df_copy[date_col]):
                df_copy[date_col] = pd.to_datetime(df_copy[date_col], errors='coerce').dt.strftime('%d-%m-%Y')
        
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

# Function to create overall summary
def create_overall_summary(df):
    # Find date column dynamically
    date_col = next((col for col in df.columns if 'date' in col.lower()), None)
    if not date_col:
        st.error("No date column found in the uploaded files. Expected a column with 'date' in its name.")
        st.write("Available columns:", list(df.columns))
        return None
    
    # Convert to datetime
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    
    # Get date range
    valid_dates = df[date_col].dropna()
    if len(valid_dates) == 0:
        st.error(f"No valid dates found in the '{date_col}' column")
        st.write("Sample values:", df[date_col].head().tolist())
        date_range_str = "Invalid Date Range"
    else:
        min_date = valid_dates.min()
        max_date = valid_dates.max()
        date_range_str = f"{min_date.strftime('%B %d')} - {max_date.strftime('%B %d, %Y')}"
    
    # Check for required columns
    required_cols = ['ENVIRONMENT', 'CLIENT', 'SMS SENDING', 'DELIVERED', 'FAILED']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"Missing required columns: {', '.join(missing_cols)}")
        st.write("Available columns:", list(df.columns))
        return None
    
    # Create overall summary
    overall_summary = df.groupby(['ENVIRONMENT', 'CLIENT']).agg({
        'SMS SENDING': 'sum',
        'DELIVERED': 'sum',
        'FAILED': 'sum'
    }).reset_index()
    
    overall_summary.insert(0, 'DATE_RANGE', date_range_str)
    return overall_summary

# Sidebar file uploader
with st.sidebar:
    st.subheader("Upload Daily Summary Files")
    uploaded_files = st.file_uploader(
        "Choose Daily Summary Excel files",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Select multiple Daily_SMS_Summary files to consolidate"
    )

# Main content
if uploaded_files:
    try:
        combined_df = load_daily_summaries(uploaded_files)
        
        overall_summary_df = create_overall_summary(combined_df)
        
        if overall_summary_df is not None:
            numeric_cols = ['SMS SENDING', 'DELIVERED', 'FAILED']
            overall_summary_display = format_with_commas(overall_summary_df, numeric_cols)
            
            st.subheader("Consolidated Overall SMS Summary")
            st.dataframe(overall_summary_display, use_container_width=True)
            
            excel_data = to_excel_single(overall_summary_df, "Consolidated_Overall_Summary")
            st.download_button(
                label="Download Consolidated Overall Summary",
                data=excel_data,
                file_name="consolidated_overall_sms_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred while processing the files: {str(e)}")
        st.write("Please check your input files and try again.")
else:
    st.info("Please upload one or more Daily_SMS_Summary Excel files to begin.")
