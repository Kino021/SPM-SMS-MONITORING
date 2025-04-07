import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide", page_title="DIALER PRODUCTIVITY PER CRITERIA OF BALANCE", page_icon="📊", initial_sidebar_state="expanded")

# Apply dark mode and custom styling (unchanged)
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
    .stExpander {
        width: 100%;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title('SPM SMS MONITORING ALL ENVI')

@st.cache_data
def load_data(uploaded_files):
    dfs = []
    for file in uploaded_files:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        df['Source_File'] = file.name
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def format_with_commas(df, numeric_cols):
    df_copy = df.copy()
    for col in numeric_cols:
        if col in df_copy.columns:
            df_copy[col] = df_copy[col].apply(lambda x: f"{x:,.0f}" if pd.notnull(x) else x)
    return df_copy

def to_excel_single(df, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_copy = df.copy()
        if 'DATE_RANGE' in df_copy.columns:
            df_copy['DATE_RANGE'] = df_copy['DATE_RANGE'].astype(str)
        elif 'DATE' in df_copy.columns:
            if pd.api.types.is_datetime64_any_dtype(df_copy['DATE']):
                df_copy['DATE'] = df_copy['DATE'].dt.strftime('%d-%m-%Y')
            elif pd.api.types.is_object_dtype(df_copy['DATE']):
                df_copy['DATE'] = pd.to_datetime(df_copy['DATE'], errors='coerce').dt.strftime('%d-%m-%Y')
        df_copy.to_excel(writer, index=False, sheet_name=sheet_name)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        header_format = workbook.add_format({'bold': True, 'bg_color': '#87CEEB', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
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

def to_excel_multiple(dfs_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        header_format = workbook.add_format({'bold': True, 'bg_color': '#87CEEB', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        cell_format = workbook.add_format({'border': 1, 'num_format': '#,##0'})
        text_format = workbook.add_format({'border': 1})
        
        for sheet_name, df in dfs_dict.items():
            df_copy = df.copy()
            if 'DATE_RANGE' in df_copy.columns:
                df_copy['DATE_RANGE'] = df_copy['DATE_RANGE'].astype(str)
            elif 'DATE' in df_copy.columns:
                if pd.api.types.is_datetime64_any_dtype(df_copy['DATE']):
                    df_copy['DATE'] = df_copy['DATE'].dt.strftime('%d-%m-%Y')
                elif pd.api.types.is_object_dtype(df_copy['DATE']):
                    df_copy['DATE'] = pd.to_datetime(df_copy['DATE'], errors='coerce').dt.strftime('%d-%m-%Y')
            
            df_copy.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            
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

def create_sms_summaries(df):
    required_columns = {
        'date_col': 'SMS Status Response Date/Time',  # Changed from 'Submission Date / Time'
        'env_col': 'Environment',
        'client_col': 'Client',
        'status_col': 'SMS Status Response Date/Time',
        'phone_col': 'Phone Number'
    }
    
    available_cols = [col.strip() for col in df.columns]
    missing_cols = []
    col_mapping = {}
    
    for key, expected_col in required_columns.items():
        found = False
        for actual_col in available_cols:
            if actual_col.lower() == expected_col.lower().strip():
                col_mapping[key] = actual_col
                found = True
                break
        if not found:
            missing_cols.append(expected_col)
    
    if missing_cols:
        st.error(f"The following required columns are missing from your data: {', '.join(missing_cols)}")
        st.write("Available columns in your data:", list(df.columns))
        return None, None
    
    df_processed = df.copy()
    df_processed = df_processed.rename(columns={
        col_mapping['date_col']: 'DATE',
        col_mapping['env_col']: 'ENVIRONMENT',
        col_mapping['client_col']: 'CLIENT',
        col_mapping['status_col']: 'STATUS',
        col_mapping['phone_col']: 'PHONE'
    })
    
    df_processed['DATE'] = pd.to_datetime(df_processed['DATE']).dt.date
    
    daily_summary = df_processed.groupby(['DATE', 'ENVIRONMENT', 'CLIENT', 'Source_File']).agg({
        'PHONE': 'count',
        'STATUS': [lambda x: x.notnull().sum(), lambda x: x.isnull().sum()]
    }).reset_index()
    
    daily_summary.columns = ['DATE', 'ENVIRONMENT', 'CLIENT', 'SOURCE_FILE', 'SMS SENDING', 'DELIVERED', 'FAILED']
    daily_summary = daily_summary.sort_values(['DATE', 'CLIENT'])
    
    min_date = df_processed['DATE'].min()
    max_date = df_processed['DATE'].max()
    date_range_str = f"{min_date.strftime('%B %d')} - {max_date.strftime('%B %d, %Y')}"
    
    overall_summary = df_processed.groupby(['ENVIRONMENT', 'CLIENT', 'Source_File']).agg({
        'PHONE': 'count',
        'STATUS': [lambda x: x.notnull().sum(), lambda x: x.isnull().sum()]
    }).reset_index()
    
    overall_summary.columns = ['ENVIRONMENT', 'CLIENT', 'SOURCE_FILE', 'SMS SENDING', 'DELIVERED', 'FAILED']
    overall_summary.insert(0, 'DATE_RANGE', date_range_str)
    overall_summary = overall_summary.sort_values(['CLIENT'])
    
    return daily_summary, overall_summary

with st.sidebar:
    st.subheader("Upload Files")
    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Select any number of Excel files to process"
    )

if uploaded_files:
    df = load_data(uploaded_files)
    
    daily_summary_df, overall_summary_df = create_sms_summaries(df)
    
    if daily_summary_df is not None and overall_summary_df is not None:
        numeric_cols = ['SMS SENDING', 'DELIVERED', 'FAILED']
        
        daily_summary_display = format_with_commas(daily_summary_df, numeric_cols)
        overall_summary_display = format_with_commas(overall_summary_df, numeric_cols)
        
        st.subheader("Overall SMS Summary")
        st.dataframe(overall_summary_display, use_container_width=True)
        
        overall_excel_data = to_excel_single(overall_summary_df, "Overall_SMS_Summary")
        st.download_button(
            label="Download Overall SMS Summary",
            data=overall_excel_data,
            file_name="overall_sms_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.subheader("SMS Summary per Client per Day")
        st.dataframe(daily_summary_display, use_container_width=True)
        
        daily_excel_data = to_excel_single(daily_summary_df, "Daily_SMS_Summary")
        st.download_button(
            label="Download Daily SMS Summary",
            data=daily_excel_data,
            file_name="daily_sms_summary_per_client_per_day.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        all_data_dict = {
            "Overall_Summary": overall_summary_df,
            "Daily_Summary": daily_summary_df
        }
        all_excel_data = to_excel_multiple(all_data_dict)
        st.download_button(
            label="Download All Summaries",
            data=all_excel_data,
            file_name="all_sms_summaries.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with st.expander("View Combined Raw Data"):
        st.dataframe(df, use_container_width=True)
