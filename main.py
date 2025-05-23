import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide", page_title="SPM SMS MONITORING", page_icon="📊", initial_sidebar_state="expanded")

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
def load_single_file(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    return df

def format_with_commas(df, numeric_cols):
    df_copy = df.copy()
    for col in numeric_cols:
        if col in df_copy.columns:
            df_copy[col] = df_copy[col].apply(lambda x: f"{int(x):,}" if pd.notnull(x) else x)
    return df_copy

# Excel writing functions remain unchanged
def to_excel_single(df, sheet_name):
    # [Previous implementation remains the same]
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
    # [Previous implementation remains the same]
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

def create_sms_summary_single_file(df):
    required_columns = {
        'date_col': 'SMS Status Response Date/Time',
        'env_col': 'Environment',
        'client_col': 'Client',
        'status_col': 'Status',
        'submission_date_col': 'Submission Date / Time',
        'phone_col': 'Phone Number'
    }
    
    available_cols = [col.strip().lower() for col in df.columns]
    missing_cols = []
    col_mapping = {}
    
    for key, expected_col in required_columns.items():
        found = False
        for actual_col in available_cols:
            if actual_col == expected_col.lower().strip():
                col_mapping[key] = next(col for col in df.columns if col.strip().lower() == actual_col)
                found = True
                break
        if not found:
            missing_cols.append(expected_col)
    
    if missing_cols:
        return None, missing_cols
    
    df_processed = df.copy()
    df_processed = df_processed.rename(columns={
        col_mapping['date_col']: 'DATE',
        col_mapping['env_col']: 'ENVIRONMENT',
        col_mapping['client_col']: 'CLIENT',
        col_mapping['status_col']: 'STATUS',
        col_mapping['submission_date_col']: 'SUBMISSION_DATE',
        col_mapping['phone_col']: 'PHONE'
    })
    
    df_processed['CLIENT'] = df_processed['CLIENT'].replace('', 'SYSTEM').fillna('SYSTEM')
    
    try:
        df_processed['DATE'] = pd.to_datetime(df_processed['DATE'], format='%d-%m-%Y %H:%M:%S', errors='coerce').dt.date
    except Exception:
        df_processed['DATE'] = pd.to_datetime(df_processed['DATE'], errors='coerce').dt.date
    
    try:
        df_processed['SUBMISSION_DATE'] = pd.to_datetime(df_processed['SUBMISSION_DATE'], format='%d-%m-%Y %H:%M:%S', errors='coerce').dt.date
    except Exception:
        df_processed['SUBMISSION_DATE'] = pd.to_datetime(df_processed['SUBMISSION_DATE'], errors='coerce').dt.date
    
    daily_summary = df_processed.groupby(['DATE', 'ENVIRONMENT', 'CLIENT']).apply(
        lambda x: pd.Series({
            'SMS SENDING': x['ENVIRONMENT'].notna().sum(),
            'DELIVERED': (x['STATUS'].str.lower() == 'delivered').sum(),
            'FAILED': (x['STATUS'].str.lower() == 'failed').sum()
        })
    ).reset_index()
    daily_summary = daily_summary.sort_values(['DATE', 'CLIENT'])
    
    return daily_summary, None

with st.sidebar:
    st.subheader("Upload Files")
    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Select any number of Excel files to process"
    )

if uploaded_files:
    all_daily_summaries = []
    all_dates = []
    
    # Process each file individually
    for file in uploaded_files:
        df = load_single_file(file)
        daily_summary, missing_cols = create_sms_summary_single_file(df)
        
        if daily_summary is not None:
            all_daily_summaries.append(daily_summary)
            valid_dates = daily_summary['DATE'].dropna()
            if not valid_dates.empty:
                all_dates.extend(valid_dates)
        else:
            st.error(f"Error processing {file.name}: Missing columns - {', '.join(missing_cols)}")
            st.write("Available columns:", list(df.columns))
    
    if all_daily_summaries:
        # Concatenate all daily summaries
        combined_daily_summary = pd.concat(all_daily_summaries, ignore_index=True)
        
        # Create overall summary from combined daily summary
        overall_summary = combined_daily_summary.groupby(['ENVIRONMENT', 'CLIENT']).agg({
            'SMS SENDING': 'sum',
            'DELIVERED': 'sum',
            'FAILED': 'sum'
        }).reset_index()
        
        # Calculate date range
        if all_dates:
            min_date = min(all_dates)
            max_date = max(all_dates)
            date_range_str = f"{min_date.strftime('%B %d')} - {max_date.strftime('%B %d, %Y')}"
        else:
            date_range_str = "Invalid Date Range"
        
        overall_summary.insert(0, 'DATE_RANGE', date_range_str)
        overall_summary = overall_summary.sort_values(['CLIENT'])
        
        numeric_cols = ['SMS SENDING', 'DELIVERED', 'FAILED']
        daily_summary_display = format_with_commas(combined_daily_summary, numeric_cols)
        overall_summary_display = format_with_commas(overall_summary, numeric_cols)
        
        st.subheader("Overall SMS Summary")
        st.dataframe(overall_summary_display, use_container_width=True)
        
        overall_excel_data = to_excel_single(overall_summary, "Overall_SMS_Summary")
        st.download_button(
            label="Download Overall SMS Summary",
            data=overall_excel_data,
            file_name="overall_sms_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.subheader("SMS Summary per Client per Day")
        st.dataframe(daily_summary_display, use_container_width=True)
        
        daily_excel_data = to_excel_single(combined_daily_summary, "Daily_SMS_Summary")
        st.download_button(
            label="Download Daily SMS Summary",
            data=daily_excel_data,
            file_name="daily_sms_summary_per_client_per_day.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        all_data_dict = {
            "Overall_Summary": overall_summary,
            "Daily_Summary": combined_daily_summary
        }
        all_excel_data = to_excel_multiple(all_data_dict)
        st.download_button(
            label="Download All Summaries",
            data=all_excel_data,
            file_name="all_sms_summaries.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
