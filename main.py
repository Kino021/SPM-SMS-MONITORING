import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide", page_title="DIALER PRODUCTIVITY PER CRITERIA OF BALANCE", page_icon="📊", initial_sidebar_state="expanded")

# Apply dark mode and custom header styling
st.markdown(
    """
    <style>
    .reportview-container {
        background: #2E2E2E;
        color: white;
    }
    .sidebar .sidebar-content {
        background: #2E2E2E;
    }
    h1, h2, h3 {
        color: #87CEEB !important;
        font-weight: bold !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title('SPM SMS MONITORING ALL ENVI')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Function to convert single DataFrame to Excel bytes with formatting
def to_excel_single(df, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_copy = df.copy()
        if 'DATE' in df_copy.columns:
            if pd.api.types.is_datetime64_any_dtype(df_copy['DATE']):
                df_copy['DATE'] = df_copy['DATE'].dt.strftime('%d-%m-%Y')
            elif pd.api.types.is_object_dtype(df_copy['DATE']):
                df_copy['DATE'] = pd.to_datetime(df_copy['DATE'], errors='coerce').dt.strftime('%d-%m-%Y')
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
        cell_format = workbook.add_format({'border': 1})
        
        for col_num, value in enumerate(df_copy.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        for row_num in range(1, len(df_copy) + 1):
            for col_num in range(len(df_copy.columns)):
                worksheet.write(row_num, col_num, df_copy.iloc[row_num-1, col_num], cell_format)
        
        for i, col in enumerate(df_copy.columns):
            max_length = max(
                df_copy[col].astype(str).map(len).max(),
                len(str(col))
            )
            worksheet.set_column(i, i, max_length + 2)
    
    return output.getvalue()

# Function to process data and create summary
def create_sms_summary(df):
    # Required columns with corrected names
    required_columns = {
        'date_col': 'Submission Date / Time',
        'env_col': 'Environment 2',
        'client_col': 'Client',
        'account_col': 'Account No.',
        'status_col': 'SMS Status Response Date/Time'
    }
    
    # Check if all required columns exist
    missing_cols = [col for name, col in required_columns.items() if col not in df.columns]
    if missing_cols:
        st.error(f"The following required columns are missing from your data: {', '.join(missing_cols)}")
        st.write("Available columns in your data:", list(df.columns))
        return None
    
    # Create a copy with renamed columns for consistency
    df_processed = df.copy()
    df_processed = df_processed.rename(columns={
        required_columns['date_col']: 'DATE',
        required_columns['env_col']: 'ENVIRONMENT',
        required_columns['client_col']: 'CLIENT',
        required_columns['account_col']: 'ACCOUNT',
        required_columns['status_col']: 'STATUS'
    })
    
    # Convert date column to date only
    df_processed['DATE'] = pd.to_datetime(df_processed['DATE']).dt.date
    
    # Infer SMS status: if STATUS (SMS Status Response Date/Time) is not null, it's sent; otherwise, it's not sent
    df_processed['SMS_SENT'] = df_processed['STATUS'].notnull().astype(int)  # 1 if sent, 0 if not
    df_processed['SMS_NOT_SENT'] = df_processed['STATUS'].isnull().astype(int)  # 1 if not sent, 0 if sent
    
    # Create summary DataFrame
    summary = df_processed.groupby(['DATE', 'ENVIRONMENT', 'CLIENT']).agg({
        'ACCOUNT': 'nunique',  # Count unique accounts
        'SMS_SENT': 'sum',     # Total SMS Sent
        'SMS_NOT_SENT': 'sum'  # Total SMS Not Sent
    }).reset_index()
    
    # Rename columns
    summary.columns = ['DATE', 'ENVIRONMENT', 'CLIENT', 'ACCOUNTS', 'SMS SENT', 'SMS NOT SENT']
    
    # Sort by date and client
    summary = summary.sort_values(['DATE', 'CLIENT'])
    
    return summary

with st.sidebar:
    st.subheader("Upload File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    # Load and process data
    df = load_data(uploaded_file)
    
    # Create SMS summary
    summary_df = create_sms_summary(df)
    
    if summary_df is not None:
        # Display the summary
        st.subheader("SMS Summary per Client per Day")
        st.dataframe(summary_df)
        
        # Download button for summary
        excel_data = to_excel_single(summary_df, "SMS_Summary")
        st.download_button(
            label="Download SMS Summary",
            data=excel_data,
            file_name="sms_summary_per_client_per_day.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Display raw data (optional)
    with st.expander("View Raw Data"):
        st.dataframe(df)
