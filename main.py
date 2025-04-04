import streamlit as st
import pandas as pd
from io import BytesIO
import math

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
        color: #87CEEB !important;  /* Light blue color */
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
    # Convert Submission Date / Time to date only
    df['Submission Date / Time'] = pd.to_datetime(df['Submission Date / Time']).dt.date
    
    # Create summary DataFrame
    summary = df.groupby(['Submission Date / Time', 'ENVIRONMENT', 'CLIENT']).agg({
        'ACCOUNT NO.': 'nunique',  # Count unique accounts
        'STATUS': [
            lambda x: (x.str.upper() == 'DELIVERED').sum(),  # SMS Sent
            lambda x: (x.str.upper() == 'FAILED').sum()      # SMS Not Sent
        ]
    }).reset_index()
    
    # Rename columns
    summary.columns = ['DATE', 'ENVIRONMENT', 'CLIENT', 'ACCOUNTS', 'SMS SENT', 'SMS NOT SENT']
    
    return summary

with st.sidebar:
    st.subheader("Upload File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    # Load and process data
    df = load_data(uploaded_file)
    
    # Create SMS summary
    summary_df = create_sms_summary(df)
    
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
