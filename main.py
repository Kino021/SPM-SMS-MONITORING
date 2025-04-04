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

st.title('SPM MONITORING ALL ENVI')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Function to convert single DataFrame to Excel bytes with formatting
def to_excel_single(df, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Format dates properly before writing if 'Date' column exists
        df_copy = df.copy()
        if 'Date' in df_copy.columns:
            if pd.api.types.is_datetime64_any_dtype(df_copy['Date']):
                df_copy['Date'] = df_copy['Date'].dt.strftime('%d-%m-%Y')
            elif pd.api.types.is_object_dtype(df_copy['Date']):
                df_copy['Date'] = pd.to_datetime(df_copy['Date'], errors='coerce').dt.strftime('%d-%m-%Y')
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

# Function to combine all DataFrames into one Excel file
def to_excel_all(dfs, sheet_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheet_name in zip(dfs, sheet_names):
            df_copy = df.copy()
            if 'Date' in df_copy.columns:
                if pd.api.types.is_datetime64_any_dtype(df_copy['Date']):
                    df_copy['Date'] = df_copy['Date'].dt.strftime('%d-%m-%Y')
                elif pd.api.types.is_object_dtype(df_copy['Date']):
                    df_copy['Date'] = pd.to_datetime(df_copy['Date'], errors='coerce').dt.strftime('%d-%m-%Y')
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

with st.sidebar:
    st.subheader("Upload File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    df = load_data(uploaded_file)
