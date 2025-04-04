# Function to process data and create daily and overall summaries
def create_sms_summaries(df):
    required_columns = {
        'date_col': 'Submission Date / Time',
        'env_col': 'ENVIRONMENT',  # Changed to uppercase for consistency
        'client_col': 'Client',
        'status_col': 'SMS Status Response Date/Time',
        'phone_col': 'Phone Number'
    }
    
    # Check for missing columns with case-insensitive comparison
    missing_cols = [col for name, col in required_columns.items() if col not in [x.upper() for x in df.columns]]
    if missing_cols:
        st.error(f"THE FOLLOWING REQUIRED COLUMNS ARE MISSING FROM YOUR DATA: {', '.join(missing_cols)}")  # Uppercase error message
        st.write("AVAILABLE COLUMNS IN YOUR DATA:", [col.upper() for col in df.columns])  # Uppercase available columns
        return None, None
    
    df_processed = df.copy()
    # Rename columns to uppercase for consistency
    df_processed = df_processed.rename(columns={
        required_columns['date_col']: 'DATE',
        required_columns['env_col']: 'ENVIRONMENT',
        required_columns['client_col']: 'CLIENT',
        required_columns['status_col']: 'STATUS',
        required_columns['phone_col']: 'PHONE'
    })
    
    df_processed['DATE'] = pd.to_datetime(df_processed['DATE']).dt.date
    
    daily_summary = df_processed.groupby(['DATE', 'ENVIRONMENT', 'CLIENT']).agg({
        'PHONE': 'count',
        'STATUS': [
            lambda x: x.notnull().sum(),
            lambda x: x.isnull().sum()
        ]
    }).reset_index()
    
    daily_summary.columns = ['DATE', 'ENVIRONMENT', 'CLIENT', 'SMS SENDING', 'DELIVERED', 'FAILED']
    daily_summary = daily_summary.sort_values(['DATE', 'CLIENT'])
    
    min_date = df_processed['DATE'].min()
    max_date = df_processed['DATE'].max()
    date_range_str = f"{min_date.strftime('%B %d')} - {max_date.strftime('%B %d, %Y')}"
    
    overall_summary = df_processed.groupby(['ENVIRONMENT', 'CLIENT']).agg({
        'PHONE': 'count',
        'STATUS': [
            lambda x: x.notnull().sum(),
            lambda x: x.isnull().sum()
        ]
    }).reset_index()
    
    overall_summary.columns = ['ENVIRONMENT', 'CLIENT', 'SMS SENDING', 'DELIVERED', 'FAILED']
    overall_summary.insert(0, 'DATE_RANGE', date_range_str)
    overall_summary = overall_summary.sort_values(['CLIENT'])
    
    return daily_summary, overall_summary
