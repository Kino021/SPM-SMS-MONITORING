# Add these new functions after the existing ones

def calculate_sms_status(df):
    # Convert date columns to datetime
    df['SUBMISSION DATE / TIME'] = pd.to_datetime(df['SUBMISSION DATE / TIME'], errors='coerce')
    df['SMS STATUS RESPONSE DATE/TIME'] = pd.to_datetime(df['SMS STATUS RESPONSE DATE/TIME'], errors='coerce')
    
    # Calculate delivered count
    delivered_df = df[df['COL STATUS'].str.contains('DELIVERED', case=False, na=False)].copy()
    delivered_by_date = delivered_df.groupby(delivered_df['SUBMISSION DATE / TIME'].dt.date).agg({
        'ACCOUNT NO.': 'count'
    }).rename(columns={'ACCOUNT NO.': 'DELIVERED SMS'})
    
    # Calculate failed count (blank SMS Status Response Date/Time)
    failed_df = df[df['SMS STATUS RESPONSE DATE/TIME'].isna() & 
                  df['COL STATUS'].notna()].copy()
    failed_by_date = failed_df.groupby(failed_df['SUBMISSION DATE / TIME'].dt.date).agg({
        'ACCOUNT NO.': 'count'
    }).rename(columns={'ACCOUNT NO.': 'FAILED SMS'})
    
    return delivered_by_date, failed_by_date

# Modify the existing calculate_summary_per_collector function to include SMS stats
def calculate_summary_per_collector(df, remark_types):
    summary_columns = [
        'DATE', 'CLIENT', 'COLLECTOR', 'MANUAL CALL', 'MANUAL ACCOUNT', 
        'PREDICTIVE CONNECTED', 'MANUAL CONNECTED', 'CONNECTED UNIQUE', 
        'CONNECTED NOT UNIQUE', 'TOTAL TALK TIME', 'MANUAL TALK TIME', 
        'PREDICTIVE TALK TIME', 'PTP ACC', 'TOTAL PTP AMOUNT', 'TOTAL BALANCE',
        'DELIVERED SMS', 'FAILED SMS'  # Added new columns
    ]
    
    summary_table = pd.DataFrame(columns=summary_columns)
    
    df_filtered = df[df['REMARK TYPE'].isin(remark_types)].copy()
    df_filtered['DATE'] = df_filtered['DATE'].dt.date  

    for (date, client, collector), group in df_filtered.groupby(['DATE', 'CLIENT', 'REMARK BY']):
        if not group['CALL DURATION'].notna().any():
            continue
        
        # Existing calculations...
        manual_call = group[group['REMARK TYPE'] == 'Outgoing']['ACCOUNT NO.'].count()
        manual_account = group[group['REMARK TYPE'] == 'Outgoing']['ACCOUNT NO.'].nunique()
        predictive_connected = group[(group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])) & 
                                   (group['CALL STATUS'] == 'CONNECTED')]['ACCOUNT NO.'].count()
        manual_connected = group[(group['REMARK TYPE'] == 'Outgoing') & 
                               (group['CALL STATUS'] == 'CONNECTED')]['ACCOUNT NO.'].count()
        connected_unique = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].nunique()
        connected_not_unique = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].count()
        ptp_acc = group[(group['STATUS'].str.contains('PTP', na=False)) & 
                      (group['PTP AMOUNT'] != 0)]['ACCOUNT NO.'].nunique()
        total_ptp_amount = group[(group['STATUS'].str.contains('PTP', na=False)) & 
                               (group['PTP AMOUNT'] != 0)]['PTP AMOUNT'].sum()
        total_balance = group[(group['PTP AMOUNT'] != 0)]['BALANCE'].sum()

        total_talk_seconds = group['TALK TIME DURATION'].sum()
        total_talk_time = format_seconds_to_hms(total_talk_seconds)
        manual_talk_seconds = group[group['REMARK TYPE'] == 'Outgoing']['TALK TIME DURATION'].sum()
        manual_talk_time = format_seconds_to_hms(manual_talk_seconds)
        predictive_talk_seconds = group[group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])]['TALK TIME DURATION'].sum()
        predictive_talk_time = format_seconds_to_hms(predictive_talk_seconds)

        # New SMS calculations
        delivered_sms = group[group['COL STATUS'].str.contains('DELIVERED', case=False, na=False)]['ACCOUNT NO.'].count()
        failed_sms = group[group['SMS STATUS RESPONSE DATE/TIME'].isna() & 
                         group['COL STATUS'].notna()]['ACCOUNT NO.'].count()

        summary_data = {
            'DATE': date,
            'CLIENT': client,
            'COLLECTOR': collector,
            'MANUAL CALL': manual_call,
            'MANUAL ACCOUNT': manual_account,
            'PREDICTIVE CONNECTED': predictive_connected,
            'MANUAL CONNECTED': manual_connected,
            'CONNECTED UNIQUE': connected_unique,
            'CONNECTED NOT UNIQUE': connected_not_unique,
            'TOTAL TALK TIME': total_talk_time,
            'MANUAL TALK TIME': manual_talk_time,
            'PREDICTIVE TALK TIME': predictive_talk_time,
            'PTP ACC': ptp_acc,
            'TOTAL PTP AMOUNT': total_ptp_amount,
            'TOTAL BALANCE': total_balance,
            'DELIVERED SMS': delivered_sms,
            'FAILED SMS': failed_sms
        }
        
        summary_table = pd.concat([summary_table, pd.DataFrame([summary_data])], ignore_index=True)
    
    return summary_table.sort_values(by=['DATE', 'COLLECTOR'])

# Modify the calculate_summary_per_day function to include SMS stats
def calculate_summary_per_day(df, remark_types):
    summary_columns = [
        'DATE', 'CLIENT', 'COLLECTOR', 'MANUAL CALL', 'MANUAL ACCOUNT', 
        'PREDICTIVE CONNECTED', 'MANUAL CONNECTED', 'CONNECTED UNIQUE', 
        'CONNECTED NOT UNIQUE', 'TOTAL TALK TIME', 'MANUAL TALK TIME', 
        'PREDICTIVE TALK TIME', 'PTP ACC', 'TOTAL PTP AMOUNT', 'TOTAL BALANCE',
        'MANUAL CALL AVG', 'MANUAL CONNECTED AVG', 'PREDICTIVE CONNECTED AVG',
        'TOTAL CONNECTED AVG', 'MANUAL TALK TIME AVG', 'PREDICTIVE TALK TIME AVG',
        'TOTAL TALK TIME AVG', 'PTP COUNT AVG', 'PTP AMOUNT AVG', 'BALANCE AVG',
        'DELIVERED SMS', 'FAILED SMS', 'DELIVERED SMS AVG', 'FAILED SMS AVG'  # Added new columns
    ]
    
    summary_table = pd.DataFrame(columns=summary_columns)
    
    df_filtered = df[df['REMARK TYPE'].isin(remark_types)].copy()
    df_filtered['DATE'] = df_filtered['DATE'].dt.date  

    for (date, client), group in df_filtered.groupby(['DATE', 'CLIENT']):
        if not group['CALL DURATION'].notna().any():
            continue
        
        collector_count = group['REMARK BY'].nunique()
        
        # Existing calculations...
        manual_call = group[group['REMARK TYPE'] == 'Outgoing']['ACCOUNT NO.'].count()
        manual_account = group[group['REMARK TYPE'] == 'Outgoing']['ACCOUNT NO.'].nunique()
        predictive_connected = group[(group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])) & 
                                   (group['CALL STATUS'] == 'CONNECTED')]['ACCOUNT NO.'].count()
        manual_connected = group[(group['REMARK TYPE'] == 'Outgoing') & 
                               (group['CALL STATUS'] == 'CONNECTED')]['ACCOUNT NO.'].count()
        connected_unique = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].nunique()
        connected_not_unique = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].count()
        ptp_acc = group[(group['STATUS'].str.contains('PTP', na=False)) & 
                      (group['PTP AMOUNT'] != 0)]['ACCOUNT NO.'].nunique()
        total_ptp_amount = group[(group['STATUS'].str.contains('PTP', na=False)) & 
                               (group['PTP AMOUNT'] != 0)]['PTP AMOUNT'].sum()
        total_balance = group[(group['PTP AMOUNT'] != 0)]['BALANCE'].sum()

        total_talk_seconds = group['TALK TIME DURATION'].sum()
        total_talk_time = format_seconds_to_hms(total_talk_seconds)
        manual_talk_seconds = group[group['REMARK TYPE'] == 'Outgoing']['TALK TIME DURATION'].sum()
        manual_talk_time = format_seconds_to_hms(manual_talk_seconds)
        predictive_talk_seconds = group[group['REMARK TYPE'].isin(['Predictive', 'Follow Up'])]['TALK TIME DURATION'].sum()
        predictive_talk_time = format_seconds_to_hms(predictive_talk_seconds)

        # Existing averages...
        manual_call_avg = round(manual_call / collector_count if collector_count > 0 else 0, 2)
        manual_connected_avg = round(manual_connected / collector_count if collector_count > 0 else 0, 2)
        predictive_connected_avg = round(predictive_connected / collector_count if collector_count > 0 else 0, 2)
        total_connected_avg = round(connected_not_unique / collector_count if collector_count > 0 else 0, 2)
        
        manual_talk_time_avg_seconds = manual_talk_seconds / collector_count if collector_count > 0 else 0
        manual_talk_time_avg = format_seconds_to_hms(manual_talk_time_avg_seconds)
        
        predictive_talk_time_avg_seconds = predictive_talk_seconds / collector_count if collector_count > 0 else 0
        predictive_talk_time_avg = format_seconds_to_hms(predictive_talk_time_avg_seconds)
        
        total_talk_time_avg_seconds = total_talk_seconds / collector_count if collector_count > 0 else 0
        total_talk_time_avg = format_seconds_to_hms(total_talk_time_avg_seconds)
        
        ptp_count_avg = round(ptp_acc / collector_count if collector_count > 0 else 0, 2)
        ptp_amount_avg = round(total_ptp_amount / collector_count if collector_count > 0 else 0, 2)
        balance_avg = round(total_balance / collector_count if collector_count > 0 else 0, 2)

        # New SMS calculations
        delivered_sms = group[group['COL STATUS'].str.contains('DELIVERED', case=False, na=False)]['ACCOUNT NO.'].count()
        failed_sms = group[group['SMS STATUS RESPONSE DATE/TIME'].isna() & 
                         group['COL STATUS'].notna()]['ACCOUNT NO.'].count()
        delivered_sms_avg = round(delivered_sms / collector_count if collector_count > 0 else 0, 2)
        failed_sms_avg = round(failed_sms / collector_count if collector_count > 0 else 0, 2)

        summary_data = {
            'DATE': date,
            'CLIENT': client,
            'COLLECTOR': collector_count,
            'MANUAL CALL': manual_call,
            'MANUAL ACCOUNT': manual_account,
            'PREDICTIVE CONNECTED': predictive_connected,
            'MANUAL CONNECTED': manual_connected,
            'CONNECTED UNIQUE': connected_unique,
            'CONNECTED NOT UNIQUE': connected_not_unique,
            'TOTAL TALK TIME': total_talk_time,
            'MANUAL TALK TIME': manual_talk_time,
            'PREDICTIVE TALK TIME': predictive_talk_time,
            'PTP ACC': ptp_acc,
            'TOTAL PTP AMOUNT': total_ptp_amount,
            'TOTAL BALANCE': total_balance,
            'MANUAL CALL AVG': manual_call_avg,
            'MANUAL CONNECTED AVG': manual_connected_avg,
            'PREDICTIVE CONNECTED AVG': predictive_connected_avg,
            'TOTAL CONNECTED AVG': total_connected_avg,
            'MANUAL TALK TIME AVG': manual_talk_time_avg,
            'PREDICTIVE TALK TIME AVG': predictive_talk_time_avg,
            'TOTAL TALK TIME AVG': total_talk_time_avg,
            'PTP COUNT AVG': ptp_count_avg,
            'PTP AMOUNT AVG': ptp_amount_avg,
            'BALANCE AVG': balance_avg,
            'DELIVERED SMS': delivered_sms,
            'FAILED SMS': failed_sms,
            'DELIVERED SMS AVG': delivered_sms_avg,
            'FAILED SMS AVG': failed_sms_avg
        }
        
        summary_table = pd.concat([summary_table, pd.DataFrame([summary_data])], ignore_index=True)
    
    return summary_table.sort_values(by=['DATE'])

# Modify the calculate_overall_summary function to include SMS stats
def calculate_overall_summary(daily_summary):
    summary_columns = [
        'DATE', 'CLIENT', 'COLLECTOR', 
        'MANUAL CALL AVG', 'MANUAL CONNECTED AVG', 'PREDICTIVE CONNECTED AVG',
        'TOTAL CONNECTED AVG', 'MANUAL TALK TIME AVG', 'PREDICTIVE TALK TIME AVG',
        'TOTAL TALK TIME AVG', 'PTP COUNT AVG', 'PTP AMOUNT AVG', 'BALANCE AVG',
        'DELIVERED SMS AVG', 'FAILED SMS AVG'  # Added new columns
    ]
    
    overall_summary = pd.DataFrame(columns=summary_columns)
    
    if 'DATE' not in daily_summary.columns:
        st.error("The 'DATE' column is missing in the Daily Summary data. Please ensure the uploaded files contain a 'Daily Summary' sheet with a 'DATE' column.")
        return overall_summary
    
    daily_summary['DATE'] = pd.to_datetime(daily_summary['DATE'], errors='coerce')
    min_date = daily_summary['DATE'].min()
    max_date = daily_summary['DATE'].max()
    if pd.isna(min_date) or pd.isna(max_date):
        date_range = "Unknown Date Range"
    else:
        date_range = f"{min_date.strftime('%B %d')} - {max_date.strftime('%B %d, %Y')}"
    
    for client, group in daily_summary.groupby('CLIENT'):
        collector_avg = int(round(group['COLLECTOR'].mean(), 0))
        manual_call_avg = int(round(group['MANUAL CALL AVG'].mean(), 0))
        manual_connected_avg = int(round(group['MANUAL CONNECTED AVG'].mean(), 0))
        predictive_connected_avg = int(round(group['PREDICTIVE CONNECTED AVG'].mean(), 0))
        total_connected_avg = int(round(group['TOTAL CONNECTED AVG'].mean(), 0))
        
        manual_talk_seconds = group['MANUAL TALK TIME AVG'].apply(hms_to_seconds).mean()
        manual_talk_time_avg = format_seconds_to_hms(manual_talk_seconds)
        
        predictive_talk_seconds = group['PREDICTIVE TALK TIME AVG'].apply(hms_to_seconds).mean()
        predictive_talk_time_avg = format_seconds_to_hms(predictive_talk_seconds)
        
        total_talk_seconds = group['TOTAL TALK TIME AVG'].apply(hms_to_seconds).mean()
        total_talk_time_avg = format_seconds_to_hms(total_talk_seconds)
        
        ptp_count_avg = int(round(group['PTP COUNT AVG'].mean(), 0))
        ptp_amount_avg = round(group['PTP AMOUNT AVG'].mean(), 2)
        balance_avg = round(group['BALANCE AVG'].mean(), 2)
        
        delivered_sms_avg = int(round(group['DELIVERED SMS AVG'].mean(), 0))
        failed_sms_avg = int(round(group['FAILED SMS AVG'].mean(), 0))

        summary_data = {
            'DATE': date_range,
            'CLIENT': client,
            'COLLECTOR': collector_avg,
            'MANUAL CALL AVG': manual_call_avg,
            'MANUAL CONNECTED AVG': manual_connected_avg,
            'PREDICTIVE CONNECTED AVG': predictive_connected_avg,
            'TOTAL CONNECTED AVG': total_connected_avg,
            'MANUAL TALK TIME AVG': manual_talk_time_avg,
            'PREDICTIVE TALK TIME AVG': predictive_talk_time_avg,
            'TOTAL TALK TIME AVG': total_talk_time_avg,
            'PTP COUNT AVG': ptp_count_avg,
            'PTP AMOUNT AVG': ptp_amount_avg,
            'BALANCE AVG': balance_avg,
            'DELIVERED SMS AVG': delivered_sms_avg,
            'FAILED SMS AVG': failed_sms_avg
        }
        
        overall_summary = pd.concat([overall_summary, pd.DataFrame([summary_data])], ignore_index=True)
    
    return overall_summary

# Update the to_excel function to handle new columns
def to_excel(df_dict):
    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter', date_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00'})
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        header_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': 'red', 'font_color': 'white', 'bold': True})
        comma_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'})
        date_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'yyyy-mm-dd'})
        time_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'hh:mm:ss'})
        avg_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0.00'})
        int_avg_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'})
        
        for sheet_name, df in df_dict.items():
            sheet_name = sanitize_sheet_name(sheet_name)
            df_for_excel = df.copy()
            
            df_for_excel.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            worksheet = writer.sheets[sheet_name]
            
            worksheet.merge_range('A1:' + chr(65 + len(df.columns) - 1) + '1', sheet_name, title_format)
            
            for col_num, col_name in enumerate(df_for_excel.columns):
                worksheet.write(1, col_num, col_name, header_format)
            
            for row_num in range(2, len(df_for_excel) + 2):
                for col_num, col_name in enumerate(df_for_excel.columns):
                    value = df_for_excel.iloc[row_num - 2, col_num]
                    if col_name == 'DATE':
                        if sheet_name == 'Overall Summary':
                            worksheet.write(row_num, col_num, value, center_format)
                        else:
                            if isinstance(value, (pd.Timestamp, datetime.date)):
                                worksheet.write_datetime(row_num, col_num, value, date_format)
                            else:
                                worksheet.write(row_num, col_num, value, date_format)
                    elif col_name in ['TOTAL PTP AMOUNT', 'TOTAL BALANCE']:
                        worksheet.write(row_num, col_num, value, comma_format)
                    elif col_name in ['TOTAL TALK TIME', 'MANUAL TALK TIME', 'PREDICTIVE TALK TIME', 
                                    'TOTAL TALK TIME AVG', 'MANUAL TALK TIME AVG', 'PREDICTIVE TALK TIME AVG']:
                        worksheet.write(row_num, col_num, value, time_format)
                    elif col_name in ['COLLECTOR', 'MANUAL CALL AVG', 'MANUAL CONNECTED AVG', 
                                    'PREDICTIVE CONNECTED AVG', 'TOTAL CONNECTED AVG', 'PTP COUNT AVG',
                                    'DELIVERED SMS', 'FAILED SMS', 'DELIVERED SMS AVG', 'FAILED SMS AVG']:
                        worksheet.write(row_num, col_num, value, int_avg_format)
                    elif 'AVG' in col_name:
                        worksheet.write(row_num, col_num, value, avg_format)
                    else:
                        worksheet.write(row_num, col_num, value, center_format)
            
            for col_num, col_name in enumerate(df_for_excel.columns):
                max_len = max(
                    df_for_excel[col_name].astype(str).str.len().max(),
                    len(col_name)
                ) + 2
                worksheet.set_column(col_num, col_num, max_len)
    return output.getvalue()
