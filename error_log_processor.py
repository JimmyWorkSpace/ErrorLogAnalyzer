import pandas as pd
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook

def process_error_log(input_path: str, existing_excel_path: str):
    # Read Excel file
    df = pd.read_excel(input_path)

    # Ensure required columns exist
    required_cols = {'ERROR_MESSAGE', 'START_TIME', 'END_TIME', 'ERROR_LEVEL', 'ERROR_CODE'}
    if not required_cols.issubset(df.columns):
        missing = required_cols - set(df.columns)
        raise ValueError(f"Missing columns in input file: {missing}")

    # Convert time columns to datetime
    df['START_TIME'] = pd.to_datetime(df['START_TIME'], errors='coerce')
    df['END_TIME'] = pd.to_datetime(df['END_TIME'], errors='coerce')

    # Apply filters
    filtered_df = df[
        (df['ERROR_LEVEL'] == 'Alarm') &
        (df['ERROR_CODE'].astype(str).str.startswith('1')) &
        (df['ERROR_CODE'].astype(str) != '1001')
    ].copy()

    # Group and summarize
    output_rows = []
    grouped = filtered_df.groupby('ERROR_MESSAGE')
    total_count = 0

    for message, group in grouped:
        group_sorted = group.sort_values(by='START_TIME')
        count = len(group_sorted)
        total_count += count

        # Summary row
        output_rows.append({
            'ERROR_MESSAGE': message,
            'START_TIME': group_sorted['START_TIME'].min(),
            'END_TIME': group_sorted['END_TIME'].max(),
            'Count': count
        })

        # Detail rows
        for _, row in group_sorted.iterrows():
            output_rows.append({
                'ERROR_MESSAGE': '',
                'START_TIME': row['START_TIME'],
                'END_TIME': row['END_TIME'],
                'Count': ''
            })

    # Load Used OHTs from Utilization sheet
    book = load_workbook(existing_excel_path)
    sheet_date = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
    util_sheet_name = f"{sheet_date}_Utilization"

    if util_sheet_name not in book.sheetnames:
        raise ValueError(f"❌ Sheet '{util_sheet_name}' not found in the Excel file.")

    util_df = pd.read_excel(existing_excel_path, sheet_name=util_sheet_name)
    used_ohts = util_df.loc[0, 'Used OHTs']
    failure_rate = total_count / (used_ohts * 24)
    failure_rate_str = f"{failure_rate:.2%}"

    # Append failure rate row to error sheet
    output_rows.append({
        'ERROR_MESSAGE': 'Failure Rate',
        'START_TIME': '',
        'END_TIME': '',
        'Count': failure_rate_str
    })

    result_df = pd.DataFrame(output_rows)

    with pd.ExcelWriter(existing_excel_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        error_sheet_name = f"{sheet_date}_errorStatistics"
        result_df.to_excel(writer, sheet_name=error_sheet_name, index=False)

    print(f"✅ Exported to: {existing_excel_path} (sheet: {error_sheet_name})")

    # Update Weekly_Report
    weekly_df = pd.read_excel(existing_excel_path, sheet_name="Weekly_Report")

    target_date = (datetime.now() - timedelta(days=1)).date()
    for i, row in weekly_df.iterrows():
        try:
            row_date = pd.to_datetime(row['Date'], errors='coerce').date()
        except:
            continue
        if row_date == target_date:
            weekly_df.at[i, 'Failure Rate (%)'] = failure_rate
            print(f"📌 Inserted Failure Rate ({failure_rate_str}) into Weekly_Report at row {i+2}")
            break

    # Write back updated Weekly_Report
    with pd.ExcelWriter(existing_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        weekly_df.to_excel(writer, sheet_name='Weekly_Report', index=False)
