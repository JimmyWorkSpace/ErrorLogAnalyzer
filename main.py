from error_log_processor import process_error_log
import os
import re
from datetime import datetime, timedelta

if __name__ == "__main__":
    data_dir = "data"
    os.makedirs(data_dir, exist_ok=True)

    # ✅ Step 1: Find the error log file
    error_file = next((f for f in os.listdir(data_dir) if f.startswith("ErrorTimeStatistics")), None)
    if not error_file:
        raise FileNotFoundError("❌ No file starting with 'ErrorTimeStatistics' found in /data")

    input_path = os.path.join(data_dir, error_file)

    # ✅ Step 2: Use full path to existing daily report file
    existing_excel_path = r"C:\Users\Jimmy\Project\DataAutoAnalyzer\output\OHT_Daily_Report.xlsx"

    # ✅ Step 3: Extract sheet_date from filename
    match = re.search(r'20\d{6}', error_file)
    if match:
        file_date = datetime.strptime(match.group(), '%Y%m%d')
        sheet_date = (file_date - timedelta(days=1)).strftime('%Y%m%d')
        print(f"📆 Detected date in filename: {match.group()} → Using sheet_date: {sheet_date}")
    else:
        sheet_date = (datetime.today() - timedelta(days=1)).strftime('%Y%m%d')
        print(f"⚠️ No valid date found in filename — defaulting to: {sheet_date}")

    # ✅ Step 4: Call the processor
    process_error_log(input_path=input_path, existing_excel_path=existing_excel_path, sheet_date=sheet_date)
