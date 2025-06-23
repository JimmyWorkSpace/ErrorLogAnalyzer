from error_log_processor import process_error_log
import os

if __name__ == "__main__":
    input_path = os.path.join("data", "ErrorTimeStatistics.xlsx")

    # This must be the full path to the existing Excel file
    existing_excel_path = r"C:\Users\Jimmy\Project\DataAutoAnalyzer\output\OHT_Daily_Report.xlsx"

    process_error_log(input_path, existing_excel_path)
