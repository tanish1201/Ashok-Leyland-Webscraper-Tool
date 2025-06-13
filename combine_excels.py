import os
import pandas as pd

# Set the date to search for (today's date or a specific date)
target_date = '12-06-2025'  # Change this as needed

# Directory where Excel files are saved
DOWNLOAD_DIR = r'C:\Users\91987\TVS\downloads'

# Exclude the output file from the input list
files = [
    os.path.join(DOWNLOAD_DIR, f)
    for f in os.listdir(DOWNLOAD_DIR)
    if f.endswith('.xlsx') and target_date in f and f != f"Combined_Report_{target_date}.xlsx"
]

output_filename = f"Combined_Report_{target_date}.xlsx"
output_path = os.path.join(DOWNLOAD_DIR, output_filename)

sheets_written = 0
if files:
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for file_path in files:
            try:
                print(f"Reading: {os.path.basename(file_path)}")
                df = pd.read_excel(file_path, engine='openpyxl')
                if not df.empty:
                    sheet_name = os.path.splitext(os.path.basename(file_path))[0][:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    sheets_written += 1
                else:
                    print(f"Skipped empty file: {file_path}")
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
    if sheets_written:
        print(f"\nCombined file saved as: {output_filename}")
        print(f"Files combined as separate sheets: {sheets_written}")
    else:
        if os.path.exists(output_path):
            os.remove(output_path)
        print("No valid, non-empty Excel files found for the date in downloads directory.")
else:
    print("No valid Excel files found for the date in downloads directory.") 