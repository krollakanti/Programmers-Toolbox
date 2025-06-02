import os
import re
import pandas as pd

def search_sas_programs(folder_path, search_terms):
    result = []

    # Traverse all files in the given folder and its subfolders
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".sas"):  # Check for .sas files
                file_path = os.path.join(root, file)
                print(f"Processing file: {file_path}")  # Debug - Check file path
                
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as sas_file:
                        lines = sas_file.readlines()

                        for line_num, line in enumerate(lines, start=1):
                            for term in search_terms:
                                # Case-insensitive search using re.IGNORECASE
                                if re.search(re.escape(term), line, re.IGNORECASE):
                                    print(f"Match found in {file} on line {line_num}: {line.strip()}")  # Debug - Check matches
                                    result.append({
                                        'Program Name': file,
                                        'Line Number': line_num,
                                        'Line Code': line.strip(),
                                        'Identified Term': term
                                    })
                except Exception as e:
                    print(f"Error reading file {file_path}: {e}")  # Debug - File read error

    # Convert results to a pandas DataFrame
    return pd.DataFrame(result)

def save_report(report_df, output_file):
    if not report_df.empty:
        report_df.to_excel(output_file, index=False)
        print(f"✅ Report saved to {output_file}")
    else:
        print("⚠️ No matches found. No data to save.")

# List of specific words or sentences to search for (case-insensitive)
search_terms = [
    'dmc4',
    'dev'
]

# Folder containing the SAS programs
sas_program_folder = r'J:\bdm\tbos\TAK279\studies\pso_3003\dmc5\programs'

# Search for terms and generate the report
report_df = search_sas_programs(sas_program_folder, search_terms)

# Save the report to an Excel file
save_report(report_df, 'sas_program_search_report.xlsx')
