import os
import re
import pandas as pd

def search_and_replace_sas_programs(folder_path, replace_dict):
    result = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".sas"):
                file_path = os.path.join(root, file)
                print(f"Processing file: {file_path}")

                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        lines = f.readlines()

                    updated_lines = []
                    file_modified = False

                    for line_num, line in enumerate(lines, start=1):
                        original_line = line
                        for search_term, replace_term in replace_dict.items():
                            # Case-insensitive match
                            if re.search(re.escape(search_term), line, re.IGNORECASE):
                                matched_text = re.findall(re.escape(search_term), line, re.IGNORECASE)
                                line = re.sub(re.escape(search_term), replace_term, line, flags=re.IGNORECASE)
                                result.append({
                                    'Program Name': file,
                                    'Line Number': line_num,
                                    'Original Line': original_line.strip(),
                                    'Modified Line': line.strip(),
                                    'Identified Term': search_term,
                                    'Replaced With': replace_term
                                })
                                file_modified = True
                        updated_lines.append(line)

                    if file_modified:
                        with open(file_path, 'w', encoding='utf-8', errors='ignore') as f:
                            f.writelines(updated_lines)

                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")

    return pd.DataFrame(result)

def save_report(report_df, output_file):
    if not report_df.empty:
        report_df.to_excel(output_file, index=False)
        print(f"✅ Report saved to {output_file}")
    else:
        print("⚠️ No matches found. No data to save.")

# Define search and replace terms (case-insensitive search)
replace_dict = {
    'check1': 'testval1',
    'Check 2': 'Test Value 2'
}

# Folder containing SAS programs
sas_program_folder = r'J:\bdm\tbos\TAK279\studies\pso_3003\dmc5\programs\test'

# Run search and replace
report_df = search_and_replace_sas_programs(sas_program_folder, replace_dict)

# Save Excel report
save_report(report_df, 'sas_program_search_replace_report.xlsx')
