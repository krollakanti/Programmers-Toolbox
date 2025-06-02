import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime

# ---------- Utility Functions ----------

def run_macro_usage_check(macro_dir, program_dir):
    macro_files = [f for f in os.listdir(macro_dir) if f.lower().endswith(".sas")]
    macro_names = [os.path.splitext(f)[0].lower() for f in macro_files]

    sas_program_texts = []
    for root, _, files in os.walk(program_dir):
        for file in files:
            if file.lower().endswith(".sas"):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                        sas_program_texts.append(f.read().lower())
                except Exception as e:
                    st.warning(f"Could not read {file_path}: {e}")

    all_program_text = "\n".join(sas_program_texts)
    macro_status = []
    for macro in macro_names:
        is_used = f"%{macro}" in all_program_text
        macro_status.append({
            "List of macros": macro,
            "Status": "Used" if is_used else "Not Used"
        })

    df = pd.DataFrame(macro_status)
    return df

def run_search_for_terms(program_dir, terms):
    result = []
    for root, dirs, files in os.walk(program_dir):
        for file in files:
            if file.endswith(".sas"):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        lines = f.readlines()
                    for line_num, line in enumerate(lines, start=1):
                        for term in terms:
                            if re.search(re.escape(term), line, re.IGNORECASE):
                                result.append({
                                    'Program Name': file,
                                    'Line Number': line_num,
                                    'Line Code': line.strip(),
                                    'Identified Term': term
                                })
                except Exception as e:
                    st.warning(f"Error reading file {file_path}: {e}")
    return pd.DataFrame(result)

def run_search_and_replace_terms(program_dir, replace_dict):
    result = []
    for root, dirs, files in os.walk(program_dir):
        for file in files:
            if file.endswith(".sas"):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        lines = f.readlines()

                    updated_lines = []
                    file_modified = False
                    for line_num, line in enumerate(lines, start=1):
                        original_line = line
                        for search_term, replace_term in replace_dict.items():
                            if re.search(re.escape(search_term), line, re.IGNORECASE):
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
                    st.warning(f"Error processing file {file_path}: {e}")
    return pd.DataFrame(result)

# ---------- Streamlit App UI ----------

st.title("ClinSage-Programmers Toolbox")

task = st.selectbox("Select a task:", ["Macro Usage Check", "Search for Terms", "Search and Replace Terms"])

with st.form("input_form"):
    if task == "Macro Usage Check":
        macro_dir = st.text_input("Path to Macro Folder")
        program_dir = st.text_input("Path to SAS Programs Folder")

    elif task == "Search for Terms":
        program_dir = st.text_input("Path to SAS Programs Folder")
        terms_text = st.text_area("Enter terms to search (comma-separated)")

    elif task == "Search and Replace Terms":
        program_dir = st.text_input("Path to SAS Programs Folder")
        terms_text = st.text_area("Enter search and replace terms as 'search:replace' per line")

    submitted = st.form_submit_button("Process Task")

if submitted:
    with st.spinner("Processing..."):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"report_{task.replace(' ', '_').lower()}_{timestamp}.xlsx"

        if task == "Macro Usage Check":
            df = run_macro_usage_check(macro_dir, program_dir)

        elif task == "Search for Terms":
            search_terms = [t.strip() for t in terms_text.split(",") if t.strip()]
            df = run_search_for_terms(program_dir, search_terms)

        elif task == "Search and Replace Terms":
            replace_dict = {}
            for line in terms_text.strip().split("\n"):
                if ":" in line:
                    key, val = line.split(":", 1)
                    replace_dict[key.strip()] = val.strip()
            df = run_search_and_replace_terms(program_dir, replace_dict)

        if not df.empty:
            df.to_excel(output_path, index=False)
            st.success("✅ Task completed successfully.")
            with open(output_path, "rb") as f:
                st.download_button("📥 Download Report", data=f, file_name=output_path, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("No matches found or no data generated.")
