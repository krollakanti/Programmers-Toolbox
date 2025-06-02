import os
import re
import pandas as pd

# === FUNCTION: Detects direct hardcoded patterns in a line ===
def find_hardcoding_in_line(line, line_num):
    issues = []

    # Each pattern is defined with:
    # - regex pattern
    # - human-readable description
    # - severity classification
    patterns = [
        (r"\b(where|if|when)\b\s+.*?=\s*['\"0-9]", "Hardcoded value in condition", "High"),
        (r"\b(set|merge)\b\s+[^;]*['\"0-9]", "Hardcoded dataset or literal in SET/MERGE", "High"),
        (r"\bput\s+['\"0-9]", "Hardcoded value in PUT statement", "High"),
        (r"['\"][A-Za-z0-9 _/-]{2,}['\"]", "Constant string literal", "Medium"),
        (r"\d{4}-\d{2}-\d{2}", "Hardcoded date", "High"),
        (r"\bselect\b\s+[^;]*['\"0-9]", "Hardcoded SELECT clause", "High"),
    ]

    for pattern, desc, severity in patterns:
        if re.search(pattern, line, re.IGNORECASE):
            issues.append((line_num, line.strip(), desc, severity))
    return issues

# === FUNCTION: Detects suspected misuse of macro variables ===
def find_macro_misuse(line, line_num):
    issues = []
    # Matches terms like study, site, dose, etc., that are often parameterized via macro variables
    if re.search(r"(?<!&)\b(study|site|visit|dose|group|arm|treatment)\b", line, re.IGNORECASE):
        if not re.search(r"&\w+", line):  # No macro variable used
            issues.append((line_num, line.strip(), "Possible macro variable misuse", "Low"))
    return issues

# === FUNCTION: Scans a single SAS file for hardcoding & macro misuse ===
def scan_sas_file(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
    issues = []

    for i, line in enumerate(lines, 1):
        issues.extend([(filepath, *issue) for issue in find_hardcoding_in_line(line, i)])
        issues.extend([(filepath, *issue) for issue in find_macro_misuse(line, i)])

    return issues

# === FUNCTION: Recursively scans a directory for SAS files and analyzes them ===
def scan_directory(folder_path):
    all_issues = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(".sas"):
                full_path = os.path.join(root, file)
                all_issues.extend(scan_sas_file(full_path))
    return all_issues

# === FUNCTION: Exports full issue list + summary report to Excel ===
def export_to_excel(results, output_file="results.xlsx"):
    df = pd.DataFrame(results, columns=["File", "Line Number", "Line", "Issue", "Severity"])
    summary = df["Severity"].value_counts().rename_axis("Severity").reset_index(name="Count")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Detailed Issues", index=False)
        summary.to_excel(writer, sheet_name="Summary by Severity", index=False)

    print(f"\n✅ Exported to {output_file}")
    print("\n📊 Summary by Severity:")
    print(summary.to_string(index=False))

# === MAIN: Prompts user for input path, scans, and reports ===
def main():
    folder_path = input("📁 Enter path to folder with SAS programs: ").strip()
    results = scan_directory(folder_path)

    if results:
        print("\n🔍 Issues Found (with Severity):\n")
        for file, line_num, line, issue, severity in results:
            print(f"[{file}] Line {line_num}: [{severity}] {issue}")
            print(f"    {line}\n")
        export_to_excel(results)
    else:
        print("✅ No issues found.")

if __name__ == "__main__":
    main()
