import os
import pandas as pd

# Local paths (Windows style)
macro_dir = r"J:\bdm\tbos\TAK279\studies\pso_3003\dmc5\macros"
program_dir = r"J:\bdm\tbos\TAK279\studies\pso_3003\dmc5\programs"

# Step 1: List all macro file names (without extension)
macro_files = [f for f in os.listdir(macro_dir) if f.lower().endswith(".sas")]
macro_names = [os.path.splitext(f)[0].lower() for f in macro_files]

# Step 2: Read all .sas program contents from program_dir and subfolders
sas_program_texts = []
for root, _, files in os.walk(program_dir):
    for file in files:
        if file.lower().endswith(".sas"):
            file_path = os.path.join(root, file)
            try:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    sas_program_texts.append(f.read().lower())
            except Exception as e:
                print(f"Could not read {file_path}: {e}")

all_program_text = "\n".join(sas_program_texts)

# Step 3: Check if each macro is used
macro_status = []
for macro in macro_names:
    is_used = f"%{macro}" in all_program_text
    macro_status.append({
        "List of macros": macro,
        "Status": "Used" if is_used else "Not Used"
    })

# Step 4: Save results to Excel
output_file = os.path.join(os.getcwd(), "macro_usage_status.xlsx")
df = pd.DataFrame(macro_status)
df.to_excel(output_file, index=False)

print(f"\n✅ Macro usage analysis complete. Output saved to:\n{output_file}")
