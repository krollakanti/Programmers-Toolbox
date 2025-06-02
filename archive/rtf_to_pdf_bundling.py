import os
import re
import win32com.client as win32
import time

# Paths
toc_rtf_path = r"J:\bdm\tbos\TAK279\studies\3001\dryrun1\tables\FTE review\rtfs\INDEX_TAK279_3001_dryrun1_TOC_bundle.rtf"
rtf_dir = os.path.dirname(toc_rtf_path)
pdf_dir = r"J:\bdm\tbos\TAK279\studies\3001\dryrun1\tables\FTE review\pdfs"

os.makedirs(pdf_dir, exist_ok=True)

# Mapping prefixes
prefix_map = {
    "Table": "t",
    "Listing": "l",
    "Figure": "f"
}

# Function to generate filename from title
def title_to_filename(title):
    for key in prefix_map:
        if title.startswith(key):
            code = title.split()[1].replace('.', '_')
            return f"{prefix_map[key]}{code}.rtf"
    return None

# Extract titles from TOC
with open(toc_rtf_path, "r", encoding="utf-8", errors="ignore") as f:
    rtf_text = f.read()

pattern = r"\b(?:Table|Listing|Figure)\s+\d+(?:\.\d+)*[a-zA-Z]?\b"
titles = re.findall(pattern, rtf_text)

# Launch Word
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

def convert_rtf_to_pdf(rtf_path, pdf_path):
    try:
        print(f"Opening: {rtf_path}")
        doc = word.Documents.Open(rtf_path)
        print(f"Saving as: {pdf_path}")
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(False)
        return True
    except Exception as e:
        print(f"❌ Failed to convert {rtf_path}: {e}")
        return False

# Convert each report RTF file
for title in titles:
    filename = title_to_filename(title)
    if filename:
        rtf_path = os.path.join(rtf_dir, filename)
        pdf_path = os.path.join(pdf_dir, filename.replace(".rtf", ".pdf"))
        if os.path.exists(rtf_path):
            success = convert_rtf_to_pdf(rtf_path, pdf_path)
            if success:
                print(f"Converted: {filename} → {os.path.basename(pdf_path)}")
        else:
            print(f"Missing RTF file: {filename}")
    else:
        print(f"Could not parse title: {title}")

# Convert the TOC RTF itself
toc_pdf_path = os.path.join(pdf_dir, os.path.basename(toc_rtf_path).replace(".rtf", ".pdf"))
convert_rtf_to_pdf(toc_rtf_path, toc_pdf_path)
print(f"\nTOC converted successfully: {toc_pdf_path}")

# Quit Word
word.Quit()
