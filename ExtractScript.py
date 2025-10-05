import os
import re
import fitz  # PyMuPDF
from docx import Document
from fuzzywuzzy import fuzz
import openpyxl
from openpyxl.styles import Alignment

# Step 1: Setup folder and keywords
folder = r"D:\Tenders"
keywords = [
    "eligibility", "the firm should", "the applicant must", "must have",
    "required experience", "qualification", "empanelment"
]

# Step 2: Define firm profile
firm_profile = {
    "CAG empanelled": True,
    "RBI Category II": True,
    "Statutory audit": True,
    "Internal audit": True,
    "System audit": True,
    "ICFR": True,
    "DISA": True,
    "CISA": True,
    "FAFD": True,
    "NBFC experience": True,
    "Turnover above 2 crore": True
}

# Step 3: Clean illegal characters for Excel
def clean_text(text):
    return re.sub(r"[\x00-\x1F\x7F-\x9F]", "", text)

# Step 4: Extract text from PDF
def extract_pdf_text(path):
    try:
        doc = fitz.open(path)
        return "\n".join(page.get_text() for page in doc)
    except:
        return ""

# Step 5: Extract text from DOC/DOCX
def extract_doc_text(path):
    try:
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    except:
        return ""

# Step 6: Extract eligibility section
def extract_eligibility(text):
    lines = text.split("\n")
    eligibility_blocks = []
    for i, line in enumerate(lines):
        if any(kw in line.lower() for kw in keywords):
            block = []
            for j in range(i, min(i+15, len(lines))):
                block.append(lines[j])
                if lines[j].strip() == "":
                    break
            eligibility_blocks.append("\n".join(block))
    return "\n\n".join(eligibility_blocks)

# Step 7: Match profile
def match_profile(eligibility_text):
    score = 0
    for trait in firm_profile:
        if fuzz.partial_ratio(trait.lower(), eligibility_text.lower()) > 70:
            score += 1
    if score >= 6:
        return "Eligible"
    elif score >= 4:
        return "Partially Eligible"
    else:
        return "Not Eligible"

# Step 8: Loop through files
results = []
for file in os.listdir(folder):
    path = os.path.join(folder, file)
    if file.endswith(".pdf"):
        text = extract_pdf_text(path)
    elif file.endswith(".doc") or file.endswith(".docx"):
        text = extract_doc_text(path)
    else:
        continue

    eligibility = extract_eligibility(text)
    match = match_profile(eligibility)

    results.append({
        "File": file,
        "Match Status": match,
        "Eligibility Summary": clean_text(eligibility[:300] + "…") if eligibility else "Not found"
    })

# Step 9: Create Excel workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Tender Eligibility"

# Step 10: Write headers
headers = ["File Name", "Match Status", "Eligibility Summary"]
ws.append(headers)

# Step 11: Write data rows
for r in results:
    ws.append([
        clean_text(r["File"]),
        clean_text(r["Match Status"]),
        clean_text(r["Eligibility Summary"])
    ])

# Step 12: Format columns
for col in ws.columns:
    for cell in col:
        cell.alignment = Alignment(wrap_text=True)

# Step 13: Save workbook
excel_path = os.path.join(folder, "Eligible_Tenders.xlsx")
try:
    wb.save(excel_path)
    print(f"\n📊 Excel file saved to: {excel_path}")
except PermissionError:
    print(f"\n❌ Permission denied. Please close '{excel_path}' and re-run the script.")
