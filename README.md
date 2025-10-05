# Tender Eligibility Extractor

This script scans tender documents in `D:\Tenders`, extracts eligibility criteria using OCR and fuzzy matching, and exports results to Excel.

## How to Run
1. Install required Python libraries:
   - PyMuPDF
   - pytesseract
   - pdf2image
   - openpyxl
   - fuzzywuzzy
2. Place tender files in `D:\Tenders`
3. Run `extractscript.py`
4. Excel file will be saved in the same folder
