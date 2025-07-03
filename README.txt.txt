Name: Kalinga Gurukiran
Task: Invoice Extraction for F-AI Internship Assignment

Description:
This Python script extracts structured data from 4 PDF invoices (2 from Amazon, 2 from Flipkart)
and converts them into a clean Excel file for further data processing and analysis.

Folder Structure:
- Input/: Contains 4 sample invoice PDFs
- Output/: All_Invoices_Extracted.xlsx (cleaned and formatted invoice data), invoice_extraction_debug.xlsx (raw/intermediate extracted data)
- Coding/: Contains Python script (extract_invoices.py)

Input:
- amazon_invoice_01.pdf
- amazon_invoice_02.pdf
- flipkart_invoice_01.pdf
- flipkart_invoice_02.pdf

Output:
- All_Invoices_Extracted.xlsx: Final cleaned and formatted data
- invoice_extraction_debug.xlsx: Intermediate or raw extraction results for validation

Tools Used:
- Python (with PyPDF2/pdfminer/pdf extraction code)
- No third-party APIs used as per instructions

Note:
The script is written to handle different invoice layouts and extract only required structured fields.
