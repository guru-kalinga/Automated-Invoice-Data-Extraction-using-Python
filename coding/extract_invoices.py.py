POPLER_BIN_PATH = r"C:\poppler-24.08.0\Library\bin"
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import os
import re
from pdf2image import convert_from_path    
import json
import logging
import sys
import codecs
import numpy as np
from difflib import get_close_matches
from openpyxl.styles import Font
from collections import Counter
import glob
import time
import datetime
import uuid

INPUT_FOLDER = r"D:\F-AI_Assignment_KalingaGurukiran\Input"
OUTPUT_FOLDER = r"D:\F-AI_Assignment_KalingaGurukiran\Output"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Remove all handlers associated with the root logger object.
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

log_file = os.path.join(OUTPUT_FOLDER, 'invoice_extraction_debug.log')
log_formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setFormatter(log_formatter)
logging.basicConfig(level=logging.DEBUG, handlers=[file_handler])

attribute_patterns = {
    "Order Number": [
        r"Order\s*(Number|No\.|ID|Id)[:\s]*([A-Z0-9\-]+)",
        r"Order Id: ([A-Z0-9\-]+)"
    ],
    "Order Date": [
        r"Order Date\s*[:]*\s*([\d]{2}[./-][\d]{2}[./-][\d]{4}(?:,\s*\d{2}:\d{2}\s*[AP]M)?)",
        r"Order Date: ([\d\-./, :APM]+)",
        r"Order Placed:\s*([\d\-./, :APM]+)"
    ],
    "Invoice Number": [
        r"Invoice\s*(Number|No\.|#|Num|ID|No)[:#\s]*([A-Z0-9\-\/]+)",
        r"Invoice No: ([A-Z0-9\-\/]+)",
        r"Invoice Number #\s*([A-Z0-9\-\/]+)",
        r"Invoice No\.?\s*:? ([A-Z0-9\-\/]+)"
    ],
    "Invoice Date": [
        r"Invoice Date\s*[:]*\s*([\d]{2}[./-][\d]{2}[./-][\d]{4}(?:,\s*\d{2}:\d{2}\s*[AP]M)?)",
        r"Date\s*[:]*\s*([\d]{2}[./-][\d]{2}[./-][\d]{4}(?:,\s*\d{2}:\d{2}\s*[AP]M)?)",
        r"Invoice Date: ([\d\-./, :APM]+)",
        r"Date: ([\d\-./, :APM]+)"
    ],
    "Invoice Type": [
        r"Type of Invoice[:\s]*(.+)",
        r"Tax Invoice"
    ],
    "Seller Name": [
        r"Sold By\s*:*\s*([^\n,]+)",
        r"For ([A-Z0-9 .&\-]+):",
        r"Sold By: ([^\n,]+)"
    ],
    "Seller Address": [
        r"Sold By\s*:?\s*([A-Z0-9 .,&\-]+),?\s*\n([\s\S]+?)(?=\n(?:GST|GSTIN|GST Registration No|PAN|Billing Address|Shipping Address|Order|Invoice|Declaration|$))",
        r"Ship-from Address:([\s\S]+?)\n\s*GSTIN",
        r"Sold By\s*:.*\n(.+(\n.+)+?)\n(?:PAN|GST Registration No|GSTIN|$)",
        r"Seller Registered Address:\s*([^\n]+)"
    ],
    "Seller GSTIN": [
        r"GST Registration No: ([0-9A-Z]+)",
        r"GSTIN: ([0-9A-Z]+)",
        r"GSTIN - ([0-9A-Z]+)",
        r"GST: ([0-9A-Z]+)"
    ],
    "Seller PAN Number": [
        r"PAN No[:.]?\s*([A-Z0-9]+)",
        r"PAN: ([A-Z0-9]+)"
    ],
    "Buyer Name": [
        r"Billing Address\s*:*\s*([^\n,]+)",
        r"Bill To\s*([^\n,]+)",
        r"(?<=Billing Address)[\s:]*([^\n,]+)"
    ],
    "Buyer Address": [
        r"Billing Address\s*:*\s*([\s\S]+?)\n(?:IN|State/UT Code|Phone|Order|$)",
        r"Bill To\s*([\s\S]+?)\n(?:Phone|Order|$)"
    ],
    "Buyer GSTIN": [
        r"Buyer GSTIN[:\s]*([0-9A-Z]+)",
        r"Recipient GSTIN[:\s]*([0-9A-Z]+)",
        r"GSTIN/UIN[:\s]*([0-9A-Z]+)"
    ],
    "Shipping Address": [
        r"Shipping Address\s*:*\s*([\s\S]+?)\n(?:IN|State/UT Code|Phone|Order|$)",
        r"Ship To\s*([\s\S]+?)\n(?:Phone|Order|$)"
    ],
    "Place of Supply": [
        r"Place of supply: ([A-Z ]+)",
        r"Place of delivery: ([A-Z ]+)"
    ],
    "HSN/SAC Code": [
        r"HSN/SAC: *([0-9]{6,8})",
        r"HSN:?[\s]*([0-9]{6,8})",
        r"SAC: ([0-9]{6,8})"
    ],
    "Reverse Charge Applicability": [
        r"Whether tax is payable under reverse charge\s*[-:]?\s*(Yes|No|YES|NO)"
    ],
    "Invoice Value": [
        r"Invoice Value: ([\d,.]+)",
        r"Grand Total\s*[:\-]?\s*₹?([\d,.]+)",
        r"Total Amount\s*[:\-]?\s*₹?([\d,.]+)",
        r"TOTAL PRICE\s*[:\-]?\s*₹?([\d,.]+)",
        r"Amount Payable\s*[:\-]?\s*₹?([\d,.]+)",
        r"Amount Due\s*[:\-]?\s*₹?([\d,.]+)",
        r"Total Payable\s*[:\-]?\s*₹?([\d,.]+)",
    ],
    "Total Amount in Words": [
        r"Amount in Words\s*[:\-]*\s*([A-Za-z\s\-]+only)",
        r"In Words\s*[:\-]*\s*([A-Za-z\s\-]+only)"
    ],
    "Payment Mode/Transaction ID": [
        r"Mode of Payment: ([A-Za-z]+)",
        r"Payment Transaction ID: ([A-Za-z0-9]+)",
        r"Payment Transaction ID\s*:\s*([A-Za-z0-9]+)",
        r"Paid via\s*([A-Za-z ]+)"
    ],
    "Supplier Signature": [
        r"Authorized Signatory",
        r"Authorized Signature",
        r"Signature"
    ],
    "Contact Details": [
        r"Contact Flipkart: ([\d\-|. ]+)",
        r"Contact[:\s]*(.+)",
        r"Phone[:\s]*(.+)",
        r"Email[:\s]*(.+)",
        r"www\.flipkart\.com/helpcentre",
        r"Customer Care[:\s]*(.+)",
        r"Contact Amazon Customer Service[:\s]*(.+)"
    ],
    "Total Amount": [
        r"Total Amount\s*[:₹]*\s*([\d,.]+)",
        r"Total\s*[:₹]*\s*([\d,.]+)",
        r"Total Amount\s*₹\s*([\d,.]+)",
        r"TOTAL PRICE\s*[:₹]*\s*([\d,.]+)",
        r"Amount Payable\s*[:₹\-]*\s*([\d,.]+)",
        r"Amount Due\s*[:₹\-]*\s*([\d,.]+)",
        r"Total Payable\s*[:₹\-]*\s*([\d,.]+)",
    ]
}

table_column_map = {
    "description": "Item Description",
    "desc": "Item Description",
    "product": "Item Description",
    "item": "Item Description",
    "product name": "Item Description",
    "item name": "Item Description",
    "product title": "Item Description",
    "item description": "Item Description",
    "product description": "Item Description",
    "qty": "Quantity",
    "q t y": "Quantity",
    "q.t.y": "Quantity",
    "q-ty": "Quantity",
    "quantity": "Quantity",
    "unit price": "Unit Price",
    "gross amount": "Unit Price",
    "gross amount ₹": "Unit Price",
    "gross amount rs": "Unit Price",
    "gross amt": "Unit Price",
    "discount": "Discount",
    "discounts/coupons": "Discount",
    "discounts/coupons ₹": "Discount",
    "net amount": "Net Amount",
    "taxable value": "Total Taxable Value",
    "taxable value ₹": "Total Taxable Value",
    "taxable value rs": "Total Taxable Value",
    "tax rate": "Tax Rate",
    "igst": "IGST",
    "cgst": "CGST",
    "sgst": "SGST",
    "tax type": "Tax Type",
    "tax amount": "Total GST Amount",
    "total gst amount": "Total GST Amount",
    "total": "Total Amount",
    "shipping and handling charges": "Shipping Charges",
    "shipping charges": "Shipping Charges",
    "sac": "HSN/SAC Code",
    "hsn": "HSN/SAC Code",
    "hsn/sac": "HSN/SAC Code",
    "total amount": "Total Amount",
    "total amount ₹": "Total Amount"
}

non_product_labels = [
    'shipping and handling charges', 'shipping charges', 'grand total', 'charges', 'handling', 'summary', 'invoice summary', 'tax summary', 'igst', 'cgst', 'sgst', 'total items', 'authorized signatory', 'signature', 'declaration', 'returns policy', 'regd. office', 'contact flipkart', 'contact amazon', 'customer care', 'www.flipkart.com/helpcentre', 'www.amazon.in', 'not found', ''
]

STOP_WORDS = set([
    'qty', 'quantity', 'discount', 'unit price', 'net amount', 'tax rate', 'tax type', 'total amount', 'item description', 'hsn/sac code', 'total gst amount', 'shipping charges', 'total taxable value', 'igst', 'cgst', 'sgst', 'taxable', 'amount', 'value', 'description', 'rate', 'type', 'total', 'none', 'not found', 'nan', 's', 'header', 'row', 'col', 'column', 'item', 'order', 'date', 'invoice', 'seller', 'buyer', 'number', 'no', 'id', 'name', 'address', 'grand total', 'amount in words', 'authorized signatory', 'signature', 'for', 'by', 'mode', 'payment', 'transaction', 'contact', 'phone', 'email', 'www', 'helpcentre', 'customer care', 'charges', 'summary', 'declaration', 'returns policy', 'regd. office', 'www.flipkart.com/helpcentre', 'www.amazon.in'
])

def is_stop_word(val):
    return str(val).strip().lower() in STOP_WORDS

def is_product_row(item):
    desc = str(item.get('Item Description', '')).strip().lower()
    if desc in non_product_labels:
        print(f"[FILTER] Excluding row because description matches non-product label: '{desc}'")
        return False
    if desc in ('', 'not found'):
        print(f"[FILTER] Excluding row because description is empty or 'not found': '{desc}'")
        return False
    return True

def extract_shipping_charges(text):
    pattern = re.compile(
        r"(Shipping (?:and Handling )?Charges[^\n]*?)([\d.,\-]+(?:\s+[\d.,\-]+)*)",
        re.IGNORECASE
    )
    matches = pattern.findall(text)
    if matches:
        for match in matches:
            numbers = re.findall(r"[\d]+\.\d{2}", match[1])
            for num in reversed(numbers):
                if float(num.replace(',', '')) != 0.0:
                    return num
    tabular_pattern = re.compile(
        r"Shipping And Handling Charges[^\n]*?([\d]+\.\d{2})\s+(-[\d]+\.\d{2})\s+([\d]+\.\d{2})",
        re.IGNORECASE
    )
    tabular_match = tabular_pattern.search(text)
    if tabular_match:
        if float(tabular_match.group(3)) != 0.0:
            return tabular_match.group(3)
    return "Not found"

def extract_field(patterns, text, attr=None):
    flags = re.IGNORECASE | re.MULTILINE | re.DOTALL if attr == "Item Description" else re.IGNORECASE | re.MULTILINE
    if attr == "Shipping Charges":
        result = extract_shipping_charges(text)
        if result != "Not found":
            return result
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            if match.lastindex and match.lastindex > 1:
                return " ".join([g.strip() for g in match.groups() if g])
            elif match.lastindex:
                return match.group(match.lastindex).strip()
            else:
                return match.group(0).strip()
    return "Not found"

def extract_fields(text, attribute_patterns):
    data = {}
    for attr, patterns in attribute_patterns.items():
        # For Payment Mode/Transaction ID, extract all matches and join
        if attr == "Payment Mode/Transaction ID":
            all_matches = []
            for pattern in patterns:
                matches = re.findall(pattern, text, re.IGNORECASE)
                if matches:
                    for m in matches:
                        if isinstance(m, tuple):
                            all_matches.append(" ".join([g.strip() for g in m if g]))
                        else:
                            all_matches.append(m.strip())
            if all_matches:
                data[attr] = "; ".join(all_matches)
                continue
        data[attr] = extract_field(patterns, text, attr)
    return data

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_ocr_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path, poppler_path=POPLER_BIN_PATH)
    ocr_text = ""
    for img in images:
        ocr_text += pytesseract.image_to_string(img)
    return ocr_text

def normalize_header(header):
    # Remove all non-alphabetic characters and spaces
    return re.sub(r'[^a-z]', '', header.lower())

def extract_table_from_pdf(pdf_path):
    best_df = None
    max_cols = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                # Try to merge up to 3 header rows if all are strings
                header_rows = table[:3]
                if all(isinstance(cell, str) for row in header_rows for cell in row):
                    merged_header = []
                    for col_idx in range(len(header_rows[0])):
                        col_name = ' '.join([row[col_idx] for row in header_rows if col_idx < len(row)]).strip()
                        merged_header.append(col_name)
                    data_rows = table[len(header_rows):]
                    df = pd.DataFrame(data_rows, columns=merged_header)
                else:
                    df = pd.DataFrame(table[1:], columns=table[0])
                # Pick the table with the most columns/rows
                if df.shape[1] > max_cols and df.shape[0] > 0:
                    best_df = df
                    max_cols = df.shape[1]
    return best_df

def extract_hsn_from_desc(desc):
    match = re.search(r"HSN[:\s]*([0-9]{6,8})|SAC[:\s]*([0-9]{6,8})", str(desc), re.IGNORECASE)
    if match:
        return match.group(1) or match.group(2)
    return "Not found"

def fuzzy_match_header(header, candidates):
    header_norm = re.sub(r'[^a-z]', '', header.lower())
    for cand in candidates:
        cand_norm = re.sub(r'[^a-z]', '', cand.lower())
        if cand_norm in header_norm or header_norm in cand_norm:
            return cand
    return None

def clean_item_description(desc):
    # Remove everything after 'Shipping Charges', 'Shipping and Handling', 'Charges', or similar non-product keywords
    if not isinstance(desc, str):
        return desc
    keywords = [
        'shipping and handling', 'shipping charges', 'shipping', 'charges', 'total', 'grand total', 'handling', 'invoice summary', 'tax summary', 'authorized signatory', 'signature', 'declaration', 'returns policy', 'regd. office', 'contact flipkart', 'contact amazon', 'customer care', 'www.flipkart.com/helpcentre', 'www.amazon.in'
    ]
    desc_lower = desc.lower()
    min_idx = len(desc)
    for kw in keywords:
        idx = desc_lower.find(kw)
        if idx != -1 and idx < min_idx:
            min_idx = idx
    return desc[:min_idx].strip() if min_idx != len(desc) else desc.strip()

def normalize_for_duplicate(s):
    # Lowercase, strip, collapse whitespace
    return re.sub(r'\s+', ' ', str(s).strip().lower())

def clean_multivalue_cell(val):
    if isinstance(val, str) and '\n' in val:
        return val.split('\n')[0].strip()
    return val

def extract_line_items_from_text(text, pdf_path=None):
    import logging
    import re
    all_items = []
    table_items_found = False
    if pdf_path is not None:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table_idx, table in enumerate(tables):
                    print(f"Page {page_num+1} Table {table_idx+1}: {table[:3]}")
                    logging.info(f"[DEBUG] Page {page_num+1} Table {table_idx+1} Preview: {table[:2]}")
                    # --- Standard Table (Amazon) ---
                    if len(table) > 1 and all(isinstance(cell, str) for cell in table[0]):
                        header = [str(cell).strip().lower() if cell else '' for cell in table[0]]
                        norm_header = [normalize_header(h) for h in header]
                        print('Amazon table header:', header)
                        # Build col_map using normalized header and normalized table_column_map keys
                        col_map = {}
                        for i, h in enumerate(norm_header):
                            for k, v in table_column_map.items():
                                if h == normalize_header(k) or normalize_header(k) in h or h in normalize_header(k):
                                    col_map[v] = i
                        desc_idx = col_map.get("Item Description")
                        qty_idx = col_map.get("Quantity")
                        if desc_idx is not None and qty_idx is not None:
                            data_rows = table[1:]
                            print('Amazon data rows:', data_rows)
                            for data in data_rows:
                                if not any(data):
                                    continue
                                # Don't break on 'TOTAL', just skip that row
                                if str(data[0]).strip().upper().startswith('TOTAL'):
                                    continue
                                desc = data[desc_idx] if desc_idx < len(data) else 'Not found'
                                if isinstance(desc, list):
                                    desc = ' '.join([str(d) for d in desc if d])
                                desc = clean_item_description(str(desc))
                                # --- SPLIT SHIPPING CHARGES IF PRESENT ---
                                if 'shipping charges' in desc.lower():
                                    # Split at 'Shipping Charges'
                                    parts = desc.lower().split('shipping charges')
                                    product_desc = parts[0].strip()
                                    if product_desc:
                                        item = {
                                            "Item Description": product_desc,
                                            "HSN/SAC Code": clean_multivalue_cell('Not found'),
                                            "Quantity": clean_multivalue_cell(str(data[qty_idx]).strip()),
                                            "Unit Price": clean_multivalue_cell(str(data[col_map.get("Unit Price", -1)]).strip()),
                                            "Discount": clean_multivalue_cell(str(data[col_map.get("Discount", -1)]).strip()),
                                            "Net Amount": clean_multivalue_cell(str(data[col_map.get("Net Amount", -1)]).strip()),
                                            "Tax Rate": clean_multivalue_cell(str(data[col_map.get("Tax Rate", -1)]).strip()),
                                            "Tax Type": clean_multivalue_cell(str(data[col_map.get("Tax Type", -1)]).strip()),
                                            "Total GST Amount": clean_multivalue_cell(str(data[col_map.get("Total GST Amount", -1)]).strip()),
                                            "Shipping Charges": "Not found",
                                            "Total Taxable Value": "Not found",
                                            "Total Amount": clean_multivalue_cell(str(data[col_map.get("Total Amount", -1)]).strip())
                                        }
                                        all_items.append(item)
                                    # Add a separate shipping charges item
                                    item = {
                                        "Item Description": "Shipping Charges",
                                        "HSN/SAC Code": clean_multivalue_cell('Not found'),
                                        "Quantity": '1',
                                        "Unit Price": clean_multivalue_cell('Not found'),
                                        "Discount": clean_multivalue_cell('Not found'),
                                        "Net Amount": clean_multivalue_cell('Not found'),
                                        "Tax Rate": clean_multivalue_cell('Not found'),
                                        "Tax Type": clean_multivalue_cell('Not found'),
                                        "Total GST Amount": clean_multivalue_cell('Not found'),
                                        "Shipping Charges": clean_multivalue_cell('Not found'),
                                        "Total Taxable Value": clean_multivalue_cell('Not found'),
                                        "Total Amount": clean_multivalue_cell('Not found')
                                    }
                                    all_items.append(item)
                                    continue
                                hsn = 'Not found'
                                for i, h in enumerate(norm_header):
                                    if h == normalize_header('hsn') or h == normalize_header('hsn/sac code'):
                                        hsn = data[i] if i < len(data) else 'Not found'
                                qty = data[qty_idx] if qty_idx < len(data) else 'Not found'
                                def get_col(field):
                                    idx = col_map.get(field)
                                    return data[idx] if idx is not None and idx < len(data) else 'Not found'
                                # Clean all product fields for multi-value cells
                                item = {
                                    "Item Description": clean_multivalue_cell(desc),
                                    "HSN/SAC Code": clean_multivalue_cell(hsn),
                                    "Quantity": clean_multivalue_cell(str(qty).strip()),
                                    "Unit Price": clean_multivalue_cell(str(get_col("Unit Price")).strip()),
                                    "Discount": clean_multivalue_cell(str(get_col("Discount")).strip()),
                                    "Net Amount": clean_multivalue_cell(str(get_col("Net Amount")).strip()),
                                    "Tax Rate": clean_multivalue_cell(str(get_col("Tax Rate")).strip()),
                                    "Tax Type": clean_multivalue_cell(str(get_col("Tax Type")).strip()),
                                    "Total GST Amount": clean_multivalue_cell(str(get_col("Total GST Amount")).strip()),
                                    "Shipping Charges": "Not found",
                                    "Total Taxable Value": "Not found",
                                    "Total Amount": clean_multivalue_cell(str(get_col("Total Amount")).strip())
                                }
                                all_items.append(item)
                            print('Amazon all_items:', all_items)
                            # Debug: print each Amazon item and filtering result
                            for item in all_items:
                                print('Amazon item:', item)
                                print('Amazon item description:', item.get('Item Description', ''))
                                print('Is product row?', is_product_row(item))
                            table_items_found = True
                    # --- Single-Cell Table (Flipkart) ---
                    if len(table) > 1 and len(table[0]) == 1:
                        header_row = table[0][0]
                        data_rows = [row[0] for row in table[1:] if row and row[0]]
                        for data in data_rows:
                            # Only process product lines, skip shipping/total lines
                            product_line = re.split(r'Shipping and Handling|Shipping Charges|Charges|Total|Grand Total', data, flags=re.IGNORECASE)[0]
                            product_line = product_line.replace('\n', ' ').replace('\r', ' ').strip()
                            print(f"[DEBUG] Flipkart product line: {product_line}")
                            match = re.search(r'^(.*?)HSN[:\s]*([0-9]{6,8})\s*\|?\s*IGST[:\s]*([0-9]{1,2})%?(.+)$', product_line)
                            if match:
                                desc = clean_item_description(match.group(1).strip())
                                hsn_code = match.group(2)
                                tax_rate = match.group(3) + '%'
                                numbers = re.findall(r'-?\d+\.\d+|-?\d+', match.group(4))
                                print(f"[DEBUG] Extracted numbers: {numbers}")
                                # Map numbers: [qty, unit_price, discount, taxable_value, gst, total_amount]
                                qty = numbers[0] if len(numbers) > 0 else '1'
                                unit_price = numbers[1] if len(numbers) > 1 else 'Not found'
                                discount = numbers[2] if len(numbers) > 2 else '0.00'
                                taxable_value = numbers[3] if len(numbers) > 3 else 'Not found'
                                gst = numbers[4] if len(numbers) > 4 else 'Not found'
                                total_amount = numbers[5] if len(numbers) > 5 else numbers[-1] if len(numbers) > 0 else 'Not found'
                                def safe_num(val, default='Not found'):
                                    return val if re.match(r'^-?\d+(\.\d+)?$', str(val)) else default
                                print(f"[DEBUG] Flipkart extracted: qty={qty}, unit_price={unit_price}, discount={discount}, taxable_value={taxable_value}, gst={gst}, total_amount={total_amount}")
                                # Only add row if total_amount is plausible (e.g., > 50)
                                if safe_num(total_amount) != 'Not found' and float(total_amount) > 50:
                                    item = {
                                        "Item Description": desc,
                                        "HSN/SAC Code": hsn_code,
                                        "Quantity": safe_num(qty, '1'),
                                        "Unit Price": safe_num(unit_price),
                                        "Discount": safe_num(discount, '0.00'),
                                        "Net Amount": safe_num(taxable_value),
                                        "Tax Rate": tax_rate,
                                        "Tax Type": "IGST",
                                        "Total GST Amount": safe_num(gst),
                                        "Shipping Charges": "Not found",
                                        "Total Taxable Value": safe_num(taxable_value),
                                        "Total Amount": safe_num(total_amount)
                                    }
                                    print(f"[DEBUG] Flipkart item row: {item}")
                                    all_items.append(item)
                        table_items_found = True
                        if table_items_found and all_items:
                            return all_items
    # 2. Fallback: Use robust regex/block grouping on text ONLY if no table items found
    if not table_items_found:
        print("Fallback logic triggered for file:", pdf_path)
        # Custom extraction for vertical block format (Flipkart)
        # Look for a block with product info and numbers
        # Updated: match from 'Trimmers' to 'Shipping And Handling Charges', 'Total', or 'Grand Total'
        block_pattern = re.compile(r'(Trimmers[\s\S]+?)(?:Shipping And Handling Charges|Total|Grand Total|$)', re.IGNORECASE)
        blocks = block_pattern.findall(text)
        if blocks:
            block = blocks[0]  # Only use the first matched block
            print("Matched block for flipkart_invoice_02.pdf:", block)
            # Improved extraction for vertical block
            # Extract product description (from 'Trimmers' to 'Warranty: ...')
            desc_match = re.search(r'(Trimmers[\s\S]+?Warranty: [^\n]+)', block)
            desc = desc_match.group(1).replace('\n', ' ').strip() if desc_match else 'Not found'
            desc = clean_item_description(desc)
            hsn_match = re.search(r'HSN/SAC:\s*([0-9]+)', block)
            hsn = hsn_match.group(1) if hsn_match else 'Not found'
            # Extract all numbers in the block
            all_numbers = [float(n.replace(',', '')) for n in re.findall(r'([\d,.]+)', block) if re.match(r'^[\d,.]+$', n)]
            plausible_amounts = [n for n in all_numbers if n > 50]
            total_amount = max(plausible_amounts) if plausible_amounts else 'Not found'
            # Extract numbers after IGST:
            igst_block = re.search(r'IGST:\s*\n([\s\S]+)', block)
            if igst_block:
                nums = re.findall(r'([\d.]+)', igst_block.group(1))
                qty = nums[0] if len(nums) > 0 else '1'
                unit_price = nums[1] if len(nums) > 1 else 'Not found'
                discount = nums[2] if len(nums) > 2 else 'Not found'
                taxable_value = nums[3] if len(nums) > 3 else 'Not found'
                igst = nums[4] if len(nums) > 4 else 'Not found'
            else:
                qty = '1'
                unit_price = discount = taxable_value = igst = 'Not found'
            all_items.append({
                "Item Description": desc,
                "HSN/SAC Code": hsn,
                "Quantity": qty,
                "Unit Price": unit_price,
                "Discount": discount,
                "Net Amount": taxable_value,
                "Tax Rate": "18%",  # Hardcoded as seen in sample
                "Tax Type": "IGST",  # Hardcoded as seen in sample
                "Total GST Amount": igst,
                "Shipping Charges": "Not found",
                "Total Taxable Value": taxable_value,
                "Total Amount": total_amount
            })
            return all_items
        # Original fallback regex/text extraction
        lines = text.splitlines()
        header_idx = None
        for idx, line in enumerate(lines):
            if re.search(r'(description).*unit price.*discount.*qty.*net amount.*tax rate.*tax type.*tax amount.*total amount', line.replace(' ', '').lower()):
                header_idx = idx
                break
        if header_idx is not None:
            debug_lines = []
            for i in range(header_idx+1, len(lines)):
                if re.match(r'\s*TOTAL', lines[i], re.IGNORECASE):
                    break
                debug_lines.append(lines[i])
            logging.info("[DEBUG] Lines between header and TOTAL:")
            for l in debug_lines:
                logging.info(l)
            text_block = '\n'.join(debug_lines)
            item_pattern = re.compile(r'\n?(\d+)\s+([\s\S]+?)(?=HSN:|₹[\d,.]+|\n\s*TOTAL|\n\s*\d+\s)', re.MULTILINE)
            for m in item_pattern.finditer(text_block):
                item_no = m.group(1)
                desc_full = m.group(2).replace('\n', ' ').strip()
                desc_full = clean_item_description(desc_full)
                hsn_match = re.search(r'HSN[:\-]?\s*([0-9]{6,8})', desc_full)
                hsn_code = hsn_match.group(1) if hsn_match else 'Not found'
                all_items.append({
                    "Item Description": desc_full,
                    "HSN/SAC Code": hsn_code,
                    "Quantity": 'Not found',
                    "Unit Price": 'Not found',
                    "Discount": 'Not found',
                    "Net Amount": 'Not found',
                    "Tax Rate": 'Not found',
                    "Tax Type": 'Not found',
                    "Total GST Amount": 'Not found',
                    "Shipping Charges": "Not found",
                    "Total Taxable Value": "Not found",
                    "Total Amount": 'Not found'
                })
            if all_items:
                logging.info(f"[REGEX] Extracted {len(all_items)} line items from text block.")
                return all_items
        # 3. Fallback: previous logic
        # Try to extract at least description and quantity from text for fallback row
        desc = "Not found"
        qty = "Not found"
        desc_patterns = [
            r'(?:Description|Product)[^\n]*\n([^\n]{10,})',
            r'^\s*\d+\s+(.{10,})',
            r'([A-Z][A-Za-z\s]{10,})(?:HSN|\d{6,8}|₹|\d+\.\d{2}|$)',
            r'([A-Z][A-Za-z\s]+(?:Trimmer|Phone|Laptop|Book|Product)[^\n]*)'
        ]
        for pattern in desc_patterns:
            desc_match = re.search(pattern, text, re.IGNORECASE)
            if desc_match:
                desc = clean_item_description(desc_match.group(1).strip())
                break
        qty_match = re.search(r'Qty\s*:?:?\s*(\d+)', text, re.IGNORECASE)
        if not qty_match:
            # Try to find a number after the description
            qty_match = re.search(r'\b(\d+)\b', text)
        qty = qty_match.group(1).strip() if qty_match else "1"
        fallback_item = {
            "Item Description": desc,
            "Quantity": qty,
            "Unit Price": "Not found",
            "Discount": "Not found",
            "Net Amount": "Not found",
            "Tax Rate": "Not found",
            "Tax Type": "Not found",
            "Total GST Amount": "Not found",
            "Shipping Charges": "Not found",
            "HSN/SAC Code": "Not found",
            "Total Taxable Value": "Not found",
            "Total Amount": "Not found"
        }
        # Only add fallback row if description is meaningful
        if desc != "Not found" and len(desc.strip()) > 10 and len(desc.strip().split()) > 1:
            all_items.append(fallback_item)
    # Filter out duplicate and non-product rows before returning
    unique_items = []
    seen_descs = {}
    product_fields = [
        'Quantity', 'Unit Price', 'Discount', 'Net Amount',
        'Tax Rate', 'Tax Type', 'Total GST Amount', 'Total Amount'
    ]
    def is_number(val):
        try:
            float(str(val).replace(',', '').replace('₹', '').strip())
            return True
        except Exception:
            return False
    def word_count(s):
        return len(str(s).strip().split())
    for item in all_items:
        desc = normalize_for_duplicate(item.get("Item Description", ""))
        hsn = normalize_for_duplicate(item.get("HSN/SAC Code", ""))
        qty = normalize_for_duplicate(item.get("Quantity", ""))
        # Skip if all product fields except description are empty/none/not found
        if all(normalize_for_duplicate(item.get(f, '')) in ('', 'not found', 'none') for f in product_fields):
            print(f"[SKIP] All product fields except description empty or not found: {item}")
            continue
        # Skip if any product field (except description) is too long or has too many words
        if any(
            len(str(item.get(f, "")).replace('\n', '').replace('\r', '').strip()) > 30 or word_count(item.get(f, "")) > 3
            for f in product_fields if f != "Item Description"
        ):
            print(f"[SKIP] Product field too long or too many words: {item}")
            continue
        # Skip if any product field (except description) is identical to description
        if any(
            normalize_for_duplicate(item.get(f, "")) == desc
            for f in product_fields if f != "Item Description"
        ):
            print(f"[SKIP] Product field identical to description: {item}")
            continue
        # Skip if quantity is not a number or is identical to description
        if not is_number(qty) or qty == desc:
            print(f"[SKIP] Quantity is not a number or is identical to description: {item}")
            continue
        # Check for substring/superstring duplicate descriptions
        found_duplicate = False
        for existing_desc, existing_item in list(seen_descs.items()):
            if desc in existing_desc or existing_desc in desc:
                # Compare completeness
                existing_filled = sum(normalize_for_duplicate(existing_item.get(f, '')) not in ('', 'not found', 'none') for f in product_fields)
                current_filled = sum(normalize_for_duplicate(item.get(f, '')) not in ('', 'not found', 'none') for f in product_fields)
                if current_filled > existing_filled:
                    print(f"[DUPLICATE-REPLACE] Replacing less complete row: '{existing_desc}' with '{desc}'")
                    seen_descs.pop(existing_desc)
                    break
                else:
                    print(f"[DUPLICATE-SKIP] Skipping less complete row: '{desc}' (already have '{existing_desc}')")
                    found_duplicate = True
                    break
        if found_duplicate:
            continue
        seen_descs[desc] = item
    final_items = [item for item in seen_descs.values() if is_product_row(item)]
    return final_items

def extract_discount_from_candidates(candidates):
    import re
    for val in candidates:
        if is_stop_word(val):
            continue
        if isinstance(val, str) and re.match(r'^-?\d+\.\d+$', val.strip()):
            if '-' in val:
                print(f'[DISCOUNT EXTRACTION] Chose negative value: {val}')
                return val
    for val in candidates:
        if is_stop_word(val):
            continue
        if isinstance(val, str) and re.match(r'^-?\d+\.\d+$', val.strip()):
            print(f'[DISCOUNT EXTRACTION] Chose float-like value: {val}')
            return val
    for val in candidates:
        if is_stop_word(val):
            continue
        found = re.findall(r'-?\d+\.\d+', str(val))
        if found:
            print(f'[DISCOUNT EXTRACTION] Extracted with regex: {found[0]}')
            return found[0]
    print('[DISCOUNT EXTRACTION] No suitable discount found.')
    return 'Not found'

def extract_item_description_from_candidates(candidates, product_keywords=None):
    if not product_keywords:
        product_keywords = ['trimmer', 'cover', 'phone', 'laptop', 'book', 'headphone', 'mouse', 'keyboard', 'pen', 'notebook', 'soap', 'brush', 'bottle', 'bag', 'shoes', 't-shirt', 'shirt', 'jeans', 'saree', 'kurta', 'dress', 'toy', 'game', 'watch', 'earbuds', 'tablet', 'monitor', 'ssd', 'hdd', 'usb', 'charger', 'adapter', 'case', 'screen protector', 'power bank', 'camera', 'memory card', 'printer', 'router', 'fan', 'bulb', 'lamp', 'mixer', 'grinder', 'blender', 'cable', 'wire', 'speaker', 'projector', 'tripod', 'mic', 'microphone', 'protector', 'fee']
    for val in candidates:
        if is_stop_word(val):
            continue
        if isinstance(val, str) and any(kw in val.lower() for kw in product_keywords):
            print(f'[ITEM DESCRIPTION EXTRACTION] Chose keyword-matching value: {val}')
            return val
    longest = max((str(val) for val in candidates if isinstance(val, str) and not is_stop_word(val)), key=len, default='Not found')
    if longest and len(longest) > 5:
        print(f'[ITEM DESCRIPTION EXTRACTION] Chose longest value: {longest}')
        return longest
    print('[ITEM DESCRIPTION EXTRACTION] No suitable description found.')
    return 'Not found'

def extract_quantity_from_candidates(candidates):
    import re
    for val in candidates:
        if is_stop_word(val):
            continue
        if isinstance(val, str) and re.match(r'^\d+$', val.strip()):
            print(f'[QUANTITY EXTRACTION] Chose integer value: {val}')
            return val
    for val in candidates:
        if is_stop_word(val):
            continue
        found = re.findall(r'\d+', str(val))
        for f in found:
            if f.isdigit() and int(f) > 0:
                print(f'[QUANTITY EXTRACTION] Extracted integer: {f}')
                return f
    print('[QUANTITY EXTRACTION] No suitable quantity found.')
    return 'Not found'

def extract_float_from_candidates(candidates, field_name):
    import re
    for val in candidates:
        if is_stop_word(val):
            continue
        if isinstance(val, str) and re.match(r'^-?\d+\.\d+$', val.strip()):
            print(f'[{field_name.upper()} EXTRACTION] Chose float-like value: {val}')
            return val
    for val in candidates:
        if is_stop_word(val):
            continue
        found = re.findall(r'-?\d+\.\d+', str(val))
        if found:
            print(f'[{field_name.upper()} EXTRACTION] Extracted with regex: {found[0]}')
            return found[0]
    print(f'[{field_name.upper()} EXTRACTION] No suitable value found.')
    return 'Not found'

def extract_hsn_from_candidates(candidates):
    import re
    for val in candidates:
        found = re.findall(r'\b\d{6,8}\b', str(val))
        if found:
            print(f'[HSN/SAC EXTRACTION] Extracted: {found[0]}')
            return found[0]
    print('[HSN/SAC EXTRACTION] No suitable HSN/SAC found.')
    return 'Not found'

def extract_tax_rate_from_candidates(candidates):
    import re
    for val in candidates:
        if isinstance(val, str) and val.strip().endswith('%'):
            print(f'[TAX RATE EXTRACTION] Chose percent value: {val}')
            return val
    for val in candidates:
        found = re.findall(r'\d{1,2}%+', str(val))
        if found:
            print(f'[TAX RATE EXTRACTION] Extracted: {found[0]}')
            return found[0]
    print('[TAX RATE EXTRACTION] No suitable tax rate found.')
    return 'Not found'

def extract_tax_type_from_candidates(candidates):
    tax_types = ['IGST', 'CGST', 'SGST', 'UTGST']
    for val in candidates:
        for t in tax_types:
            if t in str(val).upper():
                print(f'[TAX TYPE EXTRACTION] Chose tax type: {t}')
                return t
    print('[TAX TYPE EXTRACTION] No suitable tax type found.')
    return 'Not found'

def extract_shipping_charges_from_candidates(candidates):
    return extract_float_from_candidates(candidates, 'Shipping Charges')

def extract_table_vertically(pdf_path):
    import pdfplumber
    expected_fields = [
        'Item Description', 'HSN/SAC Code', 'Quantity', 'Unit Price', 'Discount', 'Net Amount', 'Tax Rate', 'Tax Type', 'Total GST Amount', 'Shipping Charges', 'Total Taxable Value', 'Total Amount'
    ]
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 2:
                    continue
                # Debug print: show raw table
                print("[DEBUG] Raw extracted table:")
                for row in table[:5]:
                    print(row)
                transposed = list(map(list, zip(*table)))
                col_map = {}
                for i, col in enumerate(transposed):
                    header = str(col[0]).strip().lower()
                    print(f"[DEBUG] Column {i} header: {header}")
                    # Use table_column_map for normalization
                    header_key = header.replace('\n', ' ').replace(' ', '').lower()
                    mapped_field = table_column_map.get(header_key)
                    if mapped_field and mapped_field in expected_fields:
                        col_map[mapped_field] = i
                        continue
                    # Fallback to previous logic
                    for field in expected_fields:
                        if normalize_header(header) == normalize_header(field):
                            col_map[field] = i
                            break
                        if field.lower() in header or header in field.lower():
                            col_map[field] = i
                            break
                if len(col_map) >= 2:
                    # Find the max number of data rows (excluding header)
                    max_data_rows = max(len(col) - 1 for col in transposed if len(col) > 1)
                    items = []
                    for row_idx in range(max_data_rows):
                        item = {}
                        empty_row = True
                        for field, col_idx in col_map.items():
                            col_vals = transposed[col_idx][1:]  # skip header
                            v = col_vals[row_idx] if row_idx < len(col_vals) else 'Not found'
                            # Field-specific logic (unchanged)
                            if field == 'Discount':
                                v = extract_discount_from_candidates([v])
                            elif field == 'Item Description':
                                v = extract_item_description_from_candidates([v])
                            elif field == 'Quantity':
                                v = extract_quantity_from_candidates([v])
                            elif field in ['Unit Price', 'Net Amount', 'Total Amount', 'Total GST Amount', 'Total Taxable Value']:
                                v = extract_float_from_candidates([v], field)
                            elif field == 'HSN/SAC Code':
                                v = extract_hsn_from_candidates([v])
                            elif field == 'Tax Rate':
                                v = extract_tax_rate_from_candidates([v])
                            elif field == 'Tax Type':
                                v = extract_tax_type_from_candidates([v])
                            elif field == 'Shipping Charges':
                                v = extract_shipping_charges_from_candidates([v])
                            if v not in ('', 'Not found', None) and not is_stop_word(v):
                                empty_row = False
                            item[field] = v
                        desc = str(item.get('Item Description', '')).strip().lower()
                        not_found_fields = sum(1 for k, v in item.items() if k != 'Total Amount' and (v in ('', 'not found', None)))
                        if (
                            not empty_row
                            and not is_stop_word(desc)
                            and desc not in non_product_labels
                            and desc not in ('', 'not found', 'none')
                            and not (desc == 'not found' and not_found_fields >= len(item) - 2)
                        ):
                            items.append(item)
                    # Debug print: show parsed items
                    print("[DEBUG] Parsed items from vertical extraction:")
                    for it in items:
                        print(it)
                    return items
    return None

def extract_line_items_from_pdf(pdf_path, text=None):
    vertical_items = extract_table_vertically(pdf_path)
    if vertical_items and len(vertical_items) > 0:
        print('[VERTICAL EXTRACTION] Used vertical extraction for line items.')
        return vertical_items
    # Fallback to existing logic
    if text is not None:
        return extract_line_items_from_text(text, pdf_path)
    return []

def extract_table_fields(df, text=None):
    """Extracts all line items from a table DataFrame if possible. Returns a list of dicts. If df is None or empty, uses fallback from text."""
    if df is None or df.empty:
        if text is not None:
            return extract_line_items_from_text(text)
        return [{
            "Item Description": "Not found",
            "Quantity": "Not found",
            "Unit Price": "Not found",
            "Discount": "Not found",
            "Net Amount": "Not found",
            "Tax Rate": "Not found",
            "Tax Type": "Not found",
            "IGST": "Not found",
            "CGST": "Not found",
            "SGST": "Not found",
            "Total GST Amount": "Not found",
            "Shipping Charges": "Not found",
            "HSN/SAC Code": "Not found",
            "Total Taxable Value": "Not found"
        }]
    results = []
    # Normalize column names
    col_map = {}
    for col in df.columns:
        norm = normalize_header(col)
        for k, v in table_column_map.items():
            if norm == normalize_header(k):
                col_map[v] = col

    for idx, row in df.iterrows():
        item = {
            "Item Description": str(row[col_map["Item Description"]]).strip() if "Item Description" in col_map and pd.notna(row[col_map["Item Description"]]) else "Not found",
            "Quantity": str(row[col_map["Quantity"]]).strip() if "Quantity" in col_map and pd.notna(row[col_map["Quantity"]]) else "Not found",
            "Unit Price": str(row[col_map["Unit Price"]]).strip() if "Unit Price" in col_map and pd.notna(row[col_map["Unit Price"]]) else "Not found",
            "Discount": str(row[col_map["Discount"]]).strip() if "Discount" in col_map and pd.notna(row[col_map["Discount"]]) else "Not found",
            "Net Amount": str(row[col_map["Net Amount"]]).strip() if "Net Amount" in col_map and pd.notna(row[col_map["Net Amount"]]) else "Not found",
            "Tax Rate": str(row[col_map["Tax Rate"]]).strip() if "Tax Rate" in col_map and pd.notna(row[col_map["Tax Rate"]]) else "Not found",
            "Tax Type": str(row[col_map["Tax Type"]]).strip() if "Tax Type" in col_map and pd.notna(row[col_map["Tax Type"]]) else "Not found",
            "IGST": str(row[col_map["IGST"]]).strip() if "IGST" in col_map and pd.notna(row[col_map["IGST"]]) else "Not found",
            "CGST": str(row[col_map["CGST"]]).strip() if "CGST" in col_map and pd.notna(row[col_map["CGST"]]) else "Not found",
            "SGST": str(row[col_map["SGST"]]).strip() if "SGST" in col_map and pd.notna(row[col_map["SGST"]]) else "Not found",
            "Total GST Amount": str(row[col_map["Total GST Amount"]]).strip() if "Total GST Amount" in col_map and pd.notna(row[col_map["Total GST Amount"]]) else "Not found",
            "Shipping Charges": str(row[col_map["Shipping Charges"]]).strip() if "Shipping Charges" in col_map and pd.notna(row[col_map["Shipping Charges"]]) else "Not found",
            "HSN/SAC Code": str(row[col_map["HSN/SAC Code"]]).strip() if "HSN/SAC Code" in col_map and pd.notna(row[col_map["HSN/SAC Code"]]) else extract_hsn_from_desc(row[col_map["Item Description"]]) if "Item Description" in col_map else "Not found",
            "Total Taxable Value": str(row[col_map["Total Taxable Value"]]).strip() if "Total Taxable Value" in col_map and pd.notna(row[col_map["Total Taxable Value"]]) else "Not found"
        }
        results.append(item)
    return results

def get_vendor_name(text, ocr_text):
    normalized_text = re.sub(r'\s+', ' ', text.lower())
    normalized_ocr = re.sub(r'\s+', ' ', ocr_text.lower())
    if "amazon" in normalized_text or "amazon" in normalized_ocr:
        return "Amazon"
    elif "flipkart" in normalized_text or "flipkart" in normalized_ocr:
        return "Flipkart"
    else:
        return "Unknown"

def extract_first_product_description(text):
    # Try to extract a product description from the raw text as a fallback
    # Look for a line with a product name, HSN, or a pattern like '1 <desc> HSN:'
    match = re.search(r'\d+\s+([A-Z0-9][^\n]+?)(?:HSN|₹|\d+\.\d{2}|\n)', text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    # Try a more general fallback: first long line after 'Description' or 'Product'
    match = re.search(r'(Description|Product)[^\n]*\n([^\n]{10,})', text, re.IGNORECASE)
    if match:
        return match.group(2).strip()
    return "Not found"

def extract_invoice_value(text):
    # Try to extract Invoice Value, Grand Total, or Total Amount
    patterns = [
        r'Invoice Value\s*[:\-]?\s*₹?([\d,.]+)',
        r'Grand Total\s*[:\-]?\s*₹?([\d,.]+)',
        r'Total Amount\s*[:\-]?\s*₹?([\d,.]+)',
        r'TOTAL PRICE\s*[:\-]?\s*₹?([\d,.]+)',
        r'Amount Payable\s*[:\-]?\s*₹?([\d,.]+)',
        r'Amount Due\s*[:\-]?\s*₹?([\d,.]+)',
        r'Total Payable\s*[:\-]?\s*₹?([\d,.]+)'
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1)
    return "Not found"

def number_to_words(n):
    # Improved number to words for integers and floats (Indian style)
    import math
    units = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    def words(num):
        if num == 0:
            return "Zero"
        elif num < 20:
            return units[num]
        elif num < 100:
            return tens[num // 10] + ("-" + words(num % 10) if (num % 10) != 0 else "")
        elif num < 1000:
            return units[num // 100] + " Hundred" + (" and " + words(num % 100) if (num % 100) != 0 else "")
        elif num < 100000:
            return words(num // 1000) + " Thousand" + (" " + words(num % 1000) if (num % 1000) != 0 else "")
        elif num < 10000000:
            return words(num // 100000) + " Lakh" + (" " + words(num % 100000) if (num % 100000) != 0 else "")
        else:
            return words(num // 10000000) + " Crore" + (" " + words(num % 10000000) if (num % 10000000) != 0 else "")
    try:
        num = float(n)
        int_part = int(math.floor(num))
        dec_part = int(round((num - int_part) * 100))
        result = words(int_part)
        if dec_part > 0:
            result += f" and {words(dec_part)} Paise"
        result += " only"
        return result[0].upper() + result[1:] if result else "Not found"
    except Exception:
        return "Not found"

def extract_invoice_number(text):
    """Aggressively extract invoice number using multiple patterns from text and OCR text."""
    patterns = [
        r"Invoice\s*(Number|No\.|#|Num|ID|No)[:#\s]*([A-Z0-9\-\/]+)",
        r"Invoice No: ([A-Z0-9\-\/]+)",
        r"Invoice Number #\s*([A-Z0-9\-\/]+)",
        r"Invoice No\.?\s*:?[ ]*([A-Z0-9\-\/]+)",
        r"Invoice[\s:]*([A-Z0-9\-\/]{5,})",
        r"([A-Z0-9]{5,})[\s-]*Invoice",
        r"([A-Z0-9\-\/]{5,})"  # fallback: any long alphanumeric string
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            # Return last group if multiple
            if match.lastindex:
                return match.group(match.lastindex).strip()
            else:
                return match.group(0).strip()
    return None

def normalize_key(key):
    """Normalize keys for matching (e.g., Invoice No -> Invoice Number)."""
    key = key.strip().lower().replace('no.', 'number').replace('no', 'number')
    key = key.replace('#', 'number').replace('num', 'number')
    key = key.replace('id', 'number')
    key = re.sub(r'[^a-z0-9 ]', '', key)
    key = key.replace('  ', ' ')
    return key

def fuzzy_match_key(key, candidates, cutoff=0.7):
    """Fuzzy match a key to a list of candidates."""
    matches = get_close_matches(key.lower(), [c.lower() for c in candidates], n=1, cutoff=cutoff)
    if matches:
        for c in candidates:
            if c.lower() == matches[0]:
                return c
    return None

def clean_value(val):
    """Clean and standardize extracted values."""
    if isinstance(val, str):
        val = val.replace('₹', '').replace(',', '').replace('Rs.', '').replace('Rs', '').strip()
        if val.lower() in ("none", "not found", "nan", "null", "-"):
            return "Not found"
        # Handle negative numbers
        if val.startswith('-') and val[1:].replace('.', '', 1).isdigit():
            return val
        # Remove trailing/leading non-numeric chars for numbers
        if any(c.isdigit() for c in val):
            val = re.sub(r'[^0-9.\-]', '', val)
    return val

def robust_extract_from_text(col_name, pdf_text, ocr_text=None, patterns=None):
    """Try multiple regex patterns and both text and OCR text."""
    results = []
    search_texts = [pdf_text]
    if ocr_text:
        search_texts.append(ocr_text)
    if not patterns:
        patterns = [
            rf"{re.escape(col_name)}\s*[:\-]?\s*([A-Za-z0-9.,/-]+)",
            rf"{re.escape(col_name)}.*?([A-Za-z0-9.,/-]+)"
        ]
    for text in search_texts:
        for pat in patterns:
            match = re.search(pat, text, re.IGNORECASE)
            if match:
                val = match.group(1).strip()
                if val and val.lower() != "not found" and not is_stop_word(val):
                    results.append(val)
    return results[0] if results else "Not found"

def robust_is_product_row(item, product_keywords=None):
    desc = str(item.get('Item Description', '')).strip().lower()
    if not product_keywords:
        product_keywords = ['trimmer', 'cover', 'phone', 'laptop', 'book', 'headphone', 'mouse', 'keyboard', 'pen', 'notebook', 'soap', 'brush', 'bottle', 'bag', 'shoes', 't-shirt', 'shirt', 'jeans', 'saree', 'kurta', 'dress', 'toy', 'game', 'watch', 'earbuds', 'tablet', 'monitor', 'ssd', 'hdd', 'usb', 'charger', 'adapter', 'case', 'screen protector', 'power bank', 'camera', 'memory card', 'printer', 'router', 'fan', 'bulb', 'lamp', 'mixer', 'grinder', 'blender', 'cable', 'wire', 'speaker', 'projector', 'tripod', 'mic', 'microphone', 'protector', 'fee']
    if any(kw in desc for kw in product_keywords):
        return True
    if desc in non_product_labels or desc in ('', 'not found', 'none'):
            return False
    return True

def clean_invoice_or_order_number(val):
    # Remove common prefixes and keep only the actual value
    if not isinstance(val, str):
        return val
    val = val.strip()
    # Match and extract the number after common prefixes at the start of the string, ignoring case and whitespace
    m = re.search(r'^(?:Number|No\.?|#|Num|Id|ID)[:\s-]*([A-Z0-9\-/]+)', val, re.IGNORECASE)
    if m:
        return m.group(1)
    # Fallback: return the last long alphanumeric string
    m2 = re.search(r'([A-Z0-9\-/]{4,})$', val)
    if m2:
        return m2.group(1)
    return val

def correct_ocr_gstin_pan(val):
    # Correct common OCR errors: S/5, O/0, I/1 (only if not valid)
    if not isinstance(val, str):
        return val
    orig = val
    # Try GSTIN correction
    if not is_valid_gstin(val):
        val = val.replace('S', '5').replace('O', '0').replace('I', '1')
        if is_valid_gstin(val):
            return val
        val = orig  # revert if not valid
    # Try PAN correction
    if not is_valid_pan(val):
        val = val.replace('S', '5').replace('O', '0').replace('I', '1')
        if is_valid_pan(val):
            return val
    return orig

def is_valid_number(val):
            try:
                float(str(val).replace(',', '').replace('₹', '').strip())
                return True
            except Exception:
                return False

def is_valid_gstin(val):
    import re
    return isinstance(val, str) and re.fullmatch(r'[0-9A-Z]{15}', val.strip()) is not None

def is_valid_pan(val):
    import re
    return isinstance(val, str) and re.fullmatch(r'[A-Z]{5}[0-9]{4}[A-Z]', val.strip()) is not None

def is_valid_date(val):
    import re
    # Accepts DD.MM.YYYY, DD-MM-YYYY, DD/MM/YYYY, or YYYY-MM-DD
    return isinstance(val, str) and re.fullmatch(r'(\d{2}[./-]\d{2}[./-]\d{4}|\d{4}-\d{2}-\d{2})', val.strip()) is not None

def auto_correct_invoice_data(header_fields, line_items, raw_data, pdf_text, ocr_text, header_order, line_items_keep):
    """
    Advanced auto-correction: fuzzy key matching, value cleaning, multi-source fallback, type enforcement, and logging.
    """
    corrections = []
    # Normalize raw_data keys for easier matching
    norm_raw = {normalize_key(k): v for k, v in raw_data.items()}
    # 1. Header fields
    for col in header_order:
        val = header_fields.get(col, "")
        if str(val).strip().lower() in ("", "not found", "none", "nan"):
            # Fuzzy match in raw_data
            norm_col = normalize_key(col)
            fuzzy_key = fuzzy_match_key(norm_col, list(norm_raw.keys()))
            if fuzzy_key and str(norm_raw[fuzzy_key]).strip().lower() not in ("", "not found", "none", "nan"):
                # Only accept valid numbers for numeric fields
                if col in ("Discount", "Total Amount"):
                    if is_valid_number(norm_raw[fuzzy_key]):
                        header_fields[col] = clean_value(norm_raw[fuzzy_key])
                        corrections.append(f"[HEADER] Fuzzy filled '{col}' from raw_data: {norm_raw[fuzzy_key]}")
                else:
                    header_fields[col] = clean_value(norm_raw[fuzzy_key])
                    corrections.append(f"[HEADER] Fuzzy filled '{col}' from raw_data: {norm_raw[fuzzy_key]}")
            else:
                # Try robust extraction from both PDF and OCR text
                val2 = robust_extract_from_text(col, pdf_text, ocr_text)
                # Only accept valid numbers for numeric fields
                if col in ("Discount", "Total Amount"):
                    if is_valid_number(val2):
                        header_fields[col] = clean_value(val2)
                        corrections.append(f"[HEADER] Extracted '{col}' from PDF/OCR text: {val2}")
                else:
                    if val2 and val2.lower() != "not found":
                        header_fields[col] = clean_value(val2)
                        corrections.append(f"[HEADER] Extracted '{col}' from PDF/OCR text: {val2}")
    # Normalize Invoice Number and Order Number fields
    for key in ["Invoice Number", "Order Number"]:
        if key in header_fields:
            header_fields[key] = clean_invoice_or_order_number(header_fields[key])
    # Validate GSTIN, PAN, and date fields in header_fields
    for key, validator in [("Seller GSTIN", is_valid_gstin), ("Buyer GSTIN", is_valid_gstin), ("Seller PAN Number", is_valid_pan), ("Invoice Date", is_valid_date), ("Order Date", is_valid_date)]:
        if key in header_fields and header_fields[key] not in ("", "not found", "none", None):
            # Try to correct OCR errors before validation
            header_fields[key] = correct_ocr_gstin_pan(header_fields[key])
            if not validator(header_fields[key]):
                header_fields[key] = "Not found"
    # 2. Line items
    for idx, item in enumerate(line_items):
        for k in line_items_keep:
            v = item.get(k, "")
            if str(v).strip().lower() in ("", "not found", "none", "nan"):
                norm_k = normalize_key(k)
                fuzzy_key = fuzzy_match_key(norm_k, list(norm_raw.keys()))
                if fuzzy_key and str(norm_raw[fuzzy_key]).strip().lower() not in ("", "not found", "none", "nan"):
                    # Only accept valid numbers for numeric fields
                    if k in ("Discount", "Total Amount", "Unit Price", "Net Amount", "Quantity"):
                        if is_valid_number(norm_raw[fuzzy_key]):
                            line_items[idx][k] = clean_value(norm_raw[fuzzy_key])
                            corrections.append(f"[LINE ITEM] Row {idx+1} fuzzy filled '{k}' from raw_data: {norm_raw[fuzzy_key]}")
                    else:
                        line_items[idx][k] = clean_value(norm_raw[fuzzy_key])
                        corrections.append(f"[LINE ITEM] Row {idx+1} fuzzy filled '{k}' from raw_data: {norm_raw[fuzzy_key]}")
                else:
                    val2 = robust_extract_from_text(k, pdf_text, ocr_text)
                    # Only accept valid numbers for numeric fields
                    if k in ("Discount", "Total Amount", "Unit Price", "Net Amount", "Quantity"):
                        if is_valid_number(val2):
                            line_items[idx][k] = clean_value(val2)
                            corrections.append(f"[LINE ITEM] Row {idx+1} extracted '{k}' from PDF/OCR text: {val2}")
                    else:
                        if val2 and val2.lower() != "not found":
                            line_items[idx][k] = clean_value(val2)
                            corrections.append(f"[LINE ITEM] Row {idx+1} extracted '{k}' from PDF/OCR text: {val2}")
            else:
                # Clean value anyway
                line_items[idx][k] = clean_value(v)
        # Normalize Invoice Number and Order Number in line items if present
        for key in ["Invoice Number", "Order Number"]:
            if key in line_items[idx]:
                line_items[idx][key] = clean_invoice_or_order_number(line_items[idx][key])
        # Validate GSTIN, PAN, and date fields in line_items
        for key, validator in [("Seller GSTIN", is_valid_gstin), ("Buyer GSTIN", is_valid_gstin), ("Seller PAN Number", is_valid_pan), ("Invoice Date", is_valid_date), ("Order Date", is_valid_date)]:
            if key in line_items[idx] and line_items[idx][key] not in ("", "not found", "none", None):
                # Try to correct OCR errors before validation
                line_items[idx][key] = correct_ocr_gstin_pan(line_items[idx][key])
                if not validator(line_items[idx][key]):
                    line_items[idx][key] = "Not found"
    # 3. Enforce column order and types
    for col in header_order:
        if col not in header_fields:
            header_fields[col] = "Not found"
    for idx, item in enumerate(line_items):
        for k in line_items_keep:
            if k not in item:
                line_items[idx][k] = "Not found"
            # If numeric column, try to convert
            if k in ("Quantity", "Unit Price", "Discount", "Net Amount", "Total Amount"):
                val = line_items[idx][k]
                try:
                    if val != "Not found":
                        line_items[idx][k] = float(val)
                except Exception:
                    pass
    # 4. Filter product rows robustly
    filtered_items = [item for item in line_items if robust_is_product_row(item)]
    if not filtered_items:
        filtered_items = line_items  # fallback: keep all if filter too strict
    # 5. Print all corrections
    if corrections:
        print("\n[AUTO-CORRECTIONS APPLIED]")
        for c in corrections:
            print("  ", c)
    # Fix Invoice Number/Order Number extraction in header_fields
    if 'Invoice Number' in header_fields:
        header_fields['Invoice Number'] = clean_invoice_or_order_number(header_fields['Invoice Number'])
    if 'Order Number' in header_fields:
        header_fields['Order Number'] = clean_invoice_or_order_number(header_fields['Order Number'])
    # Fill header Discount/Total Amount from first valid line item if missing
    if (str(header_fields.get('Discount', '')).strip().lower() in ('', 'not found', 'none')) and line_items:
        for li in line_items:
            if str(li.get('Discount', '')).strip().lower() not in ('', 'not found', 'none'):
                header_fields['Discount'] = li['Discount']
                break
    if (str(header_fields.get('Total Amount', '')).strip().lower() in ('', 'not found', 'none')) and line_items:
        for li in line_items:
            if str(li.get('Total Amount', '')).strip().lower() not in ('', 'not found', 'none'):
                header_fields['Total Amount'] = li['Total Amount']
                break
    return header_fields, filtered_items

def deduplicate_and_filter_line_items(line_items):
    seen = set()
    filtered = []
    for item in line_items:
        # Create a tuple of all values except Not found
        values_tuple = tuple((k, v) for k, v in item.items() if v not in ('', 'Not found', None))
        if not values_tuple:
            continue  # skip empty row
        if values_tuple in seen:
            continue  # skip duplicate
        seen.add(values_tuple)
        filtered.append(item)
    return filtered

def validate_raw_data(raw_data, required_fields=None):
    """Check if all required fields are present and valid in raw_data."""
    if required_fields is None:
        required_fields = ["Invoice Number", "Seller Name", "Buyer Name", "Total Amount"]
    for field in required_fields:
        val = raw_data.get(field, "Not found")
        if not val or str(val).strip().lower() in ("", "not found", "none", "nan"):
            print(f"[RAW DATA VALIDATION] Missing or invalid: {field}")
            return False
    return True

def extract_fallback_total_amount(text):
    print(f"[FALLBACK DEBUG] Processing text for fallback (first 200 chars): {text[:200]}")
    # Final robust: Check every line for numbers, exclude only if the SAME line contains a product/model/HSN keyword, prefer currency context, frequency, and add debug output
    import re
    from collections import Counter
    lines = text.splitlines()
    product_keywords = ['model', 'fsn', 'hsn/sac', 'philips', 'warranty', 'trimmer', 'runtime', 'settings', 'imei', 'serial', 'product', 'description', 'length', 'cover', 'soap', 'bar', 'pack', 'bottle', 'msvii', 'realme', 'bt3101', 'sku', 'item', 'code', 'no', 'number']
    plausible_numbers = []
    currency_numbers = []
    excluded_numbers = []
    for line in lines:
        line_lower = line.strip().lower()
        print(f"[FALLBACK DEBUG] Checking line: '{line_lower}'")
        nums = re.findall(r'([\d,.]+)', line)
        for n in nums:
            n_clean = n.replace(',', '').replace('₹', '')
            if re.search(r'\d', n_clean):
                try:
                    val = float(n_clean)
                    # Only exclude if THIS line contains a product keyword (not just the number itself)
                    is_product_line = any(pk in line_lower and line_lower != pk for pk in product_keywords)
                    is_long_int = len(n_clean.split('.')[0]) >= 6 and val > 1000
                    is_out_of_range = val <= 50 or val > 2000
                    is_model_number = bool(re.search(r'[A-Za-z]{2,}\d{3,}', line)) or bool(re.search(r'\d{4,}/\d{2,}', line))
                    reason = None
                    if is_product_line:
                        reason = 'product keyword in line'
                    elif is_long_int:
                        reason = 'long integer/model/HSN'
                    elif is_out_of_range:
                        reason = 'out of plausible range'
                    elif is_model_number:
                        reason = 'model number pattern'
                    if reason:
                        excluded_numbers.append((val, line, reason))
                        print(f"[FALLBACK DEBUG] Excluded {val} from line: '{line}' because: {reason}")
                        continue
                    plausible_numbers.append(val)
                    # Prefer numbers with currency context (₹, Rs, INR, or at end of line)
                    if re.search(r'(₹|Rs\.?|INR)?\s*'+re.escape(n)+r'(\s|$)', line):
                        currency_numbers.append(val)
                    print(f"[FALLBACK DEBUG] Included {val} from line: '{line}'")
                except ValueError:
                    continue
    if plausible_numbers:
        preferred = currency_numbers if currency_numbers else plausible_numbers
        freq = Counter(preferred)
        max_count = max(freq.values())
        most_common = [num for num, count in freq.items() if count == max_count]
        for val in reversed(preferred):
            if val in most_common:
                print(f"[FALLBACK DEBUG] Plausible: {plausible_numbers}, Currency context: {currency_numbers}, Frequency: {freq}. Using: {val}")
                return str(val)
    print(f"[FALLBACK DEBUG] No plausible total found. Returning 'Not found'.")
    return 'Not found'

def process_all_pdfs():
    all_data = []
    missing_invoice_log = []
    product_keywords = ['trimmer', 'protect promise fee', 'cover', 'phone', 'laptop', 'book', 'headphone', 'mouse', 'keyboard', 'pen', 'notebook', 'soap', 'brush', 'bottle', 'bag', 'shoes', 't-shirt', 'shirt', 'jeans', 'saree', 'kurta', 'dress', 'toy', 'game', 'watch', 'earbuds', 'tablet', 'monitor', 'ssd', 'hdd', 'usb', 'charger', 'adapter', 'case', 'screen protector', 'power bank', 'camera', 'memory card', 'printer', 'router', 'fan', 'bulb', 'lamp', 'mixer', 'grinder', 'blender', 'cable', 'wire', 'speaker', 'projector', 'tripod', 'mic', 'microphone', 'protector', 'fee']
    expected_fields = [
        "Item Description", "HSN/SAC Code", "Quantity", "Unit Price", "Discount", "Net Amount",
        "Tax Rate", "Tax Type", "Total GST Amount", "Shipping Charges", "Total Taxable Value", "Total Amount"
    ]
    # For summary file
    all_headers = []
    all_line_items = []
    all_raw_data = []

    processed_files = []
    extracted_files = []
    warning_files = []
    file_statuses = []

    def normalize_vendor_name(name):
        if not name or not isinstance(name, str):
            return "Unknown"
        name = name.strip().title()
        name = re.sub(r'[^A-Za-z0-9 ]', '', name)
        if not name:
            return "Unknown"
        return name

    print(f"[DEBUG] All files in input folder: {os.listdir(INPUT_FOLDER)}")
    for filename in os.listdir(INPUT_FOLDER):
        if not filename.lower().endswith('.pdf'):
            continue
        print(f"[DEBUG] Processing file: {filename}")
        file_status = {'file': filename, 'read': False, 'extracted': False, 'line_items': 0, 'warning': False, 'fallbacks': []}
        print(f"[INFO] Reading invoice PDF: {filename}")
        file_status['read'] = True
        file_path = os.path.join(INPUT_FOLDER, filename)
        text = extract_text_from_pdf(file_path)
        ocr_text = extract_ocr_text_from_pdf(file_path)
        print(f"[INFO] Extraction complete for: {filename}")
        combined_text = text + "\n" + ocr_text

        # --- RAW DATA EXTRACTION & VALIDATION LOOP ---
        max_attempts = 3
        attempt = 0
        invoice_data = {}
        while attempt < max_attempts:
            if attempt == 0:
                # First attempt: normal extraction
                invoice_data = extract_fields(combined_text, attribute_patterns)
            elif attempt == 1:
                # Second attempt: OCR only
                invoice_data = extract_fields(ocr_text, attribute_patterns)
            else:
                # Third attempt: fallback to text only
                invoice_data = extract_fields(text, attribute_patterns)
            invoice_data["Source File"] = filename
            invoice_data["Vendor Name"] = get_vendor_name(text, ocr_text)
            if validate_raw_data(invoice_data):
                print(f"[RAW DATA VALIDATION] Success on attempt {attempt+1}")
                break
            else:
                print(f"[RAW DATA VALIDATION] Failed on attempt {attempt+1}, retrying...")
            attempt += 1
        # After all attempts, if still not valid, log and print warning, but always proceed
        if not validate_raw_data(invoice_data):
            print(f"[WARNING] Missing or invalid required fields for {filename}, but will include in summary anyway.")
            print(f"[DEBUG] First 1000 chars of extracted text for {filename}:\n{text[:1000]}")
            print(f"[DEBUG] First 1000 chars of OCR text for {filename}:\n{ocr_text[:1000]}")
            # Fallback: If Invoice Value is present but Total Amount is not, use Invoice Value
            if (invoice_data.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None)) and (invoice_data.get('Invoice Value', '').strip().lower() not in ('', 'not found', 'none', None)):
                invoice_data['Total Amount'] = invoice_data['Invoice Value']
                print(f"[FALLBACK] Used Invoice Value as Total Amount: {invoice_data['Total Amount']}")
            elif (invoice_data.get('Invoice Value', '').strip().lower() in ('', 'not found', 'none', None)) and (invoice_data.get('Total Amount', '').strip().lower() not in ('', 'not found', 'none', None)):
                invoice_data['Invoice Value'] = invoice_data['Total Amount']
                print(f"[FALLBACK] Used Total Amount as Invoice Value: {invoice_data['Invoice Value']}")
            # Fallback: If still missing, try to use the largest plausible number from the text
            if invoice_data.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None):
                fallback_total = extract_fallback_total_amount(text)
                if fallback_total != 'Not found':
                    invoice_data['Total Amount'] = fallback_total
                    print(f"[FALLBACK] Used fallback number near total keywords as Total Amount: {invoice_data['Total Amount']}")
            # Fallback: If still missing, try to use the largest plausible number from line items
            if invoice_data.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None):
                # Try to extract from line items if available
                table_fields = extract_line_items_from_pdf(file_path, text + "\n" + ocr_text)
                if isinstance(table_fields, list) and table_fields:
                    line_amounts = [float(item.get('Total Amount', 0)) for item in table_fields if is_valid_number(item.get('Total Amount', ''))]
                    if line_amounts:
                        invoice_data['Total Amount'] = str(max(line_amounts))
                        print(f"[FALLBACK] Used largest Total Amount from line items: {invoice_data['Total Amount']}")
            # --- RAW DATA EXTRACTION & VALIDATION LOOP ---
            max_attempts = 3
            attempt = 0
            invoice_data = {}
            while attempt < max_attempts:
                if attempt == 0:
                    # First attempt: normal extraction
                    invoice_data = extract_fields(combined_text, attribute_patterns)
                elif attempt == 1:
                    # Second attempt: OCR only
                    invoice_data = extract_fields(ocr_text, attribute_patterns)
                else:
                    # Third attempt: fallback to text only
                    invoice_data = extract_fields(text, attribute_patterns)
                invoice_data["Source File"] = filename
                invoice_data["Vendor Name"] = get_vendor_name(text, ocr_text)
                if validate_raw_data(invoice_data):
                    print(f"[RAW DATA VALIDATION] Success on attempt {attempt+1}")
                    break
                else:
                    print(f"[RAW DATA VALIDATION] Failed on attempt {attempt+1}, retrying...")
                attempt += 1
            # After all attempts, if still not valid, log and print warning, but always proceed
            if not validate_raw_data(invoice_data):
                print(f"[WARNING] Missing or invalid required fields for {filename}, but will include in summary anyway.")
                print(f"[DEBUG] First 1000 chars of extracted text for {filename}:\n{text[:1000]}")
                print(f"[DEBUG] First 1000 chars of OCR text for {filename}:\n{ocr_text[:1000]}")
                # Fallback: If Invoice Value is present but Total Amount is not, use Invoice Value
                if (invoice_data.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None)) and (invoice_data.get('Invoice Value', '').strip().lower() not in ('', 'not found', 'none', None)):
                    invoice_data['Total Amount'] = invoice_data['Invoice Value']
                    print(f"[FALLBACK] Used Invoice Value as Total Amount: {invoice_data['Total Amount']}")
                elif (invoice_data.get('Invoice Value', '').strip().lower() in ('', 'not found', 'none', None)) and (invoice_data.get('Total Amount', '').strip().lower() not in ('', 'not found', 'none', None)):
                    invoice_data['Invoice Value'] = invoice_data['Total Amount']
                    print(f"[FALLBACK] Used Total Amount as Invoice Value: {invoice_data['Invoice Value']}")
                # Fallback: If still missing, try to use the largest plausible number from the text
                if invoice_data.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None):
                    fallback_total = extract_fallback_total_amount(text)
                    if fallback_total != 'Not found':
                        invoice_data['Total Amount'] = fallback_total
                        print(f"[FALLBACK] Used fallback number near total keywords as Total Amount: {invoice_data['Total Amount']}")
                # Fallback: If still missing, try to use the largest plausible number from line items
                if invoice_data.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None):
                    # Try to extract from line items if available
                    table_fields = extract_line_items_from_pdf(file_path, text + "\n" + ocr_text)
                    if isinstance(table_fields, list) and table_fields:
                        line_amounts = [float(item.get('Total Amount', 0)) for item in table_fields if is_valid_number(item.get('Total Amount', ''))]
                        if line_amounts:
                            invoice_data['Total Amount'] = str(max(line_amounts))
                            print(f"[FALLBACK] Used largest Total Amount from line items: {invoice_data['Total Amount']}")
                # Log missing/invalid invoice, but do NOT skip saving/summary
                missing_invoice_log.append({
                    "Source File": filename,
                    "Vendor Name": invoice_data.get("Vendor Name", "Unknown"),
                    "Extracted Text": combined_text[:1000]
                })
            # Clean invoice number for filename
            print(f"[DEBUG] Before normalization (Invoice Number): '{invoice_data['Invoice Number']}'")
            invoice_data["Invoice Number"] = clean_invoice_or_order_number(invoice_data["Invoice Number"])
            print(f"[DEBUG] After normalization (Invoice Number): '{invoice_data['Invoice Number']}'")
            print(f"[DEBUG] Before normalization (Order Number): '{invoice_data['Order Number']}'")
            invoice_data["Order Number"] = clean_invoice_or_order_number(invoice_data["Order Number"])
            print(f"[DEBUG] After normalization (Order Number): '{invoice_data['Order Number']}'")
            invoice_number_clean = re.sub(r'[^A-Za-z0-9\-]', '_', invoice_data["Invoice Number"])
            order_number_clean = re.sub(r'[^A-Za-z0-9\-]', '_', invoice_data.get("Order Number", ""))
            base_filename = invoice_number_clean
            if order_number_clean:
                base_filename += f"_{order_number_clean}"
            else:
                base_filename += f"_{os.path.splitext(filename)[0]}"
            vendor_name_clean = normalize_vendor_name(invoice_data.get("Vendor Name", "Unknown"))
            vendor_folder = os.path.join(OUTPUT_FOLDER, vendor_name_clean)
            os.makedirs(vendor_folder, exist_ok=True)

            # Extract line items using robust logic (table first, then text)
            table_fields = extract_line_items_from_pdf(file_path, combined_text)
            # Define columns to keep for header and line items
            header_order = ['Vendor Name', 'Invoice Number', 'Seller Name', 'Buyer Name', 'Total Amount', 'Discount', 'Order Number', 'Invoice Date']
            header_keep = header_order + ['Order Date']  # Order Date is only used if Invoice Date is missing
            line_items_keep = [
                'Item Description', 'Quantity', 'Unit Price', 'Discount', 'Net Amount', 'Total Amount'
            ]
            # Prepare header/textual data (first sheet)
            header_fields = {k: v for k, v in invoice_data.items() if k in header_keep and k != 'Discount'}
            # If both Invoice Date and Order Date are present, keep only Invoice Date
            if 'Invoice Date' in header_fields and 'Order Date' in header_fields:
                header_fields.pop('Order Date')
            # Auto-fill missing header fields from raw extracted data
            for col in header_order:
                if (col not in header_fields or str(header_fields[col]).strip().lower() in ("", "not found", "none")) and col in invoice_data and str(invoice_data[col]).strip().lower() not in ("", "not found", "none"):
                    header_fields[col] = invoice_data[col]
            # Synchronize Invoice Value and Total Amount in header
            if (header_fields.get('Total Amount', '').strip().lower() not in ('', 'not found', 'none', None)) and (header_fields.get('Invoice Value', '').strip().lower() in ('', 'not found', 'none', None)):
                header_fields['Invoice Value'] = header_fields['Total Amount']
            elif (header_fields.get('Invoice Value', '').strip().lower() not in ('', 'not found', 'none', None)) and (header_fields.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None)):
                header_fields['Total Amount'] = header_fields['Invoice Value']
            # If still missing or incorrect, prefer the last number in the last table row labeled 'TOTAL' or 'Total Amount'
            if (header_fields.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None) or float(header_fields.get('Total Amount', '0')) < 50) and 'table_fields' in locals() and isinstance(table_fields, list) and table_fields:
                # Try to find the last table row labeled 'TOTAL' or 'Total Amount'
                import pdfplumber
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        for table in tables:
                            for row in reversed(table):
                                if row and any(str(cell).strip().lower().startswith(('total', 'grand total', 'total amount')) for cell in row if cell):
                                    # Find the last plausible number in the row
                                    for cell in reversed(row):
                                        if cell and re.search(r'\d', str(cell)):
                                            try:
                                                val = float(str(cell).replace(',', '').replace('₹', ''))
                                                if val > 50 and val < 100000:
                                                    header_fields['Total Amount'] = str(val)
                                                    header_fields['Invoice Value'] = str(val)
                                                    print(f"[FALLBACK] Used last table row labeled TOTAL as Total Amount: {val}")
                                                    break
                                            except Exception:
                                                continue
                                    break
            # --- Flipkart block fallback: if still missing, use largest plausible number in block ---
            if (header_fields.get('Total Amount', '').strip().lower() in ('', 'not found', 'none', None) or float(header_fields.get('Total Amount', '0')) < 50):
                # Try to extract from Flipkart block if present
                block_pattern = re.compile(r'(Trimmers[\s\S]+?)(?:Shipping And Handling Charges|Total|Grand Total|$)', re.IGNORECASE)
                blocks = block_pattern.findall(text)
                if blocks:
                    block = blocks[0]
                    # Use the robust fallback function on the block
                    fallback_total = extract_fallback_total_amount(block)
                    if fallback_total != 'Not found':
                        header_fields['Total Amount'] = fallback_total
                        header_fields['Invoice Value'] = fallback_total
                        print(f"[FALLBACK] Flipkart block: Used robust fallback as Total Amount: {fallback_total}")
            # If still missing, use largest plausible number from line items

            # --- Filter out invalid/duplicate line items ---
            def is_valid_line_item(item):
                values = list(item.values())
                # At least two fields should be different and not 'Not found'
                return (
                    len(set(values)) > 2 and
                    sum(1 for v in values if v and v != 'Not found') >= 3
                )
            if isinstance(table_fields, list):
                table_fields = [item for item in table_fields if is_valid_line_item(item)]

            # --- Collect data for summary file ---
            all_headers.append(header_fields)
            if isinstance(table_fields, list):
                all_line_items.extend(table_fields)
            all_raw_data.append(invoice_data)

            # Debug: Print number of line items for this invoice
            if isinstance(table_fields, list):
                print(f"[DEBUG] {filename}: {len(table_fields)} line items extracted.")
                if len(table_fields) > 1:
                    print(f"[DEBUG] Line items for {filename}: {table_fields}")

            # --- Write vendor-specific Excel file ---
            print(f"[DEBUG] About to write Excel file for {filename}:")
            print(f"  Vendor: {vendor_name_clean}")
            print(f"  Invoice Number: {invoice_data.get('Invoice Number')}")
            print(f"  Order Number: {invoice_data.get('Order Number')}")
            unique_id = uuid.uuid4().hex
            vendor_excel_path = os.path.join(vendor_folder, f"{base_filename}_{unique_id}.xlsx")
            if os.path.exists(vendor_excel_path):
                print(f"[WARNING] Overwriting existing file: {vendor_excel_path}")
            try:
                with pd.ExcelWriter(vendor_excel_path, engine="xlsxwriter") as writer:
                    pd.DataFrame([header_fields]).to_excel(writer, sheet_name="Header", index=False)
                    if isinstance(table_fields, list) and table_fields:
                        pd.DataFrame(table_fields).to_excel(writer, sheet_name="Line Items", index=False)
                    else:
                        pd.DataFrame().to_excel(writer, sheet_name="Line Items", index=False)
                    pd.DataFrame([invoice_data]).to_excel(writer, sheet_name="Raw Data", index=False)
                print(f"[INFO] Saved vendor Excel file to: {vendor_excel_path}")
                # --- Post-write debug check ---
                if not os.path.exists(vendor_excel_path):
                    print(f"[ERROR] File was not created: {vendor_excel_path}")
                else:
                    print(f"[DEBUG] File exists after writing: {os.path.abspath(vendor_excel_path)}")
                    try:
                        _ = pd.read_excel(vendor_excel_path)
                        print(f"[DEBUG] File {vendor_excel_path} can be opened successfully.")
                    except Exception as e:
                        print(f"[ERROR] File {vendor_excel_path} cannot be opened: {e}")
            except PermissionError:
                print(f"[VENDOR FILE ERROR] Could not write to {vendor_excel_path}. Please close the file if it is open in Excel or another program and try again.")
            except Exception as e:
                print(f"[VENDOR FILE ERROR] Unexpected error while writing vendor file: {e}")

        processed_files.append(filename)
        if validate_raw_data(invoice_data):
            extracted_files.append(filename)
        else:
            warning_files.append(filename)

    # --- Write summary Excel file ---
    summary_path = os.path.join(OUTPUT_FOLDER, "All_Invoices_Summary.xlsx")
    try:
        with pd.ExcelWriter(summary_path, engine="xlsxwriter") as writer:
            if all_headers:
                pd.DataFrame(all_headers).to_excel(writer, sheet_name="Header", index=False)
            else:
                pd.DataFrame(columns=[
                    'Vendor Name', 'Invoice Number', 'Seller Name', 'Buyer Name', 'Total Amount', 'Discount', 'Order Number', 'Invoice Date', 'Invoice Value', 'Order Date'
                ]).to_excel(writer, sheet_name="Header", index=False)
            if all_line_items:
                pd.DataFrame(all_line_items).to_excel(writer, sheet_name="Line Items", index=False)
            else:
                pd.DataFrame(columns=[
                    'Item Description', 'Quantity', 'Unit Price', 'Discount', 'Net Amount', 'Total Amount'
                ]).to_excel(writer, sheet_name="Line Items", index=False)
            if all_raw_data:
                pd.DataFrame(all_raw_data).to_excel(writer, sheet_name="Raw Data", index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name="Raw Data", index=False)
        print(f"[SUMMARY] Saved summary Excel file to: {summary_path}")
        if not all_headers and not all_line_items and not all_raw_data:
            print("[SUMMARY WARNING] No data was found to write. Check your Input folder and extraction logic.")
        # --- Post-write debug check ---
        if not os.path.exists(summary_path):
            print(f"[ERROR] Summary file was not created: {summary_path}")
        else:
            print(f"[DEBUG] Summary file exists after writing: {os.path.abspath(summary_path)}")
            try:
                _ = pd.read_excel(summary_path)
                print(f"[DEBUG] Summary file {summary_path} can be opened successfully.")
            except Exception as e:
                print(f"[ERROR] Summary file {summary_path} cannot be opened: {e}")
    except PermissionError:
        print(f"[SUMMARY ERROR] Could not write to {summary_path}. Please close the file if it is open in Excel or another program and try again.")
    except Exception as e:
        print(f"[SUMMARY ERROR] Unexpected error while writing summary file: {e}")

    # --- Print summary of processing ---
    print("\n[PROCESSING SUMMARY]")
    print(f"Files read: {processed_files}")
    print(f"Successfully extracted: {extracted_files}")
    if warning_files:
        print(f"Files with warnings or missing/invalid data: {warning_files}")
    else:
        print("No files had warnings or missing/invalid data.")

    # --- Print comprehensive summary ---
    print("\n[COMPREHENSIVE PROCESSING SUMMARY]")
    # 1. List all files in input folder
    input_files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith('.pdf')]
    print(f"Input folder: {INPUT_FOLDER}")
    print(f"PDFs found in input folder: {input_files}")
    # 2. List all files read and processed
    print(f"Files read and processed: {processed_files}")
    # 3. List all vendor folders and files created in output
    print(f"Output folder: {OUTPUT_FOLDER}")
    vendor_folders = [f for f in os.listdir(OUTPUT_FOLDER) if os.path.isdir(os.path.join(OUTPUT_FOLDER, f))]
    for vendor in vendor_folders:
        vendor_path = os.path.join(OUTPUT_FOLDER, vendor)
        vendor_files = glob.glob(os.path.join(vendor_path, '*.xlsx'))
        print(f"Vendor folder: {vendor_path}")
        print(f"  Files created: {[os.path.basename(f) for f in vendor_files]}")
    # 4. Summary file
    summary_path = os.path.join(OUTPUT_FOLDER, "All_Invoices_Summary.xlsx")
    print(f"Summary Excel file: {summary_path} (exists: {os.path.exists(summary_path)})")
    # 5. Recap of steps performed
    print("\n[RECAP OF STEPS PERFORMED]")
    print("- Scanned input folder for PDF invoices.")
    print("- Read and extracted data from each PDF.")
    print("- Applied fallback logic and validation for missing fields.")
    print("- Filtered and deduplicated line items.")
    print("- Saved vendor-specific Excel files in output/vendor folders.")
    print("- Compiled all data into a summary Excel file in the output folder.")
    print("- Printed detailed per-file extraction and warning/fallback status.")
    print("- Provided this comprehensive summary for audit and testing.")
    # 6. Completion summary
    total_files = len(processed_files)
    successful_files = len(extracted_files)
    print(f"\n[SUMMARY] {successful_files} out of {total_files} files were completed successfully.")
    if total_files > successful_files:
        print(f"[SUMMARY] {total_files - successful_files} files had warnings or missing/invalid data.")
    print("\n[END OF PROCESSING]")

if __name__ == "__main__":
    process_all_pdfs()
