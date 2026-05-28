import os
import json
import io
import time  
import re
from datetime import datetime
import pandas as pd
from google import genai  
from google.genai import types  
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st
import pypdf  # Requirement: pip install pypdf

# ==========================================
# APP CONFIGURATION & UI SETUP (ALL ENGLISH)
# ==========================================
st.set_page_config(page_title="PO Checker AI", layout="wide")
st.title("📦 Purchase Order Checking Assistant")
st.write("Upload your System Master Data and PO PDFs to automatically generate a flagged discrepancy report with matching confirmation notes (Supports Split Delivery Schedules).")

# Securely fetch API key
if "GOOGLE_API_KEY" in st.secrets:
    GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
else:
    st.error("🔑 Google API Key is not set! Please configure it within Streamlit Secrets.")
    st.stop()

client = genai.Client(api_key=GOOGLE_API_KEY)
OUTPUT_FILENAME = "PO_Checking_Report.xlsx"

# Advanced matching cleanup helper to handle revision suffixes (e.g., rev01, rev03)
def clean_key(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).replace(" ", "").replace("-", "").replace(".", "").lower().strip()
    if 'rev' in s:
        s = s.split('rev')[0].strip()
    return s

# Standardize line numbers
def clean_line_num(val):
    if pd.isna(val) or val is None: return ""
    return str(val).strip().split('-')[0].strip().lstrip('0')

# Robust parsing tool to strip extra zeros (,0000) and isolate units (ST, pack, EA)
def parse_qty_and_unit(v):
    if pd.isna(v) or v is None:
        return None, None
    s = str(v).strip()
    
    match = re.match(r'^([\d\s.,]+)(.*)$', s)
    if match:
        num_part = match.group(1).strip()
        unit_part = match.group(2).strip()
        
        if ',' in num_part and '.' not in num_part:
            parts = num_part.split(',')
            if len(parts) == 2 and parts[1] == '0000':
                num_part = parts[0]
            else:
                num_part = num_part.replace(',', '.')
        elif '.' in num_part and ',' not in num_part:
            parts = num_part.split('.')
            if len(parts) == 2 and parts[1] == '0000':
                num_part = parts[0]
                
        num_part = num_part.replace(' ', '')
        
        try:
            val = float(num_part)
            if val.is_integer():
                val = int(val)
            return val, unit_part if unit_part else None
        except:
            return None, None
    return None, None

# Dynamic Price Normalizer handling 'à', 'per X', '€', and multlingual variations
def normalize_price(v):
    if pd.isna(v) or v is None: return None
    s = str(v).lower().strip()
    s = s.replace('à', '')
    
    factor = 1.0
    match_factor = re.search(r'(?:per|/)\s*(\d+)', s)
    if match_factor:
        factor = float(match_factor.group(1))
        
    s = re.sub(r'(?:per|/)\s*\d+\s*[a-z]*', '', s)
    
    for text_to_remove in ['st.', 'st', '€', 'eur', 'piece', 'pcs', 'ea', 'pack']:
        s = s.replace(text_to_remove, '')
    s = s.strip()
    
    if not s: return None
    if s.startswith(','): s = '0' + s
        
    s = s.replace(' ', '')
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
        else: s = s.replace(',', '')
    elif ',' in s and '.' not in s:
        s = s.replace(',', '.')
        
    try:
        unit_val = float(s)
        return round(unit_val / factor, 4)
    except:
        return None

# Smart date parser supporting dot (.), slash (/), and hyphen (-) formats natively
def parse_date_to_custom_format(v):
    if pd.isna(v) or v is None: return ""
    s = str(v).replace(':', '').replace('Delivery Date', '').strip()
    s = s.replace('.', '/').replace('-', '/')  
    s = s.split(',')[0].strip().split()[0]
    
    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%m/%d/%Y'):
        try: 
            return pd.to_datetime(s, format=fmt).strftime('%d/%m/%Y')
        except: 
            continue
    return str(v).strip()

# ==========================================
# UI FILE UPLOADERS
# ==========================================
excel_file = st.file_uploader("👉 Step 1: Upload System Master Data (Excel or CSV)", type=["xlsx", "xls", "csv"], key="master_data_excel_csv")
pdf_files = st.file_uploader("👉 Step 2: Upload PO PDF file(s)", type=["pdf"], accept_multiple_files=True, key="po_pdf_files_list")

# ==========================================
# APP PROCESSING LOGIC
# ==========================================
if excel_file and pdf_files:
    if st.button("🚀 Run AI Verification Report"):
        
        if excel_file.name.lower().endswith('.csv'):
            df_excel = pd.read_csv(excel_file, dtype=str)
        else:
            df_excel = pd.read_excel(excel_file, dtype=str)
            
        df_excel.columns = df_excel.columns.astype(str).str.strip()

        item_col_name = next((c for c in df_excel.columns if c.lower() in ['item', 'item number', 'material', 'part no', 'part number']), 'Item')
        if item_col_name not in df_excel.columns and len(df_excel.columns) > 0:
            item_col_name = df_excel.columns[0]

        unnamed_col = next((c for c in df_excel.columns if 'Unnamed' in c), 'Description_Extracted')
        if unnamed_col not in df_excel.columns:
            df_excel[unnamed_col] = None

        # NEW SCHEMA: Nested delivery schedule to handle multi-date split delivery rules natively
        schedule_item_schema = types.Schema(
            type=types.Type.OBJECT,
            properties={
                "Split_Quantity": types.Schema(type=types.Type.STRING, description="The specific partial quantity for this split date (e.g. '25 st')"),
                "Required_Date": types.Schema(type=types.Type.STRING, description="The delivery date for this specific batch (e.g. '7-8-2026')")
            },
            required=["Split_Quantity", "Required_Date"]
        )

        po_item_schema = types.Schema(
            type=types.Type.OBJECT,
            properties={
                "Line": types.Schema(type=types.Type.STRING, description="Sequence position or row index number"),
                "Item": types.Schema(type=types.Type.STRING, description="The primary item number listed under standard column (e.g. 514214)"),
                "Customer_Item": types.Schema(type=types.Type.STRING, description="The buyer's part number from notes labeled as 'Uw teknr', 'Your part no' (e.g. 4022.622.9146.3R103)"),
                "Description": types.Schema(type=types.Type.STRING, description="Product item text name description"),
                "Unit_Price": types.Schema(type=types.Type.STRING, description="Price text including conditions (e.g., '€ 1.423,61 per 1 st')"),
                "Deliveries": types.Schema(
                    type=types.Type.ARRAY, 
                    items=schedule_item_schema, 
                    description="List of all scheduled delivery dates and split quantities for this single line item."
                )
            },
            required=["Item", "Unit_Price", "Deliveries"]
        )

        final_response_schema = types.Schema(
            type=types.Type.ARRAY,
            items=po_item_schema
        )

        # PROMPT RE-TUNED FOR SPLIT DELIVERIES CAPTURE
        prompt = """
        You are an elite Purchase Order parsing specialist handling complex multi-lingual layouts with split delivery dates.
        
        CRITICAL EXTRACTION RULES:
        1. SPLIT DELIVERIES / VERZENDSCHEMA (HIGHEST PRIORITY):
           - A single line item might have multiple delivery dates and split quantities underneath 'Verzendschema:' or 'Delivery Schedule'.
           - For example, if it states '25 st' on '7-8-2026' AND '25 st' on '21-8-2026', you MUST capture BOTH separate entries inside the `Deliveries` array.
           - Do not skip subsequent dates or squash them into a single string.
        2. DUAL ITEM NUMBER EXTRACTION:
           - Scan standard columns for item numbers (e.g., '514214'). Also look at the descriptions/notes block right around it for the Customer Part Number labeled as 'Uw teknr:' (e.g., '4022 622 9146.3').
           - Store standard in `Item` and buyer's code in `Customer_Item`.
        3. MULTI-LINGUAL HEADERS:
           - Look for 'Verzendschema', 'Leverdatum', 'Shipping date' -> Delivery dates.
           - Look for 'Aantal', 'Menge', 'Quantity' -> Contains quantity and text prices (e.g., '50 st à € 1.423,61 per 1 st'). Extract the Unit Price accurately.
        """

        all_po_items = []
        progress_bar = st.progress(0)
        quota_exhausted = False
        
        for idx, pdf_file in enumerate(pdf_files):
            st.write(f"🔍 AI is processing file: **{pdf_file.name}**")
            try:
                pdf_file.seek(0)
                pdf_bytes = pdf_file.read()
                if len(pdf_bytes) == 0: continue

                pdf_reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
                total_pages = len(pdf_reader.pages)
                pages_per_chunk = 2 
                
                for start_page in range(0, total_pages, pages_per_chunk):
                    end_page = min(start_page + pages_per_chunk, total_pages)
                    st.write(f"   📄 Parsing pages {start_page + 1} to {end_page} (Total: {total_pages} pages)...")
                    
                    pdf_writer = pypdf.PdfWriter()
                    for p_idx in range(start_page, end_page):
                        pdf_writer.add_page(pdf_reader.pages[p_idx])
                    
                    chunk_buffer = io.BytesIO()
                    pdf_writer.write(chunk_buffer)
                    chunk_bytes = chunk_buffer.getvalue()

                    max_retries = 4
                    retry_delay = 4  
                    response = None
                    
                    for attempt in range(max_retries):
                        try:
                            response = client.models.generate_content(
                                model='gemini-2.5-flash',
                                contents=[
                                    types.Part.from_bytes(data=chunk_bytes, mime_type='application/pdf'),
                                    prompt
                                ],
                                config=types.GenerateContentConfig(
                                    response_mime_type="application/json",
                                    response_schema=final_response_schema,
                                    temperature=0.0
                                )
                            )
                            break  
                        except Exception as api_err:
                            err_msg = str(api_err)
                            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                                st.error("🛑 **Daily Quota Limit Reached (429)!**")
                                quota_exhausted = True
                                break
                            if "503" in err_msg or "UNAVAILABLE" in err_msg:
                                if attempt < max_retries - 1:
                                    time.sleep(retry_delay)
                                    retry_delay *= 2  
                                    continue
                            raise api_err  
                    
                    if quota_exhausted: st.stop()
                        
                    if response and response.text:
                        try:
                            items_data = json.loads(response.text.strip())
                            for item in items_data:
                                item['PO_Source_File'] = pdf_file.name
                                all_po_items.append(item)
                        except json.JSONDecodeError:
                            st.warning(f"⚠️ Failed to parse JSON for pages {start_page+1}-{end_page}, skipping chunk.")
                    time.sleep(2)
                        
            except Exception as e:
                st.error(f"❌ Failed to parse {pdf_file.name}: {e}")
            
            progress_bar.progress((idx + 1) / len(pdf_files))

        if not all_po_items:
            st.error("❌ No data was extracted from your PDF items. Processing stopped.")
            st.stop()

        df_po = pd.DataFrame(all_po_items)
        st.write("🔄 Aligning and matching split-delivery rows into structural layout...")

        structured_rows = []
        
        # We track how rows are generated to highlight discrepancies accurately in openpyxl later
        excel_row_blocks_meta = [] 

        for idx, row in df_excel.iterrows():
            excel_item = str(row.get(item_col_name, '')).strip()
            excel_key = clean_key(excel_item)
            ex_line = clean_line_num(row.get('Line'))
            
            # 1. Standardize Glovia Master Row values
            excel_side_row = {col: row[col] for col in df_excel.columns}
            excel_side_row['Data Block Source'] = 'GloviaG2'
            
            if 'Order Quantity' in excel_side_row:
                ex_qty_num, _ = parse_qty_and_unit(excel_side_row['Order Quantity'])
                if ex_qty_num is not None:
                    excel_side_row['Order Quantity'] = str(ex_qty_num)
            
            if 'Unit Price' in excel_side_row:
                ex_price_num = normalize_price(excel_side_row['Unit Price'])
                if ex_price_num is not None:
                    excel_side_row['Unit Price'] = str(ex_price_num)
            
            if 'Required Date/Time' in excel_side_row:
                excel_side_row['Required Date/Time'] = parse_date_to_custom_format(excel_side_row['Required Date/Time'])
            
            # 2. Look for matching candidates in the extracted PDF items pool
            matched_po_item = None
            if not df_po.empty and excel_key:
                candidates = []
                for po_idx, po_row in df_po.iterrows():
                    po_key_p = clean_key(po_row.get('Item', ''))
                    po_key_c = clean_key(po_row.get('Customer_Item', ''))
                    po_key_d = clean_key(po_row.get('Description', ''))
                    
                    if (po_key_p and (po_key_p in excel_key or excel_key in po_key_p)) or \
                       (po_key_c and (po_key_c in excel_key or excel_key in po_key_c)) or \
                       (excel_key and excel_key in po_key_d):
                        candidates.append(po_row)
                
                if candidates:
                    matched_po_item = candidates[0]
                    for cand in candidates:
                        if clean_line_num(cand.get('Line')) == ex_line:
                            matched_po_item = cand
                            break

            # 3. Handle data rendering depending on if item is found or contains multiple deliveries
            block_start_index = len(structured_rows)
            
            if matched_po_item is not None:
                deliveries = matched_po_item.get('Deliveries', [])
                if not isinstance(deliveries, list) or len(deliveries) == 0:
                    deliveries = [{"Split_Quantity": "0", "Required_Date": ""}]
                
                # We calculate aggregate quantity across all splits to check if sum total matches system
                total_pdf_item_qty = 0
                for d in deliveries:
                    q_num, _ = parse_qty_and_unit(d.get('Split_Quantity', '0'))
                    if q_num: total_pdf_item_qty += q_num

                # Append Glovia row first
                excel_side_row['Confirmation Note'] = "" # Will fill if single match or leave for splits
                structured_rows.append(excel_side_row)

                # Append a distinct line for EACH split delivery date found under the same item code!
                for d_idx, deliv in enumerate(deliveries):
                    pdf_side_row = {col: None for col in df_excel.columns}
                    pdf_side_row['Data Block Source'] = 'PDF'
                    pdf_side_row[item_col_name] = excel_item
                    if 'Line' in df_excel.columns: pdf_side_row['Line'] = matched_po_item.get('Line')
                    
                    # Map static fields
                    raw_price = matched_po_item.get('Unit_Price')
                    calc_price = normalize_price(raw_price)
                    if 'Unit Price' in df_excel.columns:
                        pdf_side_row['Unit Price'] = str(calc_price) if calc_price is not None else raw_price
                        
                    pdf_side_row[unnamed_col] = matched_po_item.get('Description', 'No Description')
                    if 'Notes' in df_excel.columns: 
                        pdf_side_row['Notes'] = f"Extracted from PDF: {matched_po_item.get('PO_Source_File', '')} [Batch {d_idx+1}]"
                    
                    # Map dynamic schedule fields
                    raw_split_qty = deliv.get('Split_Quantity')
                    clean_split_qty, extracted_unit = parse_qty_and_unit(raw_split_qty)
                    if 'Order Quantity' in df_excel.columns:
                        pdf_side_row['Order Quantity'] = str(clean_split_qty) if clean_split_qty is not None else raw_split_qty
                        
                    for um_col in ['UM', 'Stock UM', 'In Stock UM']:
                        if um_col in df_excel.columns:
                            pdf_side_row[um_col] = extracted_unit if extracted_unit else excel_side_row.get(um_col)

                    formatted_deliv_date = parse_date_to_custom_format(deliv.get('Required_Date'))
                    if 'Required Date/Time' in df_excel.columns:
                        pdf_side_row['Required Date/Time'] = formatted_deliv_date

                    # Build targeted Confirmation Notes text format: DDMY
                    conf_text = ""
                    if formatted_deliv_date:
                        today_str = datetime.now().strftime("%d%m")
                        ddmmyy = ""
                        if '/' in formatted_deliv_date:
                            parts = formatted_deliv_date.split('/')
                            if len(parts) >= 3: ddmmyy = parts[0] + parts[1] + parts[2][-2:]
                        conf_text = f"{today_str} lla dd conf. {ddmmyy}"
                    
                    pdf_side_row['Confirmation Note'] = conf_text
                    structured_rows.append(pdf_side_row)
                
                block_end_index = len(structured_rows)
                excel_row_blocks_meta.append({
                    "type": "MATCHED",
                    "start": block_start_index,
                    "end": block_end_index,
                    "total_pdf_qty": total_pdf_item_qty
                })
            else:
                # Completely missing from PDF
                if 'Notes' in df_excel.columns: excel_side_row['Notes'] = ""
                excel_side_row['Confirmation Note'] = ""
                structured_rows.append(excel_side_row)
                
                pdf_side_row = {col: None for col in df_excel.columns}
                pdf_side_row['Data Block Source'] = 'PDF'
                pdf_side_row[item_col_name] = excel_item
                if 'Notes' in df_excel.columns: pdf_side_row['Notes'] = "Not found in PO PDF"
                for um_col in ['UM', 'Stock UM', 'In Stock UM']:
                    if um_col in df_excel.columns: pdf_side_row[um_col] = excel_side_row.get(um_col)
                structured_rows.append(pdf_side_row)
                
                block_end_index = len(structured_rows)
                excel_row_blocks_meta.append({
                    "type": "MISSING",
                    "start": block_start_index,
                    "end": block_end_index,
                    "total_pdf_qty": 0
                })

            # Append trailing spacing row
            blank_row = {col: None for col in df_excel.columns}
            blank_row['Data Block Source'] = None
            blank_row['Confirmation Note'] = None
            structured_rows.append(blank_row)

        df_final = pd.DataFrame(structured_rows)
        core_cols = [c for c in df_final.columns if c not in ['Data Block Source', 'Confirmation Note']]
        cols = ['Data Block Source'] + core_cols + ['Confirmation Note']
        df_final = df_final[cols]
        df_final.to_excel(OUTPUT_FILENAME, index=False)

        # Style and Highlight via OpenPyXL using dynamic cell blocks metadata
        wb = openpyxl.load_workbook(OUTPUT_FILENAME)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]

        idx_qty = headers.index('Order Quantity') + 1 if 'Order Quantity' in headers else None
        idx_price = headers.index('Unit Price') + 1 if 'Unit Price' in headers else None
        idx_date = headers.index('Required Date/Time') + 1 if 'Required Date/Time' in headers else None

        fill_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        fill_light_gray = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        fill_yellow = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
        
        for block in excel_row_blocks_meta:
            start_r = block["start"] + 2 # offset spreadsheet header
            end_r = block["end"] + 1
            
            # Glovia Row is always the first index of the structured block
            glovia_row_num = start_r
            
            if block["type"] == "MISSING":
                for r in range(start_r, end_r):
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=r, column=c).fill = fill_yellow
                continue
            
            # Style standard matching rows block gray
            for r in range(start_r, end_r):
                for c in range(1, len(headers) + 1):
                    ws.cell(row=r, column=c).fill = fill_light_gray
            
            # Validate pricing and quantities safely against split groups
            ex_qty_val, _ = parse_qty_and_unit(ws.cell(row=glovia_row_num, column=idx_qty).value if idx_qty else 0)
            ex_price_val = normalize_price(ws.cell(row=glovia_row_num, column=idx_price).value if idx_price else 0)
            ex_date_val = str(ws.cell(row=glovia_row_num, column=idx_date).value).strip() if idx_date else ""

            # Check if total sum quantity matches system target requirement
            qty_discrepancy = (ex_qty_val != block["total_pdf_qty"])

            # Evaluate each individual PDF row sub-entry
            for pdf_row_num in range(glovia_row_num + 1, end_r):
                if idx_qty and qty_discrepancy:
                    ws.cell(row=glovia_row_num, column=idx_qty).fill = fill_red
                    ws.cell(row=pdf_row_num, column=idx_qty).fill = fill_red
                    
                if idx_price:
                    p_cell = ws.cell(row=pdf_row_num, column=idx_price)
                    if ex_price_val != normalize_price(p_cell.value) and p_cell.value is not None:
                        ws.cell(row=glovia_row_num, column=idx_price).fill = fill_red
                        p_cell.fill = fill_red
                        
                if idx_date:
                    d_cell = ws.cell(row=pdf_row_num, column=idx_date)
                    # For split delivery, individual date variances are natural but highlighted if distinct from original master row target
                    if ex_date_val != str(d_cell.value).strip() and d_cell.value is not None and d_cell.value != "":
                        # We flag date cell soft yellow or red if it shifts from target
                        d_cell.fill = fill_red

        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        st.success("🎉 Process Complete with Split Schedules Unpacked!")
        st.download_button(
            label="📥 Download Discrepancy Report",
            data=excel_buffer,
            file_name=OUTPUT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("💡 Please upload both the Master Excel file and PO PDFs to begin.")
