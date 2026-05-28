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
st.set_page_config(page_title="PO Checker AI", layout="centered")
st.title("📦 Purchase Order Checking Assistant")
st.write("Upload your System Master Data and PO PDFs to automatically generate a flagged discrepancy report with matching confirmation notes.")

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

# NEW: Highly robust parsing tool to strip extra zeros (,0000) and isolate units (ST, pack, EA)
def parse_qty_and_unit(v):
    if pd.isna(v) or v is None:
        return None, None
    s = str(v).strip()
    
    # Match the leading numbers/decimals and capture trailing text units
    match = re.match(r'^([\d\s.,]+)(.*)$', s)
    if match:
        num_part = match.group(1).strip()
        unit_part = match.group(2).strip()
        
        # Handle specific system decimal formatting like "500,0000" or "500.0000"
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

# Smart Price Normalizer to handle currency formatting and bulk rates (e.g., "per 100 st 5,34" vs "0,0534")
def normalize_price(v):
    if pd.isna(v) or v is None: return None
    s = str(v).lower().strip()
    
    factor = 1.0
    if "per 100" in s or "/100" in s or "per100" in s:
        factor = 100.0
    elif "per 1000" in s or "/1000" in s or "per1000" in s:
        factor = 1000.0
        
    for text_to_remove in ['st.', 'st', '€', 'eur', 'per 1000', 'per 100', 'per1000', 'per100', 'ex works', 'piece', 'pcs']:
        s = s.replace(text_to_remove, '')
    s = s.strip()
    
    if not s: return None
    if s.startswith(','): s = '0' + s
        
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

# Smart date parser supporting both dot (.) and slash (/) formats (e.g., 09.07.2026 -> 09/07/2026)
def parse_date_to_custom_format(v):
    if pd.isna(v) or v is None: return ""
    s = str(v).replace(':', '').replace('Delivery Date', '').strip()
    s = s.replace('.', '/')  
    s = s.split(',')[0].strip().split()[0]
    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%m/%d/%Y', '%d-%m-%Y', '%Y-%m-%d'):
        try: 
            return pd.to_datetime(s, format=fmt).strftime('%d/%m/%Y')
        except: 
            continue
    return str(v).strip()

# ==========================================
# UI FILE UPLOADERS (ALL ENGLISH)
# ==========================================
excel_file = st.file_uploader("👉 Step 1: Upload System Master Data (Excel or CSV)", type=["xlsx", "xls", "csv"], key="master_data_excel_csv")
pdf_files = st.file_uploader("👉 Step 2: Upload PO PDF file(s)", type=["pdf"], accept_multiple_files=True, key="po_pdf_files_list")

# ==========================================
# APP PROCESSING LOGIC
# ==========================================
if excel_file and pdf_files:
    if st.button("🚀 Run AI Verification Report"):
        
        # Read Master Data
        if excel_file.name.lower().endswith('.csv'):
            df_excel = pd.read_csv(excel_file, dtype=str)
        else:
            df_excel = pd.read_excel(excel_file, dtype=str)
            
        df_excel.columns = df_excel.columns.astype(str).str.strip()

        # Discover the Primary Item Identification column in Master Data
        item_col_name = next((c for c in df_excel.columns if c.lower() in ['item', 'item number', 'material', 'part no', 'part number']), 'Item')
        if item_col_name not in df_excel.columns and len(df_excel.columns) > 0:
            item_col_name = df_excel.columns[0]

        unnamed_col = next((c for c in df_excel.columns if 'Unnamed' in c), 'Description_Extracted')
        if unnamed_col not in df_excel.columns:
            df_excel[unnamed_col] = None

        # Define structured extraction schema layout
        po_item_schema = types.Schema(
            type=types.Type.OBJECT,
            properties={
                "Line": types.Schema(type=types.Type.STRING, description="The sequence position or row index number, e.g. '1', '2', '3', '7', '8'"),
                "Item": types.Schema(type=types.Type.STRING, description="The drawing number or item string like '1237190 Rev: 01'"),
                "Description": types.Schema(type=types.Type.STRING, description="Product item text name description"),
                "Order_Quantity": types.Schema(type=types.Type.STRING, description="Quantity numerical value count optionally with unit like '500 ST' or '100 pack'"),
                "Unit_Price": types.Schema(type=types.Type.STRING, description="Price text including conditions, e.g., 'per 100 st 5,34' or '0,0534'"),
                "Required_Date": types.Schema(type=types.Type.STRING, description="Delivery Date string value")
            },
            required=["Item", "Order_Quantity", "Unit_Price"]
        )

        final_response_schema = types.Schema(
            type=types.Type.ARRAY,
            items=po_item_schema
        )

        # Advanced prompting for multi-page document layout parsing
        prompt = """
        You are a meticulous purchase order parsing specialist working with multi-line aggregated layouts.
        This PDF documents text-tracks by grouping multiple records vertically in single rows.
        
        CRITICAL EXTRACTION RULES:
        1. Look closely at grouped lines (e.g., lines 3, 4, 5 or lines 8, 10). The item codes, line numbers, quantities, and prices may be listed sequentially separated by newlines in single column blocks, OR temporarily merged due to layout distortion (e.g., QTY and Price squeezed together like '5 EA\\n 582,16'). 
           YOU MUST UNPACK THEM completely and correctly split them up so that every individual item code gets its own separate JSON object in the output list.
        2. Clean and capture the specific 'Order Quantity', 'Unit Price' (preserve full text mapping context like 'per 100 st 5,34'), and 'Delivery Date' fields associated with that item sequence rank position.
        3. Extract the item part numbers wherever they sit (checking both the column headers and descriptions text blocks).
        """

        all_po_items = []
        progress_bar = st.progress(0)
        quota_exhausted = False
        
        for idx, pdf_file in enumerate(pdf_files):
            st.write(f"🔍 AI is processing file: **{pdf_file.name}**")
            try:
                pdf_file.seek(0)
                pdf_bytes = pdf_file.read()
                
                if len(pdf_bytes) == 0:
                    continue

                # --- Using pypdf for long file page chunking (2 pages per batch) ---
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

                    # --- PROTECTED API CALL WITH EMBEDDED RETRY LOGIC ---
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
                    
                    if quota_exhausted:
                        st.stop()
                        
                    if response and response.text:
                        try:
                            items_data = json.loads(response.text.strip())
                            for item in items_data:
                                item['PO_Source_File'] = pdf_file.name
                                all_po_items.append(item)
                        except json.JSONDecodeError:
                            st.warning(f"⚠️ Failed to parse JSON for pages {start_page+1}-{end_page}, automatically skipping this chunk.")
                    
                    time.sleep(2)
                        
            except Exception as e:
                st.error(f"❌ Failed to parse {pdf_file.name}: {e}")
            
            progress_bar.progress((idx + 1) / len(pdf_files))

        if not all_po_items:
            st.error("❌ No data was extracted from your PDF items. Processing stopped.")
            st.stop()

        df_po = pd.DataFrame(all_po_items)
        st.write("🔄 Aligning and structuring report layout rows...")

        structured_rows = []
        used_po_indices = set() 
        
        for idx, row in df_excel.iterrows():
            excel_item = str(row.get(item_col_name, '')).strip()
            excel_key = clean_key(excel_item)
            ex_line = clean_line_num(row.get('Line'))
            
            excel_side_row = {col: row[col] for col in df_excel.columns}
            excel_side_row['Data Block Source'] = 'GloviaG2'
            
            # Formats/cleans the Excel Quantity row display immediately (removes ,0000)
            if 'Order Quantity' in excel_side_row:
                ex_qty_num, _ = parse_qty_and_unit(excel_side_row['Order Quantity'])
                if ex_qty_num is not None:
                    excel_side_row['Order Quantity'] = str(ex_qty_num)
            
            if 'Required Date/Time' in excel_side_row:
                excel_side_row['Required Date/Time'] = parse_date_to_custom_format(excel_side_row['Required Date/Time'])
            
            match = None
            if not df_po.empty and excel_key:
                candidates = []
                for po_idx, po_row in df_po.iterrows():
                    po_item = str(po_row.get('Item', '')).strip()
                    po_key = clean_key(po_item)
                    po_desc_key = clean_key(po_row.get('Description', ''))
                    
                    if (po_key and (po_key in excel_key or excel_key in po_key)) or (excel_key and excel_key in po_desc_key):
                        candidates.append((po_idx, po_row))
                
                if candidates:
                    for po_idx, po_row in candidates:
                        if clean_line_num(po_row.get('Line')) == ex_line and po_idx not in used_po_indices:
                            match = po_row
                            used_po_indices.add(po_idx)
                            break
                    if match is None:
                        for po_idx, po_row in candidates:
                            if po_idx not in used_po_indices:
                                match = po_row
                                used_po_indices.add(po_idx)
                                break
                    if match is None:
                        match = candidates[0][1]

            pdf_side_row = {col: None for col in df_excel.columns}
            pdf_side_row['Data Block Source'] = 'PDF'
            pdf_side_row[item_col_name] = excel_item
            
            confirmation_note_text = ""
            
            if match is not None:
                if 'Line' in df_excel.columns: pdf_side_row['Line'] = match.get('Line')
                if 'Unit Price' in df_excel.columns: pdf_side_row['Unit Price'] = match.get('Unit_Price')
                if 'Required Date/Time' in df_excel.columns: pdf_side_row['Required Date/Time'] = parse_date_to_custom_format(match.get('Required_Date'))
                
                # Dynamic Extraction: Splits numeric part from unit (e.g. "500 ST" -> "500" & "ST")
                raw_pdf_qty = match.get('Order_Quantity')
                cleaned_pdf_qty, extracted_unit = parse_qty_and_unit(raw_pdf_qty)
                
                if 'Order Quantity' in df_excel.columns: 
                    pdf_side_row['Order Quantity'] = str(cleaned_pdf_qty) if cleaned_pdf_qty is not None else raw_pdf_qty
                
                # Smart UM Alignment: Push extracted units automatically into available measure columns
                for um_col in ['UM', 'Stock UM', 'In Stock UM']:
                    if um_col in df_excel.columns:
                        pdf_side_row[um_col] = extracted_unit if extracted_unit else excel_side_row.get(um_col)
                        
                pdf_side_row[unnamed_col] = match.get('Description', 'No Description')
                if 'Notes' in df_excel.columns: pdf_side_row['Notes'] = f"Extracted from PDF: {match.get('PO_Source_File', '')}"
                
                pdf_date = parse_date_to_custom_format(match.get('Required_Date'))
                
                if pdf_date and str(pdf_date).strip() != "":
                    today_str = datetime.now().strftime("%d%m")
                    delivery_ddmmyy = ""
                    pdf_date_str = str(pdf_date).strip()
                    if '/' in pdf_date_str:
                        date_parts = pdf_date_str.split('/')
                        if len(date_parts) >= 3:
                            delivery_ddmmyy = date_parts[0] + date_parts[1] + date_parts[2][-2:]
                    confirmation_note_text = f"{today_str} lla dd conf. {delivery_ddmmyy}"
            else:
                if 'Notes' in df_excel.columns: pdf_side_row['Notes'] = "Not found in PO PDF"
                # Keep unit intact from Excel if PDF item was missing completely
                for um_col in ['UM', 'Stock UM', 'In Stock UM']:
                    if um_col in df_excel.columns:
                        pdf_side_row[um_col] = excel_side_row.get(um_col)
            
            excel_side_row['Confirmation Note'] = confirmation_note_text
            pdf_side_row['Confirmation Note'] = confirmation_note_text
            
            blank_spacer_row = {col: None for col in df_excel.columns}
            blank_spacer_row['Data Block Source'] = None
            blank_spacer_row['Confirmation Note'] = None
            
            structured_rows.append(excel_side_row)
            structured_rows.append(pdf_side_row)
            structured_rows.append(blank_spacer_row)

        df_final = pd.DataFrame(structured_rows)
        core_cols = [c for c in df_final.columns if c not in ['Data Block Source', 'Confirmation Note']]
        cols = ['Data Block Source'] + core_cols + ['Confirmation Note']
        df_final = df_final[cols]
        df_final.to_excel(OUTPUT_FILENAME, index=False)

        # Style and Highlight discrepancies via OpenPyXL
        wb = openpyxl.load_workbook(OUTPUT_FILENAME)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]

        idx_qty = headers.index('Order Quantity') + 1 if 'Order Quantity' in headers else None
        idx_price = headers.index('Unit Price') + 1 if 'Unit Price' in headers else None
        idx_date = headers.index('Required Date/Time') + 1 if 'Required Date/Time' in headers else None
        idx_notes = headers.index('Notes') + 1 if 'Notes' in headers else None

        fill_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        fill_light_gray = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        fill_yellow = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
        
        for i in range(len(df_excel)):
            row_excel_idx = (i * 3) + 2
            row_pdf_idx = (i * 3) + 3
            
            is_missing_in_pdf = False
            if idx_notes:
                notes_val = ws.cell(row=row_pdf_idx, column=idx_notes).value
                if notes_val == "Not found in PO PDF":
                    is_missing_in_pdf = True

            current_base_fill = fill_yellow if is_missing_in_pdf else fill_light_gray
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=row_excel_idx, column=col_idx).fill = current_base_fill

            if is_missing_in_pdf:
                continue

            # Quantities Check (Using advanced parser to ignore ,0000 and unit formatting differences)
            if idx_qty:
                cell_e, cell_p = ws.cell(row=row_excel_idx, column=idx_qty), ws.cell(row=row_pdf_idx, column=idx_qty)
                val_e, _ = parse_qty_and_unit(cell_e.value)
                val_p, _ = parse_qty_and_unit(cell_p.value)
                if val_e != val_p and cell_p.value is not None:
                    cell_e.fill = fill_red; cell_p.fill = fill_red

            # Unit Price Check 
            if idx_price:
                cell_e, cell_p = ws.cell(row=row_excel_idx, column=idx_price), ws.cell(row=row_pdf_idx, column=idx_price)
                if normalize_price(cell_e.value) != normalize_price(cell_p.value) and cell_p.value is not None:
                    cell_e.fill = fill_red; cell_p.fill = fill_red
                    
            # Dates Check
            if idx_date:
                cell_e, cell_p = ws.cell(row=row_excel_idx, column=idx_date), ws.cell(row=row_pdf_idx, column=idx_date)
                if cell_e.value != cell_p.value and cell_p.value is not None and cell_p.value != "":
                    cell_e.fill = fill_red; cell_p.fill = fill_red

        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        st.success("🎉 Process Complete!")
        st.download_button(
            label="📥 Download Discrepancy Report",
            data=excel_buffer,
            file_name=OUTPUT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("💡 Please upload both the Master Excel file and PO PDFs to begin.")
