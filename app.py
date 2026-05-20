import os
import json
import io
import time  
from datetime import datetime
import pandas as pd
from google import genai  
from google.genai import types  
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st

# ==========================================
# APP CONFIGURATION & UI SETUP
# ==========================================
st.set_page_config(page_title="PO Checker AI", layout="centered")
st.title("📦 Purchase Order Checking Assistant")
st.write("Upload your System Master Data and PO PDFs to automatically generate a flagged discrepancy report with matching confirmation notes.")

# Securely fetch API key
GOOGLE_API_KEY = "AIzaSyD4888IrAzSh0utXCp4YQiJBGaOQR3-QiU"

if not GOOGLE_API_KEY:
    st.error("🔑 Google API Key not found! Please configure it in your environment or secrets.")
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

def normalize_numeric(v):
    if pd.isna(v) or v is None: return None
    try:
        s = str(v).replace(' ', '').replace('€', '').replace('st.', '').strip()
        if not s: return None
        if ',' in s and '.' in s: s = s.replace(',', '')
        elif ',' in s and '.' not in s: s = s.replace(',', '.')
        return round(float(s), 2)
    except:
        return None

# Custom date parser to output strictly DD/MM/YYYY
def parse_date_to_custom_format(v):
    if pd.isna(v) or v is None: return ""
    s = str(v).replace(':', '').replace('Delivery Date', '').strip()
    s = s.split(',')[0].strip().split()[0]
    for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d'):
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
                "Order_Quantity": types.Schema(type=types.Type.STRING, description="Quantity numerical value count"),
                "Unit_Price": types.Schema(type=types.Type.STRING, description="Price per piece numeric value"),
                "Required_Date": types.Schema(type=types.Type.STRING, description="Delivery Date string value")
            },
            required=["Item", "Order_Quantity", "Unit_Price"]
        )

        final_response_schema = types.Schema(
            type=types.Type.ARRAY,
            items=po_item_schema
        )

        prompt = """
        You are a meticulous purchase order parsing specialist working with multi-line aggregated layouts.
        This PDF documents text-tracks by grouping multiple records vertically in single rows.
        
        CRITICAL EXTRACTION RULES:
        1. Look closely at grouped lines (e.g., lines 3, 4, 5 or lines 8, 10). The item codes, line numbers, quantities, and prices are listed sequentially separated by newlines in single column blocks. 
           YOU MUST UNPACK THEM completely! Split them up so that every individual item code gets its own separate JSON object in the output list.
        2. Clean and capture the specific 'Order Quantity', 'Unit Price', and 'Delivery Date' fields associated with that item sequence rank position.
        3. Extract the item part numbers wherever they sit (checking both the column headers and descriptions text blocks).
        """

        all_po_items = []
        progress_bar = st.progress(0)
        quota_exhausted = False
        
        for idx, pdf_file in enumerate(pdf_files):
            st.write(f"🔍 AI is processing: **{pdf_file.name}**")
            try:
                pdf_file.seek(0)
                pdf_bytes = pdf_file.read()
                
                if len(pdf_bytes) == 0:
                    continue

                # --- PROTECTED API CALL WITH EMBEDDED RETRY LOGIC ---
                max_retries = 4
                retry_delay = 4  
                response = None
                
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[
                                types.Part.from_bytes(data=pdf_bytes, mime_type='application/pdf'),
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
                        
                        # Handle Hard Daily Limit Cap Exceeded
                        if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                            st.error("🛑 **Daily Quota Limit Reached (429)!** Your Gemini Free Tier account allows a maximum of 20 requests per day. To remove this limit, switch your project to Pay-As-You-Go billing in Google AI Studio, or use an API Key from a different Google account.")
                            quota_exhausted = True
                            break
                        
                        # Handle Server Busy
                        if "503" in err_msg or "UNAVAILABLE" in err_msg:
                            if attempt < max_retries - 1:
                                st.warning(f"⏳ Server overloaded (503). Retrying file {pdf_file.name} in {retry_delay}s... (Attempt {attempt + 1}/{max_retries})")
                                time.sleep(retry_delay)
                                retry_delay *= 2  
                                continue
                        raise api_err  
                
                if quota_exhausted:
                    st.stop()
                    
                items_data = json.loads(response.text.strip())
                for item in items_data:
                    item['PO_Source_File'] = pdf_file.name
                    all_po_items.append(item)
                    
            except Exception as e:
                st.error(f"❌ Failed to parse {pdf_file.name}: {e}")
            
            progress_bar.progress((idx + 1) / len(pdf_files))
            if idx < len(pdf_files) - 1:
                time.sleep(4)

        if not all_po_items:
            st.error("❌ No data was extracted from your PDF items. Processing stopped.")
            st.stop()

        df_po = pd.DataFrame(all_po_items)
        st.write("🔄 Alternating and structuring report layout rows...")

        structured_rows = []
        used_po_indices = set() 
        
        for idx, row in df_excel.iterrows():
            excel_item = str(row.get(item_col_name, '')).strip()
            excel_key = clean_key(excel_item)
            ex_line = clean_line_num(row.get('Line'))
            
            excel_side_row = {col: row[col] for col in df_excel.columns}
            excel_side_row['Data Block Source'] = 'Excel A (System Master Data)'
            
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
                            if clean_line_num(po_row.get('Line')) == ex_line:
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
            pdf_side_row['Data Block Source'] = 'PDF Extracted (PO Check Zone)'
            pdf_side_row[item_col_name] = excel_item
            
            confirmation_note_text = ""
            
            if match is not None:
                if 'Line' in df_excel.columns: pdf_side_row['Line'] = match.get('Line')
                if 'Unit Price' in df_excel.columns: pdf_side_row['Unit Price'] = match.get('Unit_Price')
                if 'Required Date/Time' in df_excel.columns: pdf_side_row['Required Date/Time'] = parse_date_to_custom_format(match.get('Required_Date'))
                if 'Order Quantity' in df_excel.columns: pdf_side_row['Order Quantity'] = match.get('Order_Quantity')
                pdf_side_row[unnamed_col] = match.get('Description', 'No Description')
                if 'Notes' in df_excel.columns: pdf_side_row['Notes'] = f"Extracted from PDF: {match.get('PO_Source_File', '')}"
                
                pdf_line = clean_line_num(match.get('Line'))
                ex_qty = normalize_numeric(row.get('Order Quantity'))
                pdf_qty = normalize_numeric(match.get('Order_Quantity'))
                ex_price = normalize_numeric(row.get('Unit Price'))
                pdf_price = normalize_numeric(match.get('Unit_Price'))
                ex_date = parse_date_to_custom_format(row.get('Required Date/Time'))
                pdf_date = parse_date_to_custom_format(match.get('Required_Date'))
                
                if (ex_line == pdf_line and ex_qty == pdf_qty and ex_price == pdf_price and ex_date == pdf_date and ex_date != ""):
                    today_str = datetime.now().strftime("%d%m")
                    
                    if '/' in ex_date:
                        date_parts = ex_date.split('/')
                        delivery_ddmmyy = date_parts[0] + date_parts[1] + date_parts[2][-2:]
                    else:
                        digits = "".join([c for c in ex_date if c.isdigit()])
                        if len(digits) >= 8:
                            delivery_ddmmyy = digits[:4] + digits[-2:]
                        else:
                            delivery_ddmmyy = digits[:6]
                        
                    confirmation_note_text = f"{today_str} LL confirmed delivery date {delivery_ddmmyy}"
            else:
                if 'Notes' in df_excel.columns: pdf_side_row['Notes'] = "Not found in PO PDF"
            
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

        # Style and Highlight via OpenPyXL cell indexing
        wb = openpyxl.load_workbook(OUTPUT_FILENAME)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]

        idx_line = headers.index('Line') + 1 if 'Line' in headers else None
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

            if idx_line:
                cell_e, cell_p = ws.cell(row=row_excel_idx, column=idx_line), ws.cell(row=row_pdf_idx, column=idx_line)
                if clean_line_num(cell_e.value) != clean_line_num(cell_p.value) and cell_p.value is not None:
                    cell_e.fill = fill_red; cell_p.fill = fill_red

            if idx_qty:
                cell_e, cell_p = ws.cell(row=row_excel_idx, column=idx_qty), ws.cell(row=row_pdf_idx, column=idx_qty)
                if normalize_numeric(cell_e.value) != normalize_numeric(cell_p.value) and cell_p.value is not None:
                    cell_e.fill = fill_red; cell_p.fill = fill_red

            if idx_price:
                cell_e, cell_p = ws.cell(row=row_excel_idx, column=idx_price), ws.cell(row=row_pdf_idx, column=idx_price)
                if normalize_numeric(cell_e.value) != normalize_numeric(cell_p.value) and cell_p.value is not None:
                    cell_e.fill = fill_red; cell_p.fill = fill_red
                    
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
