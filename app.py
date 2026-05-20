import os
import json
import io
import time  
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
st.write("Upload your System Master Data and PO PDFs to automatically generate a flagged discrepancy report.")

# Securely fetch API key
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY") or st.secrets.get("GOOGLE_API_KEY", "")

if not GOOGLE_API_KEY:
    st.error("🔑 Google API Key not found! Please configure it in your environment or secrets.")
    st.stop()

client = genai.Client(api_key=GOOGLE_API_KEY)
OUTPUT_FILENAME = "PO_Checking_Report.xlsx"

# Helper function to remove spaces/hyphens/dots for matching keys
def clean_key(val):
    if pd.isna(val) or val is None: return ""
    return str(val).replace(" ", "").replace("-", "").replace(".", "").lower().strip()

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
                "Line": types.Schema(type=types.Type.STRING, description="The sequence position, e.g. '1', '2', '3'"),
                "Item": types.Schema(type=types.Type.STRING, description="The drawing number or item string like '00497-00552-B' or '1117622.B'"),
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
        You are a meticulous purchase order parsing specialist. 
        This PDF document stacks data values vertically across text blocks. 
        Carefully associate each 'Item' (Drawing code/Part number) with its correct corresponding sequence order properties.
        
        CRITICAL RULES:
        1. Look closely at blocks where multiple line numbers, quantities, or prices are listed together (e.g. 10, 1, 6). Unpack them step-by-step so that every distinct Item code gets its own unique object block.
        2. Clean and capture the exact base 'Order Quantity' and 'Unit Price' values.
        3. Extract the 'Delivery Date' associated with that item block.
        """

        all_po_items = []
        progress_bar = st.progress(0)
        
        for idx, pdf_file in enumerate(pdf_files):
            st.write(f"🔍 AI is processing: **{pdf_file.name}**")
            try:
                pdf_file.seek(0)
                pdf_bytes = pdf_file.read()
                
                if len(pdf_bytes) == 0:
                    continue

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
                
                items_data = json.loads(response.text.strip())
                for item in items_data:
                    item['PO_Source_File'] = pdf_file.name
                    all_po_items.append(item)
                    
            except Exception as e:
                st.error(f"❌ Failed to parse {pdf_file.name}: {e}")
            
            progress_bar.progress((idx + 1) / len(pdf_files))
            if idx < len(pdf_files) - 1:
                time.sleep(5)

        if not all_po_items:
            st.error("❌ No data was extracted from your PDF items. Processing stopped.")
            st.stop()

        df_po = pd.DataFrame(all_po_items)
        st.write("🔄 Alternating and structuring report layout rows...")

        # Build precise custom alternated mapping matrix block:
        # [Excel Row] -> [PDF Row] -> [Blank Spacer Row] -> Repeat
        structured_rows = []
        
        for idx, row in df_excel.iterrows():
            excel_item = str(row.get(item_col_name, '')).strip()
            excel_key = clean_key(excel_item)
            
            # Formulate the Excel baseline dictionary row
            excel_side_row = {col: row[col] for col in df_excel.columns}
            excel_side_row['Data Block Source'] = 'Excel A (System Master Data)'
            if 'Required Date/Time' in excel_side_row:
                excel_side_row['Required Date/Time'] = parse_date_to_custom_format(excel_side_row['Required Date/Time'])
            
            # Search for the extracted PDF item match
            match = None
            if not df_po.empty and excel_key:
                for _, po_row in df_po.iterrows():
                    po_item = str(po_row.get('Item', '')).strip()
                    po_key = clean_key(po_item)
                    if po_key and (po_key in excel_key or excel_key in po_key):
                        match = po_row
                        break
            
            # Formulate the matched PDF dictionary row
            pdf_side_row = {col: None for col in df_excel.columns}
            pdf_side_row['Data Block Source'] = 'PDF Extracted (PO Check Zone)'
            pdf_side_row[item_col_name] = excel_item
            
            if match is not None:
                if 'Line' in df_excel.columns: pdf_side_row['Line'] = match.get('Line')
                if 'Unit Price' in df_excel.columns: pdf_side_row['Unit Price'] = match.get('Unit_Price')
                if 'Required Date/Time' in df_excel.columns: pdf_side_row['Required Date/Time'] = parse_date_to_custom_format(match.get('Required_Date'))
                if 'Order Quantity' in df_excel.columns: pdf_side_row['Order Quantity'] = match.get('Order_Quantity')
                pdf_side_row[unnamed_col] = match.get('Description', 'No Description')
                if 'Notes' in df_excel.columns: pdf_side_row['Notes'] = f"Extracted from PDF: {match.get('PO_Source_File', '')}"
            else:
                if 'Notes' in df_excel.columns: pdf_side_row['Notes'] = "Not found in PO PDF"
            
            # Empty spacer dict row
            blank_spacer_row = {col: None for col in df_excel.columns}
            blank_spacer_row['Data Block Source'] = None
            
            # Append rows strictly adhering to requested repeating layout pattern
            structured_rows.append(excel_side_row)
            structured_rows.append(pdf_side_row)
            structured_rows.append(blank_spacer_row)

        df_final = pd.DataFrame(structured_rows)
        
        # Shift Data Block Source identification tags column to the far front
        cols = ['Data Block Source'] + [c for c in df_final.columns if c != 'Data Block Source']
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

        # Style Fills
        fill_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        fill_light_gray = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        fill_yellow = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid') # Soft warning yellow
        
        # Loop over items step skipping by 3 to evaluate adjacent pairs 
        for i in range(len(df_excel)):
            row_excel_idx = (i * 3) + 2
            row_pdf_idx = (i * 3) + 3
            
            # Check if the row was missing in the PDF
            is_missing_in_pdf = False
            if idx_notes:
                notes_val = ws.cell(row=row_pdf_idx, column=idx_notes).value
                if notes_val == "Not found in PO PDF":
                    is_missing_in_pdf = True

            # Step 1: Base row color assignment (Yellow if missing, Gray if found)
            current_base_fill = fill_yellow if is_missing_in_pdf else fill_light_gray
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=row_excel_idx, column=col_idx).fill = current_base_fill

            # Step 2: Skip mismatch value checks if it wasn't even found in the PDF
            if is_missing_in_pdf:
                continue

            # Step 3: Run Value Discrepancy Checks (Mismatches override base colors with red)
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
