import os
import json
import io
import time  # Fix: Imported time to support the 12-second pause delay
import pandas as pd
import google.generativeai as genai
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st

# ==========================================
# APP CONFIGURATION & UI SETUP
# ==========================================
st.set_page_config(page_title="PO Checker AI", layout="centered")
st.title("📦 Purchase Order Checking Assistant")
st.write("Upload your System Master Data and PO PDFs to automatically generate a flagged discrepancy report.")

# Securely fetch API key from environment or Streamlit secrets
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY") or st.secrets.get("GOOGLE_API_KEY", "")

if not GOOGLE_API_KEY:
    st.error("🔑 Google API Key not found! Please configure it in your environment or secrets.")
    st.stop()

genai.configure(api_key=GOOGLE_API_KEY)
OUTPUT_FILENAME = "PO_Checking_Report.xlsx"

# ==========================================
# UI FILE UPLOADERS
# ==========================================
excel_file = st.file_uploader("👉 Step 1: Upload System Master Data (Excel or CSV)", type=["xlsx", "xls", "csv"], key="master_data_excel_csv")
pdf_files = st.file_uploader("👉 Step 2: Upload PO PDF file(s)", type=["pdf"], accept_multiple_files=True, key="po_pdf_files_list")

# Helper function to remove spaces/hyphens
def clean_key(val):
    if pd.isna(val) or val is None: return ""
    return str(val).replace(" ", "").replace("-", "").lower().strip()

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

def parse_date(v):
    if pd.isna(v) or v is None: return None
    s = str(v).split(',')[0].strip().split()[0]
    for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d'):
        try: return pd.to_datetime(s, format=fmt).strftime('%Y-%m-%d')
        except: continue
    return str(v).strip()

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

        # Detect/create target description column
        unnamed_col = None
        cols_list = list(df_excel.columns)
        if 'Notes' in cols_list:
            notes_idx = cols_list.index('Notes')
            if notes_idx + 1 < len(cols_list):
                next_col = cols_list[notes_idx + 1]
                if 'Unnamed' in next_col or next_col.strip() == '':
                    unnamed_col = next_col

        if not unnamed_col:
            unnamed_col = next((c for c in df_excel.columns if 'Unnamed' in c), 'Unnamed: 5')
            if unnamed_col not in df_excel.columns:
                df_excel.insert(cols_list.index('Notes') + 1, unnamed_col, None)

        # ==========================================
        # AI PDF Parsing Block (Model updated to gemini-1.5-flash-latest)
        # ==========================================
        model = genai.GenerativeModel('gemini-1.5-flash')
        prompt = """
        You are a purchase order parsing expert. Read the PDF purchase order and extract all "Line Items".
        Output ONLY a valid standard JSON Array. Do NOT include markdown syntax (like ```json) or any extra text.

        Required JSON Structure:
        [
          {
            "Line": "0001",
            "Item": "Product item number or material code",
            "Description": "Product name / text description",
            "Order Quantity": "100", 
            "Unit Price": "50.5",
            "Required Date/Time": "2026-08-03"
          }
        ]
        """

        all_po_items = []
        
        # Display progress status dynamically in the UI
        progress_bar = st.progress(0)
        for idx, pdf_file in enumerate(pdf_files):
            st.write(f"🔍 AI is processing: **{pdf_file.name}**")
            try:
                # Reset stream buffer pointer before reading
                pdf_file.seek(0)
                pdf_bytes = pdf_file.read()
                
                if len(pdf_bytes) == 0:
                    st.error(f"⚠️ File data for {pdf_file.name} was read empty. Skipping...")
                    continue

                response = model.generate_content([
                    {'mime_type': 'application/pdf', 'data': pdf_bytes},
                    prompt
                ])
                raw_text = response.text.strip()
                
                start_idx = raw_text.find('[')
                end_idx = raw_text.rfind(']') + 1
                if start_idx != -1 and end_idx != 0:
                    raw_text = raw_text[start_idx:end_idx]
                    
                items_data = json.loads(raw_text)
                if isinstance(items_data, dict):
                    items_data = [items_data]
                    
                for item in items_data:
                    item['PO_Source_File'] = pdf_file.name
                    all_po_items.append(item)
                    
            except Exception as e:
                # Check specifically if it's a rate limit error to give an intelligent warning
                if "429" in str(e) or "quota" in str(e).lower():
                    st.error(f"⚠️ Hit Google's speed limit on {pdf_file.name}. Pausing to reset...")
                    time.sleep(10) # Force a longer sleep if we trip the error
                else:
                    st.error(f"❌ Failed to parse {pdf_file.name}: {e}")
            
            progress_bar.progress((idx + 1) / len(pdf_files))
            
            # CRITICAL RATE-LIMIT FIX: 
            # If there are more PDFs left to parse, wait 12 seconds before the next one.
            # (12 seconds * 5 files = 60 seconds, which perfectly respects the 5 RPM limit)
            if idx < len(pdf_files) - 1:
                with st.spinner("⏳ Waiting 12 seconds to respect Gemini Free Tier limits..."):
                    time.sleep(12)

        # Check if processing completely failed before continuing to Excel building
        if not all_po_items:
            st.error("❌ No data was successfully extracted from your PDF(s). Processing stopped.")
            st.stop()

        df_po = pd.DataFrame(all_po_items)
        st.write("🔄 Generating block-comparison report...")

        # Match Mapping Logic
        pdf_rows = []
        for idx, row in df_excel.iterrows():
            excel_item = str(row.get('Item', '')).strip()
            excel_key = clean_key(excel_item)
            
            match = None
            if not df_po.empty:
                for _, po_row in df_po.iterrows():
                    po_item = str(po_row.get('Item', '')).strip()
                    po_key = clean_key(po_item)
                    if po_key and (po_key in excel_key or excel_key in po_key):
                        match = po_row
                        break
                        
            new_row = {col: None for col in df_excel.columns}
            new_row['Item'] = excel_item
            
            if match is not None:
                if 'Line' in df_excel.columns: new_row['Line'] = match.get('Line')
                if 'Unit Price' in df_excel.columns: new_row['Unit Price'] = match.get('Unit Price')
                if 'Required Date/Time' in df_excel.columns: new_row['Required Date/Time'] = match.get('Required Date/Time')
                if 'Order Quantity' in df_excel.columns: new_row['Order Quantity'] = match.get('Order Quantity')
                new_row[unnamed_col] = match.get('Description', 'No Description')
                if 'Notes' in df_excel.columns: new_row['Notes'] = f"Extracted from PDF: {match.get('PO_Source_File', '')}"
            else:
                if 'Notes' in df_excel.columns: new_row['Notes'] = "Not found in PO PDF"
                    
            pdf_rows.append(new_row)

        df_pdf = pd.DataFrame(pdf_rows)
        df_excel.insert(0, 'Data Block Source', 'Excel A (System Master Data)')
        df_pdf.insert(0, 'Data Block Source', 'PDF Extracted (PO Check Zone)')
        spacer_row = pd.DataFrame([{col: None for col in df_excel.columns}])
        df_final = pd.concat([df_excel, spacer_row, df_pdf], ignore_index=True)
        
        # Save to local file environment
        df_final.to_excel(OUTPUT_FILENAME, index=False)

        # Highlight Mismatches using OpenPyXL
        wb = openpyxl.load_workbook(OUTPUT_FILENAME)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]

        idx_line = headers.index('Line') + 1 if 'Line' in headers else None
        idx_qty = headers.index('Order Quantity') + 1 if 'Order Quantity' in headers else None
        idx_price = headers.index('Unit Price') + 1 if 'Unit Price' in headers else None
        idx_date = headers.index('Required Date/Time') + 1 if 'Required Date/Time' in headers else None

        fill_red = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        n_records = len(df_excel)

        for i in range(n_records):
            row_excel_idx = i + 2
            row_pdf_idx = i + 3 + n_records
            
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
                if parse_date(cell_e.value) != parse_date(cell_p.value) and cell_p.value is not None:
                    cell_e.fill = fill_red; cell_p.fill = fill_red

        # Save highlighted workbook to a byte stream buffer for Streamlit downloading
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        st.success("🎉 Process Complete!")
        
        # Web Download button interface
        st.download_button(
            label="📥 Download Discrepancy Report",
            data=excel_buffer,
            file_name=OUTPUT_FILENAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("💡 Please upload both the Master Excel file and PO PDFs to begin.")
