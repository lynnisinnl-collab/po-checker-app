import io
import json
import re
import time
from datetime import datetime, date

import openpyxl
import pandas as pd
import pypdf  # Requirement: pip install pypdf
import streamlit as st
from google import genai
from google.genai import types
from openpyxl.styles import PatternFill

# ==========================================
# SETTINGS
# ==========================================
HIGHLIGHT_GLOVIA_SIDE = False   # False = only the PDF cell turns red; True = the matching GloviaG2 cell turns red too
PAGES_PER_CHUNK = 3             # Pages sent to Gemini per request
CHUNK_OVERLAP = 1               # Overlapping pages so a delivery schedule that crosses a page break is not cut off
PRICE_TOLERANCE = 0.001
OUTPUT_FILENAME = "PO_Checking_Report.xlsx"

FILL_RED = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
STYLE_FILLS = {
    "matched": PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid'),   # light gray
    "missing": PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid'),   # yellow: in Glovia, not in PDF
    "pdf_only": PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid'),  # orange: in PDF, not in Glovia
}


# ==========================================
# GENERIC HELPERS
# ==========================================
def is_blank(v):
    if v is None:
        return True
    if isinstance(v, (list, dict, tuple)):
        return len(v) == 0
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip().lower() in ("", "none", "nan", "null", "nat")


def fmt_num(v):
    if v is None:
        return "-"
    if isinstance(v, (int, float)):
        return ('%f' % v).rstrip('0').rstrip('.')
    return str(v)


# Matching cleanup helper to handle revision suffixes (e.g., rev01, rev03)
def clean_key(val):
    if is_blank(val):
        return ""
    s = str(val).replace(" ", "").replace("-", "").replace(".", "").lower().strip()
    if 'rev' in s:
        s = s.split('rev')[0].strip()
    return s


WORD_TO_DIGIT = {
    'zero': '0', 'one': '1', 'two': '2', 'three': '3', 'four': '4', 'five': '5',
    'six': '6', 'seven': '7', 'eight': '8', 'nine': '9', 'ten': '10',
    'twee': '2', 'drie': '3', 'vier': '4', 'vijf': '5', 'zes': '6',
    'zeven': '7', 'acht': '8', 'negen': '9', 'tien': '10',
}


# Semantic text cleaner: word-numbers -> digits, '+' / 'and' / 'en' -> '&'
def clean_description_semantic(val):
    if is_blank(val):
        return ""
    s = str(val).lower().strip()
    s = s.replace('+', ' & ')
    # Whole words only, so e.g. 'tekening' or 'stand' are not mangled
    s = re.sub(r'\b(and|en)\b', '&', s)
    for word, digit in WORD_TO_DIGIT.items():
        s = re.sub(r'\b' + word + r'\b', digit, s)
    return re.sub(r'[^a-z0-9&]', '', s)


# Standardize line numbers (e.g., '0009' -> '9', '090/000' -> '90', '10.0' -> '10')
def clean_line_num(val):
    if is_blank(val):
        return ""
    s = str(val).strip().split('-')[0].split('/')[0].strip()
    if re.fullmatch(r'\d+\.0+', s):
        s = s.split('.')[0]
    stripped = s.lstrip('0')
    return stripped if stripped else ("0" if s else "")


def lines_match(ex_line, pdf_line):
    """Exact match, or PDF line = Glovia line x 10 (e.g. PDF '090/000' vs Glovia '9')."""
    if not ex_line or not pdf_line:
        return False
    return ex_line == pdf_line or (pdf_line.endswith('0') and pdf_line[:-1] == ex_line)


def keys_match(a, b):
    if not a or not b:
        return False
    if a == b:
        return True
    # Substring matching only for keys long enough, otherwise '12' would match almost anything
    return min(len(a), len(b)) >= 4 and (a in b or b in a)


def desc_match_loose(a, b):
    if not a or not b:
        return False
    return a in b or b in a or a[:10] in b or b[:10] in a


# ==========================================
# NUMBER / PRICE / DATE PARSING
# ==========================================
def to_number(num_str):
    """Handles '3.457,44', '3,457.44', '25,0000', '1.000' (Dutch thousands), '12.5'."""
    if is_blank(num_str):
        return None
    s = str(num_str).replace(' ', '').replace('\u00a0', '').strip()
    if not s:
        return None
    if s[0] in ',.':
        s = '0' + s
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        if re.fullmatch(r'\d{1,3}(,\d{3}){2,}', s):
            s = s.replace(',', '')
        else:
            s = s.replace(',', '.')
    elif '.' in s:
        # '1.000' / '12.500.000' on a Dutch document = thousands separator
        if re.fullmatch(r'[1-9]\d{0,2}(\.\d{3})+', s):
            s = s.replace('.', '')
    try:
        val = float(s)
    except ValueError:
        return None
    return int(val) if val.is_integer() else val


# Split quantity and unit (e.g. '25,0000 ST' -> 25, 'ST')
def parse_qty_and_unit(v):
    if is_blank(v):
        return None, None
    s = str(v).strip()
    match = re.match(r'^([\d\s.,]+)(.*)$', s)
    if not match:
        return None, None
    val = to_number(match.group(1))
    if val is None:
        return None, None
    unit = match.group(2).strip()
    return val, (unit if unit else None)


# Price normalizer handling 'à', 'per X', '/ X', '€' etc.
def normalize_price(v):
    if is_blank(v):
        return None
    s = str(v).lower().replace('à', ' ').replace('\u00a0', ' ').strip()

    factor = 1.0
    m = re.search(r'(?:per|/)\s*(\d+(?:[.,]\d+)?)\s*[a-z.]*', s)
    if m:
        f = to_number(m.group(1))
        if f:
            factor = float(f)
        s = s[:m.start()] + ' ' + s[m.end():]

    num = re.search(r'\d[\d.,]*\d|\d|[.,]\d+', s)
    if not num:
        return None
    val = to_number(num.group(0))
    if val is None:
        return None
    return round(val / factor, 4)


DATE_FORMATS = ('%d/%m/%Y', '%d/%m/%y', '%Y/%m/%d', '%m/%d/%Y')


def parse_to_date_obj(v):
    if is_blank(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    s = re.sub(r'(?i)delivery date|required date|wk:|verzendschema|leverdatum', ' ', s)
    s = s.replace(':', ' ').replace('.', '/').replace('-', '/')
    s = s.split(',')[0].strip()
    if not s:
        return None
    s = s.split()[0]
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def format_date(v):
    d = parse_to_date_obj(v)
    if d:
        return d.strftime('%d/%m/%Y')
    return "" if is_blank(v) else str(v).strip()


def is_date_discrepancy(ex_val, pdf_val):
    # Rule 1: Glovia has no date -> nothing to compare
    if is_blank(ex_val):
        return False
    # Rule 2: Glovia has a date but PDF does not
    if is_blank(pdf_val):
        return True
    ex_str, pdf_str = str(ex_val).strip(), str(pdf_val).strip()
    if ex_str == pdf_str:
        return False
    ex_obj, pdf_obj = parse_to_date_obj(ex_val), parse_to_date_obj(pdf_val)
    if ex_obj and pdf_obj:
        return ex_obj != pdf_obj
    return True


def find_col(columns, candidates):
    lower = {str(c).lower(): c for c in columns}
    for cand in candidates:
        if cand.lower() in lower:
            return lower[cand.lower()]
    return None


# ==========================================
# PDF ITEM MERGING & DELIVERY PAIRING
# ==========================================
def merge_po_items(raw_items):
    """
    Gemini sometimes returns the same line item several times (once per delivery date,
    or once per overlapping page chunk). Merge them into ONE item that carries ALL deliveries.
    """
    merged, seen = {}, {}
    for it in raw_items:
        key = (it.get('PO_Source_File', ''), clean_line_num(it.get('Line')), clean_key(it.get('Item')))
        if key not in merged:
            base = {k: v for k, v in it.items() if k != 'Deliveries'}
            base['Deliveries'] = []
            merged[key] = base
            seen[key] = set()
        else:
            for f in ('Line', 'Customer_Item', 'Description', 'Order_Quantity', 'Unit_Price'):
                if is_blank(merged[key].get(f)) and not is_blank(it.get(f)):
                    merged[key][f] = it.get(f)

        for d in it.get('Deliveries') or []:
            if not isinstance(d, dict):
                continue
            q, _ = parse_qty_and_unit(d.get('Split_Quantity'))
            d_obj = parse_to_date_obj(d.get('Required_Date'))
            sig = (q, d_obj if d_obj else str(d.get('Required_Date') or '').strip().lower())
            if sig in seen[key]:
                continue
            seen[key].add(sig)
            merged[key]['Deliveries'].append(d)

    result = []
    for po in merged.values():
        parsed = [(d, parse_qty_and_unit(d.get('Split_Quantity'))[0], parse_to_date_obj(d.get('Required_Date')))
                  for d in po['Deliveries']]
        # Drop a quantity-less copy of a date that also exists with a quantity (overlap artefact)
        dates_with_qty = {dt for _, q, dt in parsed if q is not None and dt is not None}
        parsed = [p for p in parsed if not (p[1] is None and p[2] is not None and p[2] in dates_with_qty)]
        parsed.sort(key=lambda p: (p[2] is None, p[2] or date.min))
        po['Deliveries'] = [p[0] for p in parsed]

        po['_key'] = clean_key(po.get('Item'))
        po['_cust_key'] = clean_key(po.get('Customer_Item'))
        po['_line'] = clean_line_num(po.get('Line'))
        po['_desc'] = clean_description_semantic(po.get('Description'))
        result.append(po)
    return result


def parse_pdf_deliveries(po):
    deliveries = po.get('Deliveries') or []
    if not deliveries:
        deliveries = [{"Split_Quantity": po.get('Order_Quantity'), "Required_Date": ""}]
    out = []
    for d in deliveries:
        q, unit = parse_qty_and_unit(d.get('Split_Quantity'))
        out.append({
            "qty": q, "unit": unit, "raw_qty": d.get('Split_Quantity'),
            "date_raw": d.get('Required_Date'), "date_obj": parse_to_date_obj(d.get('Required_Date')),
        })
    if len(out) == 1 and out[0]["qty"] is None:
        q, unit = parse_qty_and_unit(po.get('Order_Quantity'))
        out[0]["qty"] = q
        out[0]["unit"] = out[0]["unit"] or unit
    return out


def pair_deliveries(ex_dates, pdf_dates):
    """Pair Glovia delivery lines with PDF deliveries: identical dates first, the rest in date order."""
    pairs, used_ex, used_pdf = [], set(), set()
    for j, p_d in enumerate(pdf_dates):
        if p_d is None:
            continue
        for i, e_d in enumerate(ex_dates):
            if i not in used_ex and e_d is not None and e_d == p_d:
                pairs.append((i, j))
                used_ex.add(i)
                used_pdf.add(j)
                break
    rem_ex = [i for i in range(len(ex_dates)) if i not in used_ex]
    rem_pdf = [j for j in range(len(pdf_dates)) if j not in used_pdf]
    pairs += list(zip(rem_ex, rem_pdf))
    pairs += [(i, None) for i in rem_ex[len(rem_pdf):]]
    pairs += [(None, j) for j in rem_pdf[len(rem_ex):]]
    big = 10 ** 9
    pairs.sort(key=lambda p: (p[0] if p[0] is not None else big, p[1] if p[1] is not None else big))
    return pairs


def find_po_match(ex_key, ex_line, ex_desc, po_items, used):
    best = None
    for i, po in enumerate(po_items):
        num = keys_match(ex_key, po['_key']) or keys_match(ex_key, po['_cust_key'])
        line_exact = bool(ex_line) and ex_line == po['_line']
        line_ok = lines_match(ex_line, po['_line'])
        desc = desc_match_loose(ex_desc, po['_desc'])
        if num or (line_ok and desc) or (line_ok and not ex_key):
            score = (2 if num else 0) + (2 if line_exact else (1 if line_ok else 0)) \
                    + (1 if desc else 0) + (1 if i not in used else 0)
            if best is None or score > best[0]:
                best = (score, i)
    return None if best is None else best[1]


# ==========================================
# REPORT BUILDER
# ==========================================
def build_report(df_excel, po_items):
    df_excel = df_excel.copy()
    df_excel.columns = df_excel.columns.astype(str).str.strip()
    df_excel = df_excel.dropna(how='all').reset_index(drop=True)
    cols = list(df_excel.columns)
    warnings = []

    item_col = find_col(cols, ['Item', 'Item Number', 'Material', 'Part No', 'Part Number']) or cols[0]
    line_col = find_col(cols, ['Line', 'Line No', 'Line Number'])
    qty_col = find_col(cols, ['Order Quantity', 'Quantity', 'Qty'])
    price_col = find_col(cols, ['Unit Price', 'Price'])
    date_col = find_col(cols, ['Required Date/Time', 'Required Date', 'Delivery Date', 'Due Date'])
    notes_col = find_col(cols, ['Notes'])
    desc_col = next((c for c in cols if 'Unnamed' in c), None) or find_col(cols, ['Description', 'Item Description'])
    um_cols = [c for c in cols if c in ('UM', 'Stock UM', 'In Stock UM')]

    for label, c in [('Order Quantity', qty_col), ('Unit Price', price_col), ('Required Date/Time', date_col)]:
        if c is None:
            warnings.append(f"Column '{label}' not found in master data - this comparison is skipped.")
    if desc_col is None:
        desc_col = 'Description_Extracted'
        df_excel[desc_col] = None
        cols = list(df_excel.columns)

    out_cols = ['Data Block Source'] + cols + ['Check Result', 'Confirmation Note']
    structured_rows, row_styles, red_cells = [], {}, set()
    today_str = datetime.now().strftime("%d%m")

    def add_row(r, style=None):
        structured_rows.append(r)
        idx = len(structured_rows) - 1
        if style:
            row_styles[idx] = style
        return idx

    def flag(pdf_out, col, glovia_out=None):
        if col is None:
            return
        red_cells.add((pdf_out, col))
        if HIGHLIGHT_GLOVIA_SIDE and glovia_out is not None:
            red_cells.add((glovia_out, col))

    def pdf_item_display(po):
        item = str(po.get('Item') or '').strip()
        cust = po.get('Customer_Item')
        return f"{item} / {str(cust).strip()}" if not is_blank(cust) else item

    def build_glovia_row(row):
        r = {c: (None if is_blank(row[c]) else row[c]) for c in cols}
        r['Data Block Source'] = 'GloviaG2'
        if qty_col:
            q, _ = parse_qty_and_unit(row[qty_col])
            if q is not None:
                r[qty_col] = q
        if price_col:
            p = normalize_price(row[price_col])
            if p is not None:
                r[price_col] = p
        if date_col:
            r[date_col] = format_date(row[date_col]) or None
        return r

    def build_pdf_row(po, d, j, n, ref_row=None):
        r = {c: None for c in cols}
        r['Data Block Source'] = 'PDF'
        r[item_col] = pdf_item_display(po)
        if line_col:
            r[line_col] = po.get('Line')
        if price_col:
            p = normalize_price(po.get('Unit_Price'))
            r[price_col] = p if p is not None else po.get('Unit_Price')
        r[desc_col] = po.get('Description') or None
        if notes_col:
            r[notes_col] = f"Extracted from PDF: {po.get('PO_Source_File', '')} [Delivery {j + 1}/{n}]"
        if qty_col:
            r[qty_col] = d['qty'] if d['qty'] is not None else d['raw_qty']
        for um in um_cols:
            r[um] = d['unit'] or (ref_row[um] if ref_row is not None and not is_blank(ref_row[um]) else None)
        if date_col:
            r[date_col] = format_date(d['date_raw']) or None
        r['Confirmation Note'] = (f"{today_str} lla dd conf. {d['date_obj'].strftime('%d%m%y')}"
                                  if d['date_obj'] else "")
        return r

    # ---- 1. Group Glovia rows (same item + same base line = multiple delivery lines) ----
    groups = {}
    for idx, row in df_excel.iterrows():
        gk = (clean_key(row[item_col]), clean_line_num(row[line_col]) if line_col else "")
        if gk == ("", ""):
            gk = ("__row__", idx)
        groups.setdefault(gk, []).append(idx)

    # ---- 2. Match each group to one PDF item; groups hitting the same PDF item are merged ----
    blocks, po_to_block, used = [], {}, set()
    for gk, row_idxs in groups.items():
        first = df_excel.loc[row_idxs[0]]
        po_idx = find_po_match(
            clean_key(first[item_col]),
            clean_line_num(first[line_col]) if line_col else "",
            clean_description_semantic(first[desc_col]),
            po_items, used,
        )
        if po_idx is not None:
            used.add(po_idx)
            if po_idx in po_to_block:
                blocks[po_to_block[po_idx]]['rows'].extend(row_idxs)
                continue
            po_to_block[po_idx] = len(blocks)
        blocks.append({'rows': list(row_idxs), 'po_idx': po_idx})

    # ---- 3. Build output rows ----
    for block in blocks:
        if block['po_idx'] is None:
            for i in block['rows']:
                add_row(build_glovia_row(df_excel.loc[i]), 'missing')
            first = df_excel.loc[block['rows'][0]]
            ph = {c: None for c in cols}
            ph['Data Block Source'] = 'PDF'
            ph[item_col] = first[item_col]
            if notes_col:
                ph[notes_col] = "Not found in PO PDF"
            for um in um_cols:
                ph[um] = first[um]
            ph['Check Result'] = "Not found in PO PDF"
            add_row(ph, 'missing')
            add_row({})
            continue

        po = po_items[block['po_idx']]
        pdf_delivs = parse_pdf_deliveries(po)

        ex_list = []
        for i in block['rows']:
            row = df_excel.loc[i]
            ex_list.append({
                "row": row,
                "qty": parse_qty_and_unit(row[qty_col])[0] if qty_col else None,
                "date_raw": row[date_col] if date_col else None,
                "date_obj": parse_to_date_obj(row[date_col]) if date_col else None,
            })
        ex_list.sort(key=lambda e: (e['date_obj'] is None, e['date_obj'] or date.min))
        base = ex_list[0]['row']

        # Item-level comparisons (same for every delivery row)
        ex_key = clean_key(base[item_col])
        item_diff = bool(ex_key) and not (keys_match(ex_key, po['_key']) or keys_match(ex_key, po['_cust_key']))
        ex_desc = clean_description_semantic(base[desc_col])
        desc_diff = bool(ex_desc and po['_desc']) and not (ex_desc in po['_desc'] or po['_desc'] in ex_desc)
        ex_price = normalize_price(base[price_col]) if price_col else None
        pdf_price = normalize_price(po.get('Unit_Price'))
        price_diff = ex_price is not None and (pdf_price is None or abs(ex_price - pdf_price) > PRICE_TOLERANCE)

        single_mode = len(ex_list) == 1
        if single_mode:
            # One Glovia line: every PDF delivery is shown; quantity is checked on the total
            pairs = [(0, 0)] + [(None, j) for j in range(1, len(pdf_delivs))]
            qtys = [d['qty'] for d in pdf_delivs if d['qty'] is not None]
            total_pdf = sum(qtys) if qtys else None
            ex_q = ex_list[0]['qty']
            total_qty_diff = ex_q is not None and (total_pdf is None or abs(total_pdf - ex_q) > 1e-9)
        else:
            pairs = pair_deliveries([e['date_obj'] for e in ex_list], [d['date_obj'] for d in pdf_delivs])

        glovia_out = {}
        for ex_i, pdf_j in pairs:
            if ex_i is not None:
                glovia_out[ex_i] = add_row(build_glovia_row(ex_list[ex_i]['row']), 'matched')
            first_glovia = glovia_out.get(0)

            if pdf_j is None:
                # Glovia delivery line with no counterpart in the PDF
                ph = {c: None for c in cols}
                ph['Data Block Source'] = 'PDF'
                ph[item_col] = pdf_item_display(po)
                if notes_col:
                    ph[notes_col] = f"Extracted from PDF: {po.get('PO_Source_File', '')} [no matching delivery]"
                ph['Check Result'] = "Glovia delivery line not found in PDF"
                out = add_row(ph, 'matched')
                flag(out, qty_col, glovia_out.get(ex_i))
                flag(out, date_col, glovia_out.get(ex_i))
                continue

            d = pdf_delivs[pdf_j]
            ref_i = ex_i if ex_i is not None else (0 if single_mode else None)
            ref = ex_list[ref_i] if ref_i is not None else None
            ref_out = glovia_out.get(ref_i) if ref_i is not None else None

            r = build_pdf_row(po, d, pdf_j, len(pdf_delivs), ref['row'] if ref else base)
            issues, flags = [], []

            if item_diff:
                issues.append(f"Item: {po.get('Item')} ≠ {base[item_col]}")
                flags.append((item_col, first_glovia))
            if desc_diff:
                issues.append("Description differs")
                flags.append((desc_col, first_glovia))
            if price_diff:
                issues.append(f"Price: {fmt_num(pdf_price) if pdf_price is not None else po.get('Unit_Price')} ≠ {fmt_num(ex_price)}")
                flags.append((price_col, first_glovia))

            if qty_col:
                if single_mode:
                    if total_qty_diff:
                        issues.append(f"Total qty: {fmt_num(total_pdf)} ≠ {fmt_num(ex_q)}")
                        flags.append((qty_col, ref_out))
                elif ref is None:
                    flags.append((qty_col, None))
                elif ref['qty'] is not None and (d['qty'] is None or abs(d['qty'] - ref['qty']) > 1e-9):
                    issues.append(f"Qty: {fmt_num(d['qty'])} ≠ {fmt_num(ref['qty'])}")
                    flags.append((qty_col, ref_out))

            if date_col:
                if ref is None:
                    issues.append("Extra delivery - no matching Glovia delivery line")
                    flags.append((date_col, None))
                elif is_date_discrepancy(ref['date_raw'], d['date_raw']):
                    issues.append(f"Date: {format_date(d['date_raw']) or '-'} ≠ {format_date(ref['date_raw']) or '-'}")
                    flags.append((date_col, ref_out))

            r['Check Result'] = "; ".join(issues) if issues else "OK"
            out = add_row(r, 'matched')
            for col, g_out in flags:
                flag(out, col, g_out)

        add_row({})

    # ---- 4. PDF items that match no Glovia row ----
    for po_idx, po in enumerate(po_items):
        if po_idx in used:
            continue
        delivs = parse_pdf_deliveries(po)
        for j, d in enumerate(delivs):
            r = build_pdf_row(po, d, j, len(delivs))
            r['Check Result'] = "Only in PDF - not found in GloviaG2"
            add_row(r, 'pdf_only')
        add_row({})

    # ---- 5. Write & style ----
    df_final = pd.DataFrame(structured_rows, columns=out_cols)
    buf = io.BytesIO()
    df_final.to_excel(buf, index=False)
    buf.seek(0)

    wb = openpyxl.load_workbook(buf)
    ws = wb.active
    col_pos = {name: i + 1 for i, name in enumerate(out_cols)}
    for r_idx, style in row_styles.items():
        for c in range(1, len(out_cols) + 1):
            ws.cell(row=r_idx + 2, column=c).fill = STYLE_FILLS[style]
    for r_idx, col in red_cells:   # red is applied last so it always wins over gray/yellow
        if col in col_pos:
            ws.cell(row=r_idx + 2, column=col_pos[col]).fill = FILL_RED
    ws.freeze_panes = 'A2'

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    stats = {
        "blocks": len(blocks),
        "rows_with_discrepancy": len({r for r, _ in red_cells}),
        "missing_in_pdf": sum(1 for b in blocks if b['po_idx'] is None),
        "only_in_pdf": sum(1 for i in range(len(po_items)) if i not in used),
    }
    return out_buf, stats, warnings


# ==========================================
# GEMINI EXTRACTION
# ==========================================
schedule_item_schema = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "Split_Quantity": types.Schema(type=types.Type.STRING, description="Quantity for this specific delivery date as printed (e.g. '25 st')"),
        "Required_Date": types.Schema(type=types.Type.STRING, description="Delivery date of this batch as DD-MM-YYYY (e.g. '07-08-2026')"),
    },
    required=["Split_Quantity", "Required_Date"],
)

po_item_schema = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "Line": types.Schema(type=types.Type.STRING, description="Line / position number (e.g. '010/000' or '090/000')"),
        "Item": types.Schema(type=types.Type.STRING, description="The primary item number (e.g. 511154)"),
        "Customer_Item": types.Schema(type=types.Type.STRING, description="Buyer's part number labeled 'Uw teknr:' (blank if missing)"),
        "Description": types.Schema(type=types.Type.STRING, description="Item description (e.g. 'Set Point 4 & 5 disk vlgs tek')"),
        "Order_Quantity": types.Schema(type=types.Type.STRING, description="Total quantity of the line as printed (e.g. '75 st')"),
        "Unit_Price": types.Schema(type=types.Type.STRING, description="Price text incl. conditions (e.g. '€ 3.457,44 per 1 st')"),
        "Deliveries": types.Schema(
            type=types.Type.ARRAY,
            items=schedule_item_schema,
            description="ALL delivery dates with their quantities for this line item - one entry per date.",
        ),
    },
    required=["Item", "Unit_Price", "Deliveries"],
)

final_response_schema = types.Schema(type=types.Type.ARRAY, items=po_item_schema)

PROMPT = """
You are an elite Purchase Order / Order Confirmation parsing specialist handling multi-lingual layouts
(Dutch, English, German) that differ per supplier.

CRITICAL EXTRACTION RULES:
1. ONE OBJECT PER LINE ITEM:
   - Extract every line item, even if it has no customer part number ('Uw teknr').
   - Capture 'Line' (e.g. '090/000'), 'Item' (e.g. '511154'), 'Description', 'Order_Quantity' and 'Unit_Price'.
2. MULTIPLE DELIVERY DATES (Verzendschema / Leverschema / Delivery schedule / Lieferplan):
   - A line item can have SEVERAL delivery dates, each with its own quantity. They may appear as a schedule
     block under the line, as extra rows with only a date and quantity, or as repeated date columns.
   - Put EVERY delivery of that line in its 'Deliveries' array, in the printed order.
   - NEVER keep only the first date. NEVER combine several dates into one entry.
   - NEVER output the same line item as separate objects per delivery date.
3. If a line has only one delivery date, 'Deliveries' has exactly one entry with the full line quantity.
   If a date has no quantity printed next to it and it is the only date, use the line's total quantity.
4. Required_Date: always DD-MM-YYYY. If only a week number is given, write e.g. 'week 32 2026'.
5. Copy numbers exactly as printed, including thousands and decimal separators and units.
6. If these pages start with a delivery schedule whose line item header (Line/Item) is not visible
   in these pages, ignore that fragment.
"""


class QuotaExhausted(Exception):
    pass


def iter_page_chunks(total_pages, size, overlap):
    step = max(1, size - overlap)
    start = 0
    while start < total_pages:
        end = min(start + size, total_pages)
        yield start, end
        if end >= total_pages:
            break
        start += step


def call_gemini(client, chunk_bytes, max_retries=4, retry_delay=4):
    for attempt in range(max_retries):
        try:
            return client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[types.Part.from_bytes(data=chunk_bytes, mime_type='application/pdf'), PROMPT],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=final_response_schema,
                    temperature=0.0,
                ),
            )
        except Exception as api_err:
            err_msg = str(api_err)
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                raise QuotaExhausted() from api_err
            if ("503" in err_msg or "UNAVAILABLE" in err_msg) and attempt < max_retries - 1:
                time.sleep(retry_delay)
                retry_delay *= 2
                continue
            raise
    return None


# ==========================================
# STREAMLIT UI
# ==========================================
def main():
    st.set_page_config(page_title="PO Checker AI", layout="wide")
    st.title("📦 Purchase Order Checking Assistant")
    st.write("Upload your System Master Data and PO PDFs to automatically generate a flagged discrepancy report "
             "with matching confirmation notes (Supports Split Deliveries & Loose Text/Line Matching).")

    if "GOOGLE_API_KEY" not in st.secrets:
        st.error("🔑 Google API Key is not set! Please configure it within Streamlit Secrets.")
        st.stop()
    client = genai.Client(api_key=st.secrets["GOOGLE_API_KEY"])

    excel_file = st.file_uploader("👉 Step 1: Upload System Master Data (Excel or CSV)",
                                  type=["xlsx", "xls", "csv"], key="master_data_excel_csv")
    pdf_files = st.file_uploader("👉 Step 2: Upload PO PDF file(s)", type=["pdf"],
                                 accept_multiple_files=True, key="po_pdf_files_list")

    if not (excel_file and pdf_files):
        st.info("💡 Please upload both the Master Excel file and PO PDFs to begin.")
        return
    if not st.button("🚀 Run AI Verification Report"):
        return

    if excel_file.name.lower().endswith('.csv'):
        df_excel = pd.read_csv(excel_file, dtype=str)
    else:
        df_excel = pd.read_excel(excel_file, dtype=str)

    raw_items = []
    progress_bar = st.progress(0)
    for idx, pdf_file in enumerate(pdf_files):
        st.write(f"🔍 AI is processing file: **{pdf_file.name}**")
        try:
            pdf_file.seek(0)
            pdf_bytes = pdf_file.read()
            if not pdf_bytes:
                continue
            reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
            total_pages = len(reader.pages)

            for start_page, end_page in iter_page_chunks(total_pages, PAGES_PER_CHUNK, CHUNK_OVERLAP):
                st.write(f"   📄 Parsing pages {start_page + 1} to {end_page} (Total: {total_pages} pages)...")
                writer = pypdf.PdfWriter()
                for p_idx in range(start_page, end_page):
                    writer.add_page(reader.pages[p_idx])
                chunk_buffer = io.BytesIO()
                writer.write(chunk_buffer)

                try:
                    response = call_gemini(client, chunk_buffer.getvalue())
                except QuotaExhausted:
                    st.error("🛑 **Daily Quota Limit Reached (429)!**")
                    st.stop()

                if response and response.text:
                    try:
                        for item in json.loads(response.text.strip()):
                            item['PO_Source_File'] = pdf_file.name
                            raw_items.append(item)
                    except json.JSONDecodeError:
                        st.warning(f"⚠️ Failed to parse JSON for pages {start_page + 1}-{end_page}, skipping chunk.")
                time.sleep(2)
        except Exception as e:
            st.error(f"❌ Failed to parse {pdf_file.name}: {e}")
        progress_bar.progress((idx + 1) / len(pdf_files))

    if not raw_items:
        st.error("❌ No data was extracted from your PDF items. Processing stopped.")
        st.stop()

    po_items = merge_po_items(raw_items)
    multi = sum(1 for p in po_items if len(p['Deliveries']) > 1)
    st.write(f"🔄 {len(po_items)} PDF line items found ({multi} with multiple delivery dates). Aligning rows...")

    report, stats, warnings = build_report(df_excel, po_items)
    for w in warnings:
        st.warning(f"⚠️ {w}")

    st.success(f"🎉 Process Complete! {stats['rows_with_discrepancy']} PDF row(s) with discrepancies, "
               f"{stats['missing_in_pdf']} item(s) missing in PDF, {stats['only_in_pdf']} item(s) only in PDF.")
    st.download_button(
        label="📥 Download Discrepancy Report",
        data=report,
        file_name=OUTPUT_FILENAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


main()
