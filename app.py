# ==========================================
# 修正：具備日期容錯與解析能力的底色標記邏輯
# ==========================================

# 1. 建立一個強大的日期標準化工具函數
def normalize_to_date_obj(date_str):
    """將任何常見日期格式轉換為統一的 date 物件，無法轉換則回傳 None"""
    if pd.isna(date_str) or str(date_str).strip() == "":
        return None
    
    # 清理可能的干擾文字
    s = str(date_str).replace("wk:", "").replace("week", "").strip()
    # 嘗試用 pandas 解析日期
    try:
        # dayfirst=True 處理歐式日期(DD/MM/YYYY)
        return pd.to_datetime(s, dayfirst=True).date()
    except:
        return None

# ... (在原本的檢查迴圈中) ...

                # 5. 日期比對邏輯 (關鍵優化)
                if idx_date:
                    d_cell = ws.cell(row=pdf_row_num, column=idx_date)
                    
                    # 將兩邊的日期都轉換為標準 date 物件
                    ex_date_obj = normalize_to_date_obj(ex_date_val)
                    pdf_date_obj = normalize_to_date_obj(d_cell.value)
                    
                    # 比較邏輯：
                    # 只有當「兩邊都有值」且「兩邊值不相等」時，才標註底色
                    if ex_date_obj is not None and pdf_date_obj is not None:
                        if ex_date_obj != pdf_date_obj:
                            ws.cell(row=glovia_row_num, column=idx_date).fill = fill_red
                            d_cell.fill = fill_red
                    # 如果 Excel 是空的，或者 PDF 解析不出日期，則視為「無需標示」，避免誤標
