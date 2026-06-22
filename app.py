import streamlit as st
import pandas as pd
import openpyxl
import re
from datetime import datetime

st.set_page_config(page_title="SUP 強力清洗系統", layout="wide")
st.title("📊 SUP 精準統計系統 (清洗版)")

SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

def get_clean_date(val):
    """嘗試將值轉換為 YYYY-MM-DD 格式，若失敗則回傳 None"""
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    if isinstance(val, str):
        # 嘗試找尋字串中的 YYYY-MM-DD 模式
        match = re.search(r'\d{4}-\d{2}-\d{2}', val)
        if match:
            return match.group(0)
    return None

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            # 【關鍵步驟】：預先清洗日期行，確保欄位對應正確
            raw_dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            clean_dates = [get_clean_date(d) for d in raw_dates]
            
            for r in SUP_ROWS:
                for col_idx, col in enumerate(range(2, 9)):
                    # 如果該欄日期本身就是無效的，直接跳過
                    if clean_dates[col_idx] is None:
                        continue
                        
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        if TIME_PATTERN.match(t_str):
                            all_records.append({
                                "Date": clean_dates[col_idx], 
                                "Time": t_str, 
                                "Count": 1
                            })
        
        if not all_records:
            return None
            
        df = pd.DataFrame(all_records)
        # 建立 Pivot Table (此時 DataFrame 已無任何雜訊)
        result = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        return result

    except Exception as e:
        st.error(f"❌ 系統錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("成功匯出乾淨報表！")
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_精準統計結果.csv", "text/csv")
    else:
        st.warning("⚠️ 未檢測到符合格式的資料，請確認行數或日期欄位。")