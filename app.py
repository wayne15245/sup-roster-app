import streamlit as st
import pandas as pd
import openpyxl
import re
from datetime import datetime

st.set_page_config(page_title="SUP 強制清洗系統", layout="wide")
st.title("📊 SUP 精準統計系統 (徹底過濾版)")

SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

def get_clean_date(val):
    """只抓取 YYYY-MM-DD，其餘一律視為垃圾資料"""
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    if isinstance(val, str):
        match = re.search(r'\d{4}-\d{2}-\d{2}', val)
        if match:
            return match.group(0)
    # 只要不是日期格式，一律回傳 None
    return None

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            # 1. 讀取並清洗日期列 (B2 到 H2)
            raw_dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            clean_dates = [get_clean_date(d) for d in raw_dates]
            
            # 2. 檢查每一行
            for r in SUP_ROWS:
                for col_idx in range(len(clean_dates)):
                    # 【核心過濾】：如果該欄位的日期是 None，直接跳過這整欄，絕不處理
                    if clean_dates[col_idx] is None:
                        continue
                        
                    val = sheet.cell(row=r, column=col+2).value # 修正 col 索引對應
                    # 重新計算 col 為正確索引 (B欄是2)
                    col_num = col_idx + 2 
                    val = sheet.cell(row=r, column=col_num).value
                    
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
        # 此時 df 絕對不會有 'Mon' 或 'Fri'，因為我們在前面就過濾掉了
        result = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        return result

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("統計完成，已排除所有雜訊欄位！")
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_精準統計結果.csv", "text/csv")
    else:
        st.warning("⚠️ 未檢測到資料。請確認 Excel 分頁 B2-H2 欄位是否為日期格式。")