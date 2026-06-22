import streamlit as st
import pandas as pd
import openpyxl
import re

# 【精準鎖定】只看我們指定的行，無視周圍的所有雜訊
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        # 只取第 2 至 第 7 個分頁
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            # 讀取第 2 行的日期
            dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            for r in SUP_ROWS:
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        # 【過濾機制】只抓取符合時間格式的資料
                        if TIME_PATTERN.match(t_str):
                            # 確保該欄位對應的是一個有效的日期
                            date_val = str(dates[col_idx])
                            all_records.append({"Date": date_val, "Time": t_str, "Count": 1})
        
        df = pd.DataFrame(all_records)
        
        # 【關鍵防呆】徹底刪除所有包含非日期字串的欄位
        # 這裡假設你的正確日期格式包含 '-' (例如 2026-07-06)
        # 任何不含 '-' 的標籤 (如 Fri, Mon) 都會被這行過濾掉
        df = df[df['Date'].str.contains('-', na=False)]
        
        # 建立 Pivot Table
        pivot = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        
        return pivot

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_穩定版結果.csv", "text/csv")