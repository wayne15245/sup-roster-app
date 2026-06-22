import streamlit as st
import pandas as pd
import openpyxl
import re

# 這裡保留了 N1-N7 最成功的靈魂：精準的行數控制
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
# 這裡加入過濾機制，確保只抓時間格式
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        # 只取第 2 至 第 7 個分頁，避免抓到不必要的 sheet
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            # 日期固定在第 2 行
            dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            for r in SUP_ROWS:
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        # 【核心過濾】只有符合 0000-0000 格式的時間才抓取
                        if TIME_PATTERN.match(t_str):
                            all_records.append({
                                "Date": str(dates[col_idx]), 
                                "Time": t_str, 
                                "Count": 1
                            })
        
        df = pd.DataFrame(all_records)
        # 【過濾雜訊】只保留包含 '-' 的日期欄位，徹底剔除 'Mon', 'Fri' 等字串
        df = df[df['Date'].str.contains('-', na=False)]
        
        return df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_穩定統計結果.csv", "text/csv")