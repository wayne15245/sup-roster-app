import streamlit as st
import pandas as pd
import openpyxl
import re

# 【設定區】
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53] # 目標數據行
HEADER_ROW = 3 # 您現在想改為第 3 行
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            # 從指定的 HEADER_ROW 讀取標頭 (B3 到 H3)
            headers = [sheet.cell(row=HEADER_ROW, column=c).value for c in range(2, 9)]
            
            for r in SUP_ROWS:
                for col_idx in range(7):
                    header_val = str(headers[col_idx]) if headers[col_idx] else "未知"
                    
                    val = sheet.cell(row=r, column=col_idx + 2).value
                    if val:
                        t_str = str(val).strip()
                        if TIME_PATTERN.match(t_str):
                            all_records.append({
                                "Date/Day": header_val,
                                "Time": t_str,
                                "Count": 1
                            })
        
        if not all_records:
            return None
            
        df = pd.DataFrame(all_records)
        # 建立樞紐分析表
        result = df.pivot_table(index="Time", columns="Date/Day", values="Count", aggfunc="sum", fill_value=0)
        return result

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載 CSV", csv, "SUP_統計結果.csv", "text/csv")