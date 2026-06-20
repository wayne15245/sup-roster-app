import streamlit as st
import pandas as pd
import openpyxl
import re

st.set_page_config(page_title="SUP 統計系統 (精準版)", layout="wide")
st.title("📊 SUP 人數統計系統 (僅處理 2-7 頁)")

# 嚴格的時間格式：例如 0730-1630
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')
TARGET_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]

def get_merged_cell_value(sheet, row, col):
    for merged_range in sheet.merged_cells.ranges:
        if row >= merged_range.min_row and row <= merged_range.max_row and \
           col >= merged_range.min_col and col <= merged_range.max_col:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return sheet.cell(row=row, column=col).value

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # 1. 先讀取 Pattern 分頁 (取得 SUP 狀態)
        if 'Pattern' not in workbook.sheetnames:
            st.error("找不到 'Pattern' 分頁！")
            return None
        
        pattern_ws = workbook['Pattern']
        # 建立 SUP 狀態表：{row_index: True/False}
        sup_map = {r: False for r in TARGET_ROWS}
        for r in TARGET_ROWS:
            j_val = get_merged_cell_value(pattern_ws, r, 10)
            if j_val and str(j_val).strip() not in ["None", "", "0"]:
                sup_map[r] = True

        # 2. 處理 Roster 分頁 (精準控制在 2-7 頁)
        # workbook.worksheets[1:7] 對應 index 1 到 6，即第 2 到第 7 個分頁
        all_records = []
        target_sheets = workbook.worksheets[1:7]
        
        for sheet in target_sheets:
            # 取得日期 (第2行)
            dates = [get_merged_cell_value(sheet, 2, c) for c in range(2, 9)]
            
            for r in TARGET_ROWS:
                # 檢查這行是否有 SUP 標記
                if sup_map.get(r, False):
                    # 讀取 B~H 欄的時間
                    for i, col in enumerate(range(2, 9)):
                        val = get_merged_cell_value(sheet, r, col)
                        if val:
                            t_str = str(val).strip().replace(" ", "")
                            # 🚨 關鍵篩選：只有完全符合時間格式的字串才會被計算
                            if TIME_PATTERN.match(t_str):
                                all_records.append({
                                    "Date": str(dates[i]),
                                    "Time": t_str,
                                    "Count": 1
                                })

        if not all_records: 
            st.warning("沒有抓取到資料，請確認 2-7 頁的 J 欄是否有標記。")
            return None
            
        df = pd.DataFrame(all_records)
        pivot = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        return pivot

    except Exception as e:
        st.error(f"分析錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")