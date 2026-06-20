import streamlit as st
import pandas as pd
import openpyxl
import re

st.set_page_config(page_title="SUP 最終統計系統", layout="wide")
st.title("📊 SUP 人數統計系統 (J欄偵錯版)")

# 嚴格定義：HHMM-HHMM 格式
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')
TARGET_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]

def get_merged_cell_value(sheet, row, col):
    """處理合併儲存格的工具函數"""
    for merged_range in sheet.merged_cells.ranges:
        if row >= merged_range.min_row and row <= merged_range.max_row and \
           col >= merged_range.min_col and col <= merged_range.max_col:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return sheet.cell(row=row, column=col).value

def process_roster(uploaded_file):
    try:
        # data_only=True 讀取計算後的值
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        if 'Pattern' not in workbook.sheetnames:
            st.error("❌ 找不到 'Pattern' 分頁！")
            return None
        
        pattern_ws = workbook['Pattern']
        st.write("🔍 **正在偵測 Pattern 分頁 J 欄 (第 10 欄)...**")
        
        # 建立 SUP 標記地圖
        sup_map = {}
        found_rows = []
        for r in TARGET_ROWS:
            val = get_merged_cell_value(pattern_ws, r, 10)
            # 轉換為字串並去除空白，方便比對
            str_val = str(val).strip() if val is not None else ""
            st.write(f"- 第 {r} 行, J 欄數值: `{val}` (判定: {'✅ 有效' if str_val == '1' else '❌ 無效'})")
            
            if str_val == '1':
                sup_map[r] = True
                found_rows.append(r)
            else:
                sup_map[r] = False
        
        if not found_rows:
            st.warning("⚠️ 系統未偵測到任何 J 欄的 '1'，請檢查 Excel 是否已儲存計算結果。")
            return None

        # 處理資料分頁
        all_records = []
        # 僅處理第 2 到第 7 個分頁
        target_sheets = workbook.worksheets[1:7]
        
        for sheet in target_sheets:
            dates = [get_merged_cell_value(sheet, 2, c) for c in range(2, 9)]
            for r in TARGET_ROWS:
                if sup_map.get(r, False): # 只有被 Pattern 標記過的行才處理
                    for i, col in enumerate(range(2, 9)):
                        val = get_merged_cell_value(sheet, r, col)
                        if val:
                            t_str = str(val).strip().replace(" ", "")
                            if TIME_PATTERN.match(t_str):
                                all_records.append({
                                    "Date": str(dates[i]),
                                    "Time": t_str,
                                    "Count": 1
                                })

        if not all_records:
            st.warning("⚠️ 沒有抓取到資料，可能分頁範圍不對。")
            return None
            
        return pd.DataFrame(all_records).pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("統計完成！")
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")