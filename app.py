import streamlit as st
import pandas as pd
import openpyxl
import re

st.set_page_config(page_title="SUP 診斷系統", layout="wide")
st.title("📊 SUP 統計系統 (診斷模式)")

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
        sheets = workbook.worksheets
        st.write(f"📂 系統偵測到檔案共有 {len(sheets)} 個分頁。")
        
        if 'Pattern' not in workbook.sheetnames:
            st.error("❌ 找不到名為 'Pattern' 的分頁，請檢查 Excel 分頁名稱。")
            return None
        
        pattern_ws = workbook['Pattern']
        
        # 診斷：顯示 Pattern 頁面的 SUP 標記狀況
        st.write("🔍 正在掃描 'Pattern' 分頁的 J 欄標記...")
        sup_map = {}
        for r in TARGET_ROWS:
            val = get_merged_cell_value(pattern_ws, r, 10)
            st.write(f"   -> 第 {r} 行 J 欄數值: '{val}'")
            if val and str(val).strip() not in ["None", "", "0"]:
                sup_map[r] = True
            else:
                sup_map[r] = False
        
        all_records = []
        # 只處理第 2 至 第 7 個分頁 (Index 1 到 6)
        target_sheets = sheets[1:7]
        st.write(f"🔢 正在處理分頁索引 1 到 6 (第 2-7 頁)...")
        
        for sheet in target_sheets:
            st.write(f"正在掃描分頁: {sheet.title}")
            dates = [get_merged_cell_value(sheet, 2, c) for c in range(2, 9)]
            
            for r in TARGET_ROWS:
                if sup_map.get(r, False):
                    for i, col in enumerate(range(2, 9)):
                        val = get_merged_cell_value(sheet, r, col)
                        if val:
                            t_str = str(val).strip().replace(" ", "")
                            if TIME_PATTERN.match(t_str):
                                all_records.append({"Date": str(dates[i]), "Time": t_str, "Count": 1})
                            else:
                                pass # 這裡不顯示錯誤以免洗版

        if not all_records:
            st.warning("⚠️ 統計結果為空。請檢查上述『J 欄數值』是否全為 False 或空值。")
            return None
            
        return pd.DataFrame(all_records).pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"❌ 系統發生錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("統計完成！")
        st.dataframe(result)