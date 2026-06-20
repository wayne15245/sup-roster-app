import streamlit as st
import pandas as pd
import openpyxl
import datetime

st.set_page_config(page_title="SUP 統計系統 (Pattern對照版)", layout="wide")
st.title("📊 SUP 人數統計系統 (Pattern 分頁對照)")

TARGET_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]

TEAM_MAPPING = {
    "C": "Team C", "D": "Team D", "E": "Team E", "F": "Team F", "G": "Team G", "H": "Team H", "J": "Team J",
    "L1": "Team L1", "L2": "Team L2", "L3": "Team L3", "L4": "Team L4", "L5": "Team L5", "L6": "Team L6",
    "W": "Team W", "X": "Team X", "Y": "Team Y", "Z": "Team Z",
    "N1": "Team N1", "N2": "Team N2", "N3": "Team N3", "N4": "Team N4", "N5": "Team N5", "N6": "Team N6", "N7": "Team N7"
}

IGNORE_WORDS = ["DO", "AL", "PH", "SL", "NIL", "OFF", "REST", "V", ""]

def get_merged_cell_value(sheet, row, col):
    for merged_range in sheet.merged_cells.ranges:
        if row >= merged_range.min_row and row <= merged_range.max_row and \
           col >= merged_range.min_col and col <= merged_range.max_col:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return sheet.cell(row=row, column=col).value

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # 1. 讀取 Pattern 分頁
        if 'Pattern' not in workbook.sheetnames:
            st.error("找不到 'Pattern' 分頁！請檢查分頁名稱是否正確。")
            return None
            
        pattern_ws = workbook['Pattern']
        # 建立 SUP 狀態表: 格式為 {sheet_name: {row_index: True/False}}
        sup_status = {} 
        
        # 假設 Pattern 分頁結構與其他分頁一致，我們遍歷所有分頁來紀錄 SUP 狀態
        for sheet in workbook.worksheets:
            if sheet.title == 'Pattern': continue
            
            sup_status[sheet.title] = {}
            for r in TARGET_ROWS:
                j_val = get_merged_cell_value(pattern_ws, r, 10) # 讀取 Pattern 分頁的 J 欄
                # 只要有內容就視為有 SUP
                if j_val and str(j_val).strip() not in ["None", "", "0"]:
                    sup_status[sheet.title][r] = True
                else:
                    sup_status[sheet.title][r] = False

        # 2. 處理 Roster 分頁進行統計
        all_records = []
        for sheet in workbook.worksheets:
            if sheet.title == 'Pattern': continue
            
            # 取得日期列 (第 2 行)
            dates = [get_merged_cell_value(sheet, 2, c) for c in range(2, 9)]
            
            for r in TARGET_ROWS:
                # 檢查該行在 Pattern 中是否有 SUP 標記
                if sup_status.get(sheet.title, {}).get(r, False):
                    
                    # 讀取當週時間 (B~H 欄)
                    for i, col in enumerate(range(2, 9)):
                        time_val = get_merged_cell_value(sheet, 2, col) # 這邊修正為抓取該行該列的時間
                        # 如果 time_val 是 datetime 物件，要轉成日期字串
                        current_date = dates[i].strftime("%Y-%m-%d") if isinstance(dates[i], datetime.datetime) else "未知日期"
                        
                        # 讀取實際上班時間
                        work_time = get_merged_cell_value(sheet, r, col)
                        if work_time:
                            t_str = str(work_time).replace(" ", "").upper()
                            if t_str not in IGNORE_WORDS and len(t_str) > 2:
                                all_records.append({"Date": current_date, "Time": t_str, "Count": 1})
                                    
        if not all_records: return None
        df = pd.DataFrame(all_records)
        return df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"分析錯誤: {e}")
        return None

# --- 介面 ---
uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("統計完成！")
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")