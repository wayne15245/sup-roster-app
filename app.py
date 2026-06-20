import streamlit as st
import pandas as pd
import openpyxl
import datetime

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統 (防重疊修正版)")

TEAM_MAPPING = {
    "C": "Team C", "D": "Team D", "E": "Team E", "F": "Team F", "G": "Team G", "H": "Team H", "J": "Team J",
    "L1": "Team L1", "L2": "Team L2", "L3": "Team L3", "L4": "Team L4", "L5": "Team L5", "L6": "Team L6",
    "W": "Team W", "X": "Team X", "Y": "Team Y", "Z": "Team Z",
    "N1": "Team N1", "N2": "Team N2", "N3": "Team N3", "N4": "Team N4", "N5": "Team N5", "N6": "Team N6", "N7": "Team N7"
}

IGNORE_WORDS = ["DO", "AL", "PH", "SL", "NIL", "OFF", "REST", "V", ""]

def get_merged_info(sheet, row, col):
    """回傳 (值, min_row, min_col) 用來辨識是否為同一個合併區塊"""
    for merged_range in sheet.merged_cells.ranges:
        if row >= merged_range.min_row and row <= merged_range.max_row and \
           col >= merged_range.min_col and col <= merged_range.max_col:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value, merged_range.min_row, merged_range.min_col
    return sheet.cell(row=row, column=col).value, row, col

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        all_records = []
        
        for sheet in workbook.worksheets[1:7]:
            # 取得日期
            dates = []
            for c in range(2, 9):
                val, _, _ = get_merged_info(sheet, 2, c)
                if isinstance(val, datetime.datetime):
                    dates.append(val.strftime("%Y-%m-%d"))
                else:
                    dates.append(str(val).strip() if val else "未知日期")

            current_team = None
            # 紀錄已經計算過的 (min_row, min_col, date_index) 以防止重複
            counted_blocks = set()

            for row in range(3, sheet.max_row + 1):
                # 檢查 Team
                cell_a_val, _, _ = get_merged_info(sheet, row, 1)
                cell_a = str(cell_a_val).strip()
                if cell_a in TEAM_MAPPING:
                    current_team = TEAM_MAPPING[cell_a]
                
                if current_team:
                    # 檢查 I(9) 或 J(10)
                    for col_idx in [9, 10]:
                        s_val, r_min, c_min = get_merged_info(sheet, row, col_idx)
                        if s_val and str(s_val).strip() not in ["None", ""]:
                            
                            # 這邊是 SUP，檢查當天各欄位
                            for i, col in enumerate(range(2, 9)):
                                time_val, _, _ = get_merged_info(sheet, row, col)
                                t_str = str(time_val).replace(" ", "").upper()
                                
                                # 只要這個合併儲存格區塊在當天還沒被算過，就加入
                                block_key = (r_min, c_min, i) 
                                if t_str not in IGNORE_WORDS and len(t_str) > 2 and block_key not in counted_blocks:
                                    all_records.append({
                                        "Date": dates[i],
                                        "Time": t_str,
                                        "Count": 1
                                    })
                                    counted_blocks.add(block_key)
                                    
        if not all_records: return None
            
        df = pd.DataFrame(all_records)
        pivot = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        return pivot

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