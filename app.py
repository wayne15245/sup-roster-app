import streamlit as st
import pandas as pd
import openpyxl
import datetime

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統 (精準定位版)")

# 定義目標行 (根據你提供的資訊)
TARGET_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]

TEAM_MAPPING = {
    "C": "Team C", "D": "Team D", "E": "Team E", "F": "Team F", "G": "Team G", "H": "Team H", "J": "Team J",
    "L1": "Team L1", "L2": "Team L2", "L3": "Team L3", "L4": "Team L4", "L5": "Team L5", "L6": "Team L6",
    "W": "Team W", "X": "Team X", "Y": "Team Y", "Z": "Team Z",
    "N1": "Team N1", "N2": "Team N2", "N3": "Team N3", "N4": "Team N4", "N5": "Team N5", "N6": "Team N6", "N7": "Team N7"
}

IGNORE_WORDS = ["DO", "AL", "PH", "SL", "NIL", "OFF", "REST", "V", ""]

def get_merged_cell_value(sheet, row, col):
    """處理合併儲存格"""
    for merged_range in sheet.merged_cells.ranges:
        if row >= merged_range.min_row and row <= merged_range.max_row and \
           col >= merged_range.min_col and col <= merged_range.max_col:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return sheet.cell(row=row, column=col).value

def process_roster(uploaded_file):
    try:
        # data_only=True 可以讀取 Formula 計算後的結果
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        all_records = []
        
        # 處理第 2 到第 7 頁 (Index 1 到 6)
        for sheet in workbook.worksheets[1:7]:
            
            # 1. 取得日期列 (第 2 行)
            dates = []
            for c in range(2, 9): # B~H 欄
                val = get_merged_cell_value(sheet, 2, c)
                if isinstance(val, datetime.datetime):
                    dates.append(val.strftime("%Y-%m-%d"))
                else:
                    dates.append(str(val).strip() if val else "未知日期")

            # 2. 針對目標行進行精準掃描
            for r in TARGET_ROWS:
                # 確認 Team (A 欄)
                team_cell = str(get_merged_cell_value(sheet, r, 1)).strip()
                
                # 如果這行有對應到我們要的 Team
                if team_cell in TEAM_MAPPING:
                    team_name = TEAM_MAPPING[team_cell]
                    
                    # 檢查 J 欄 (第 10 欄) 是否有標記
                    j_val = get_merged_cell_value(sheet, r, 10)
                    
                    # 如果有標記 (不為空且非 None)
                    if j_val and str(j_val).strip() not in ["None", ""]:
                        
                        # 讀取當週時間 (B~H 欄)
                        for i, col in enumerate(range(2, 9)):
                            time_val = get_merged_cell_value(sheet, r, col)
                            if time_val:
                                t_str = str(time_val).replace(" ", "").upper()
                                if t_str not in IGNORE_WORDS and len(t_str) > 2:
                                    all_records.append({
                                        "Date": dates[i],
                                        "Time": t_str,
                                        "Count": 1
                                    })
                                    
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
    else:
        st.warning("未讀取到資料，請確認 J 欄是否有輸入標記。")