import streamlit as st
import pandas as pd
import openpyxl
import datetime
import re

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統 (嚴格篩選版)")

TEAM_MAPPING = {
    "C": "Team C", "D": "Team D", "E": "Team E", "F": "Team F", "G": "Team G", "H": "Team H", "J": "Team J",
    "L1": "Team L1", "L2": "Team L2", "L3": "Team L3", "L4": "Team L4", "L5": "Team L5", "L6": "Team L6",
    "W": "Team W", "X": "Team X", "Y": "Team Y", "Z": "Team Z",
    "N1": "Team N1", "N2": "Team N2", "N3": "Team N3", "N4": "Team N4", "N5": "Team N5", "N6": "Team N6", "N7": "Team N7"
}

IGNORE_WORDS = ["DO", "AL", "PH", "SL", "NIL", "OFF", "REST", "V", ""]

def is_valid_time_format(time_str):
    """嚴格檢查：只接受包含數字且符合時間格式的字串，過濾掉雜訊"""
    # 檢查是否為空
    if not time_str: return False
    # 簡單正規表示式檢查：是否有類似 HHMM-HHMM 的格式
    if re.search(r'\d{3,4}-\d{3,4}', time_str):
        return True
    return False

def get_merged_info(sheet, row, col):
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
            dates = []
            for c in range(2, 9):
                val, _, _ = get_merged_info(sheet, 2, c)
                dates.append(val.strftime("%Y-%m-%d") if isinstance(val, datetime.datetime) else "未知日期")

            current_team = None
            counted_blocks = set()

            for row in range(3, sheet.max_row + 1):
                # 檢查 A 欄
                cell_a_val, _, _ = get_merged_info(sheet, row, 1)
                cell_a = str(cell_a_val).strip()
                
                # 如果是新的 Team，切換狀態；如果不是，保持原本的 Team 狀態
                if cell_a in TEAM_MAPPING:
                    current_team = TEAM_MAPPING[cell_a]
                elif cell_a not in TEAM_MAPPING and cell_a != "":
                    # 如果遇到一個不屬於 Team 的行，重置 (防止抓到無關的 header 或結尾)
                    current_team = None

                if current_team:
                    # 只有在明確處於 Team 內，才檢查 I 或 J 欄
                    for col_idx in [9, 10]:
                        s_val, r_min, c_min = get_merged_info(sheet, row, col_idx)
                        if s_val and str(s_val).strip() not in ["None", ""]:
                            for i, col in enumerate(range(2, 9)):
                                time_val, _, _ = get_merged_info(sheet, row, col)
                                t_str = str(time_val).replace(" ", "").upper()
                                
                                # 💡 嚴格篩選：必須符合時間格式，且還沒算過，才納入計算
                                if is_valid_time_format(t_str) and (r_min, c_min, i) not in counted_blocks:
                                    all_records.append({"Date": dates[i], "Time": t_str, "Count": 1})
                                    counted_blocks.add((r_min, c_min, i))
                                    
        if not all_records: return None
        df = pd.DataFrame(all_records)
        return df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"分析錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("統計完成！")
        st.dataframe(result)
        st.download_button("下載統計表 CSV", result.to_csv().encode('utf-8-sig'), "SUP_統計結果.csv", "text/csv")