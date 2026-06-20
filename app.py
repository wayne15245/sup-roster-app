import streamlit as st
import pandas as pd
import openpyxl
import datetime

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統 (簡寫對應版)")

# 1. 建立對照表：將 Excel 中的簡寫，對應回你要的 Team 名稱
TEAM_MAPPING = {
    "C": "Team C", "D": "Team D", "E": "Team E", "F": "Team F", "G": "Team G", "H": "Team H", "J": "Team J",
    "L1": "Team L1", "L2": "Team L2", "L3": "Team L3", "L4": "Team L4", "L5": "Team L5", "L6": "Team L6",
    "W": "Team W", "X": "Team X", "Y": "Team Y", "Z": "Team Z",
    "N1": "Team N1", "N2": "Team N2", "N3": "Team N3", "N4": "Team N4", "N5": "Team N5", "N6": "Team N6", "N7": "Team N7"
}

IGNORE_WORDS = ["DO", "AL", "PH", "SL", "NIL", "OFF", "REST", "V", ""]

def get_merged_cell_value(sheet, row, col):
    """自動偵測合併儲存格並返回正確值"""
    for merged_range in sheet.merged_cells.ranges:
        if row >= merged_range.min_row and row <= merged_range.max_row and \
           col >= merged_range.min_col and col <= merged_range.max_col:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return sheet.cell(row=row, column=col).value

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        all_records = []
        
        # 只處理第 2 到第 7 頁 (Index 1 到 6)
        for sheet in workbook.worksheets[1:7]:
            
            # 1. 取得日期列 (第 2 行)
            dates = []
            for c in range(2, 9):
                val = get_merged_cell_value(sheet, 2, c)
                if isinstance(val, datetime.datetime):
                    dates.append(val.strftime("%Y-%m-%d"))
                else:
                    dates.append(str(val).strip() if val else "未知日期")

            current_team = None

            # 2. 逐行掃描
            for row in range(3, sheet.max_row + 1):
                # 檢查 A 欄 (簡寫)
                cell_a = str(get_merged_cell_value(sheet, row, 1)).strip()
                
                # 如果這行是我們對照表中的簡寫 (如 'C', 'L1')，就切換目前的 Team
                if cell_a in TEAM_MAPPING:
                    current_team = TEAM_MAPPING[cell_a]
                
                # 如果已經在某個 Team 內，檢查 SUP
                if current_team:
                    # I 欄(9) 或 J 欄(10) 有數值代表是 SUP
                    s_val = get_merged_cell_value(sheet, row, 9)
                    d_val = get_merged_cell_value(sheet, row, 10)
                    
                    if (s_val and str(s_val).strip() not in ["None", ""]) or \
                       (d_val and str(d_val).strip() not in ["None", ""]):
                        
                        # 抓取該行的時間 (B~H)
                        for i, col in enumerate(range(2, 9)):
                            time_val = get_merged_cell_value(sheet, row, col)
                            if time_val:
                                t_str = str(time_val).replace(" ", "").upper()
                                # 排除休假，只紀錄真實上班時間
                                if t_str not in IGNORE_WORDS and len(t_str) > 2:
                                    all_records.append({
                                        "Date": dates[i],
                                        "Time": t_str,
                                        "Count": 1
                                    })
                                    
        if not all_records:
            return None
            
        df = pd.DataFrame(all_records)
        pivot = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        return pivot

    except Exception as e:
        st.error(f"分析過程出錯: {e}")
        return None

# --- 介面 ---
uploaded_file = st.file_uploader("請重新上傳 .xlsm 檔案", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        result = process_roster(uploaded_file)
        
        if result is not None:
            st.success("統計完成！")
            st.dataframe(result)
            csv = result.to_csv().encode('utf-8-sig')
            st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
        else:
            st.warning("無法抓取數據。請檢查：\n1. A 欄是否為正確的簡寫（如 C, L1, N7）\n2. I 或 J 欄是否有標記\n3. 時間欄位是否為正確時間格式")