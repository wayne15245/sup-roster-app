import streamlit as st
import pandas as pd
import openpyxl
import datetime

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統 (穩定版)")

# 定義 24 個 Team
TEAMS = [
    "Team C", "Team D", "Team E", "Team F", "Team G", "Team H", "Team J",
    "Team L1", "Team L2", "Team L3", "Team L4", "Team L5", "Team L6",
    "Team W", "Team X", "Team Y", "Team Z",
    "Team N1", "Team N2", "Team N3", "Team N4", "Team N5", "Team N6", "Team N7"
]

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
        
        # 處理第 2 到第 7 頁
        # workbook.worksheets 的索引是 0 開始，所以第 2 頁是 index 1
        for sheet in workbook.worksheets[1:7]:
            
            # 1. 取得日期列 (第 2 行)
            dates = []
            for c in range(2, 9): # B~H 欄
                val = get_merged_cell_value(sheet, 2, c)
                if isinstance(val, datetime.datetime):
                    dates.append(val.strftime("%Y-%m-%d"))
                else:
                    dates.append(str(val).strip() if val else "未知日期")

            current_team = None

            # 2. 逐行掃描
            for row in range(3, sheet.max_row + 1):
                # 檢查 Team 名稱 (A 欄)
                team_cell = str(get_merged_cell_value(sheet, row, 1)).strip()
                if team_cell in TEAMS:
                    current_team = team_cell
                
                if current_team:
                    # 檢查 SUP 標記 (I 欄=9, J 欄=10)
                    s_val = get_merged_cell_value(sheet, row, 9)
                    d_val = get_merged_cell_value(sheet, row, 10)
                    
                    if (s_val and str(s_val).strip() not in ["None", ""]) or \
                       (d_val and str(d_val).strip() not in ["None", ""]):
                        
                        # 抓取這週的時間 (B~H)
                        for i, col in enumerate(range(2, 9)):
                            time_val = get_merged_cell_value(sheet, row, col)
                            if time_val:
                                t_str = str(time_val).replace(" ", "").upper()
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
uploaded_file = st.file_uploader("請重新上傳您的 .xlsm 檔案", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        result = process_roster(uploaded_file)
        
        if result is not None:
            st.success("統計完成！")
            st.dataframe(result)
            csv = result.to_csv().encode('utf-8-sig')
            st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
        else:
            st.warning("無法從檔案中抓取數據。請檢查：\n1. A 欄是否有正確填寫『Team X』\n2. I 或 J 欄是否有輸入標記\n3. 檔案內容是否符合格式")