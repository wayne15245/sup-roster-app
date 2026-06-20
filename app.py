import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統")

# 更新後的 Team 名稱格式
TEAMS = [
    "Team C", "Team D", "Team E", "Team F", "Team G", "Team H",
    "Team L1", "Team L2", "Team L3", "Team L4", "Team L5", "Team L6",
    "Team W", "Team X", "Team Y", "Team Z",
    "Team N1", "Team N2", "Team N3", "Team N4", "Team N5", "Team N6", "Team N7",
    "Team J"
]

def process_roster(uploaded_file):
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    all_records = []

    for sheet in workbook.worksheets:
        # 1. 強制轉換日期與時間為文字，防止 Excel 自動格式化
        # 我們直接讀取 cell.value 並轉為 str，確保它保持原本的樣貌
        dates = [str(sheet.cell(row=2, column=c).value).strip() for c in range(2, 9)]
        times = [str(sheet.cell(row=3, column=c).value).strip() for c in range(2, 9)]

        # 2. 搜尋 Team 區域
        for row in sheet.iter_rows():
            for cell in row:
                # 這裡比對時也將 Excel 中的 Team 名稱標準化 (去除多餘空格)
                val = str(cell.value).strip()
                if val in TEAMS:
                    # 3. 檢查 B:H 欄位的 SUP
                    for i in range(7):
                        col_idx = i + 2
                        sup_val = sheet.cell(row=cell.row, column=col_idx).value
                        # 只要格內有字就計 1
                        if sup_val and str(sup_val).strip() not in ["None", ""]:
                            all_records.append({
                                "Date": dates[i],
                                "Time": times[i],
                                "Count": 1
                            })
    
    if not all_records: return None
        
    # 4. 統計數據
    df = pd.DataFrame(all_records)
    
    # 使用 Pivot Table 建立矩陣
    pivot_df = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
    
    # 移除索引名稱，讓 CSV 更乾淨
    pivot_df.index.name = "Time"
    return pivot_df

uploaded_file = st.file_uploader("請上傳 Roster Excel (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        try:
            result = process_roster(uploaded_file)
            st.dataframe(result)
            # 使用 utf-8-sig 確保中文 CSV 不會亂碼
            csv = result.to_csv().encode('utf-8-sig')
            st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
        except Exception as e:
            st.error(f"錯誤: {e}")