import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統")

TEAMS = [
    "Team C", "Team D", "Team E", "Team F", "Team G", "Team H",
    "Team L1", "Team L2", "Team L3", "Team L4", "Team L5", "Team L6",
    "Team W", "Team X", "Team Y", "Team Z",
    "Team N1", "Team N2", "Team N3", "Team N4", "Team N5", "Team N6", "Team N7",
    "Team J"
]

def clean_time(time_val):
    """將所有時間格式強制轉換為標準格式，例如移除空格"""
    if not time_val: return None
    t = str(time_val).replace(" ", "").strip()
    return t

def process_roster(uploaded_file):
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    all_records = []

    for sheet in workbook.worksheets:
        # 讀取日期(Row 2)與時間(Row 3)，使用 clean_time 標準化時間
        dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
        times = [clean_time(sheet.cell(row=3, column=c).value) for c in range(2, 9)]

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value in TEAMS:
                    # 檢查該 Team 在 B~H 欄位的 SUP
                    for i, col_idx in enumerate(range(2, 9)):
                        sup_val = sheet.cell(row=cell.row, column=col_idx).value
                        if sup_val and str(sup_val).strip() != "":
                            all_records.append({
                                "Date": dates[i],
                                "Time": times[i],
                                "Count": 1
                            })
    
    if not all_records: return None
        
    df = pd.DataFrame(all_records)
    # 使用 pivot_table：時間為 Index，日期為 Column
    pivot_df = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
    
    # 強制排序時間（如果你有特定的順序要求，可以在這裡指定）
    return pivot_df

uploaded_file = st.file_uploader("請上傳 Roster Excel (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        try:
            result = process_roster(uploaded_file)
            # 顯示結果
            st.dataframe(result)
            csv = result.to_csv().encode('utf-8-sig')
            st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
        except Exception as e:
            st.error(f"錯誤: {e}")