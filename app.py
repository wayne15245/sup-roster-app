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

def process_roster(uploaded_file):
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    
    # 用來存放所有數據的列表
    data_list = []

    for sheet in workbook.worksheets:
        # 讀取日期(B2:H2)與時間(B3:H3)
        dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
        times = [sheet.cell(row=3, column=c).value for c in range(2, 9)]

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value in TEAMS:
                    # 檢查 B:H 欄位的 SUP
                    for col_idx in range(2, 9):
                        val = sheet.cell(row=cell.row, column=col_idx).value
                        if val and str(val).strip() != "":
                            data_list.append({
                                "Date": dates[col_idx-2],
                                "Time": times[col_idx-2],
                                "Count": 1
                            })

    df = pd.DataFrame(data_list)
    # 使用 pivot_table，確保結構為：時間為行，日期為列
    pivot_df = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
    
    # 這裡調整為：將原本的 Index (Time) 轉為第一欄
    pivot_df = pivot_df.reset_index()
    return pivot_df

uploaded_file = st.file_uploader("請上傳 Roster Excel (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        try:
            result = process_roster(uploaded_file)
            st.dataframe(result)
            
            # 使用 utf-8-sig 編碼，確保 CSV 在 Excel 打開時中文不會亂碼
            csv = result.to_csv(index=False).encode('utf-8-sig')
            st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
        except Exception as e:
            st.error(f"錯誤: {e}")