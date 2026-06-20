import streamlit as st
import pandas as pd
import openpyxl

# 設定網頁標題
st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統")

# 定義要統計的 24 個 Team
TEAMS = [
    "TeamC", "TeamD", "TeamE", "TeamF", "TeamG", "TeamH",
    "TeamL1", "TeamL2", "TeamL3", "TeamL4", "TeamL5", "TeamL6",
    "TeamW", "TeamX", "TeamY", "TeamZ",
    "TeamN1", "TeamN2", "TeamN3", "TeamN4", "TeamN5", "TeamN6", "TeamN7",
    "TeamJ"
]

uploaded_file = st.file_uploader("請上傳 Roster Excel (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        try:
            # 讀取 Excel 檔案
            workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
            results = []

            # 遍歷所有分頁
            for sheet in workbook.worksheets:
                # 在分頁中尋找關鍵字
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value in TEAMS:
                            team_name = cell.value
                            # 假設 SUP 在該 Team 標題的 I 和 J 欄位 (根據你的描述進行偏移)
                            # 這裡需要根據你 Excel 的實際偏移量進行調整
                            # 例如：SUP 在該儲存格右方 8 格 (I) 與 9 格 (J)
                            sup_cells = [sheet.cell(row=cell.row, column=cell.column+8), 
                                         sheet.cell(row=cell.row, column=cell.column+9)]
                            
                            sup_count = sum(1 for c in sup_cells if c.value is not None and str(c.value).strip() != "")
                            
                            results.append({
                                "Team": team_name,
                                "SUP_Count": sup_count
                            })

            # 整理成 DataFrame
            df = pd.DataFrame(results)
            pivot_df = df.groupby("Team")["SUP_Count"].sum().reset_index()

            st.success("統計完成！")
            st.dataframe(pivot_df)

            # 下載統計報表
            csv = pivot_df.to_csv(index=False).encode('utf-8')
            st.download_button("下載統計結果 CSV", csv, "SUP_Report.csv", "text/csv")

        except Exception as e:
            st.error(f"讀取錯誤: {e}")