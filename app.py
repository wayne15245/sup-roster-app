import streamlit as st
import pandas as pd
import openpyxl

# 設定網頁佈局
st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統")

# 定義 24 個 Team
TEAMS = [
    "TeamC", "TeamD", "TeamE", "TeamF", "TeamG", "TeamH",
    "TeamL1", "TeamL2", "TeamL3", "TeamL4", "TeamL5", "TeamL6",
    "TeamW", "TeamX", "TeamY", "TeamZ",
    "TeamN1", "TeamN2", "TeamN3", "TeamN4", "TeamN5", "TeamN6", "TeamN7",
    "TeamJ"
]

def process_roster(uploaded_file):
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    all_records = []

    for sheet in workbook.worksheets:
        # 1. 讀取該分頁的日期(Row 2, B~H)與時間(Row 3, B~H)
        # 範圍 B 到 H 對應 column 2 到 8
        dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
        times = [sheet.cell(row=3, column=c).value for c in range(2, 9)]

        # 2. 搜尋 Team 區域
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value in TEAMS:
                    # 檢查該 Team 對應的 B~H 欄位的 SUP 數據
                    for col_idx in range(2, 9):
                        sup_val = sheet.cell(row=cell.row, column=col_idx).value
                        # 若該格有值，計為 1
                        if sup_val and str(sup_val).strip() != "":
                            all_records.append({
                                "Date": dates[col_idx-2],
                                "Time": times[col_idx-2],
                                "Count": 1
                            })
    
    if not all_records:
        return None
        
    # 3. 建立樞紐分析表 (Pivot Table)
    df = pd.DataFrame(all_records)
    # 將時間設為 Row，日期設為 Column
    pivot_df = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
    
    return pivot_df

# Streamlit UI
uploaded_file = st.file_uploader("請上傳 Roster Excel (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        try:
            result = process_roster(uploaded_file)
            if result is not None:
                st.success("統計完成！")
                st.dataframe(result)
                
                # 下載統計報表 (CSV)
                csv = result.to_csv().encode('utf-8-sig')
                st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
            else:
                st.warning("未偵測到任何 SUP 數據，請檢查 Excel 格式。")
        except Exception as e:
            st.error(f"發生錯誤: {e}")