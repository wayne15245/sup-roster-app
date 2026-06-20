import streamlit as st
import pandas as pd
import openpyxl
import re

st.set_page_config(page_title="SUP 強制統計系統", layout="wide")
st.title("📊 SUP 人數統計 (強制定義版)")

# --- 設定區：直接指定哪些行是 SUP，不需要去讀取 Pattern 分頁 ---
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
# -----------------------------------------------------------

def process_roster(uploaded_file):
    try:
        # data_only=True 是標準寫法
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # 處理第 2 至 第 7 個分頁
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            # 讀取日期 (第 2 行)
            dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            # 使用我們手動指定的行數
            for r in SUP_ROWS:
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        # 簡單過濾：包含 '-' 且長度夠長
                        if "-" in t_str and len(t_str) >= 8:
                            all_records.append({
                                "Date": str(dates[col_idx]),
                                "Time": t_str,
                                "Count": 1
                            })
                            
        if not all_records:
            st.warning("⚠️ 沒有抓取到任何資料。請檢查檔案分頁順序是否正確。")
            return None
            
        df = pd.DataFrame(all_records)
        return df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"系統錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")