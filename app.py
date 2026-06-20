import streamlit as st
import pandas as pd
import openpyxl
import re

st.set_page_config(page_title="SUP 最終統計系統", layout="wide")
st.title("📊 SUP 人數統計 (最終修正版)")

# 目標行
TARGET_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # 1. 建立 SUP 狀態表 (來自 Pattern 分頁)
        if 'Pattern' not in workbook.sheetnames:
            st.error("找不到 'Pattern' 分頁！")
            return None
            
        pattern_ws = workbook['Pattern']
        sup_rows = []
        for r in TARGET_ROWS:
            j_val = pattern_ws.cell(row=r, column=10).value # J 欄是第 10 欄
            # 只要 J 欄不是空的，就判定這行為 SUP
            if j_val is not None and str(j_val).strip() != "":
                sup_rows.append(r)
        
        st.write(f"✅ 已識別出 Pattern 中有 SUP 標記的行數: {sup_rows}")

        # 2. 處理指定分頁 (第 2 至 第 7 頁)
        # index 1 是第 2 頁，index 7 是第 8 頁 (不含)，所以切片 [1:7]
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            # 讀取日期
            dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            # 針對目標行讀取
            for r in sup_rows:
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        # 寬鬆過濾：只要字串包含 "-" 且長度大於 8，就視為時間
                        if "-" in t_str and len(t_str) >= 8:
                            all_records.append({
                                "Date": str(dates[col_idx]),
                                "Time": t_str,
                                "Count": 1
                            })
                            
        if not all_records:
            st.warning("⚠️ 統計結果為空。這通常是因為 Excel 中的日期或時間格式無法被識別。")
            return None
            
        df = pd.DataFrame(all_records)
        pivot = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        return pivot

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