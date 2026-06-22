import streamlit as st
import pandas as pd
import openpyxl
import re
from datetime import datetime, date

st.set_page_config(page_title="SUP 終極穩定版", layout="wide")
st.title("📊 SUP 人數統計 (精準鎖定版)")

# 指定要讀取的行列
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

def get_valid_date(val):
    """嚴格判斷：只有真正的日期才回傳格式化字串，否則回傳 None"""
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m-%d')
    return None

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        target_sheets = workbook.worksheets[1:7] # 只抓第2-7分頁
        all_records = []
        
        for sheet in target_sheets:
            # 【關鍵步驟】：明確從 B2 到 H2 讀取日期
            # 只有這一行讀到的日期才會被保留，其他欄位一律視為雜訊
            header_dates = []
            for col in range(2, 9):
                val = sheet.cell(row=2, column=col).value
                header_dates.append(get_valid_date(val))
            
            # 對應這 7 欄進行讀取
            for r in SUP_ROWS:
                for col_idx in range(7): # 0到6，對應 B到H
                    # 如果 header_dates[col_idx] 是 None，說明這欄不是有效日期，直接跳過
                    if header_dates[col_idx] is None:
                        continue
                        
                    # 讀取目標儲存格
                    val = sheet.cell(row=r, column=col_idx + 2).value
                    if val:
                        t_str = str(val).strip()
                        if TIME_PATTERN.match(t_str):
                            all_records.append({
                                "Date": header_dates[col_idx],
                                "Time": t_str,
                                "Count": 1
                            })
        
        if not all_records:
            return None
            
        df = pd.DataFrame(all_records)
        # 因為前面已經過濾了 None，這裡的 Date 絕對只會是 YYYY-MM-DD
        result = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        return result

    except Exception as e:
        st.error(f"❌ 系統錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("統計完成！欄位已鎖定。")
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_精準統計.csv", "text/csv")
    else:
        st.warning(⚠️ 未偵測到資料。請確認分頁第 2 行的 B2-H2 是否已設定為 Excel 日期格式。")