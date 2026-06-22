import streamlit as st
import pandas as pd
import openpyxl
import re
from datetime import datetime

st.set_page_config(page_title="SUP 嚴格統計系統", layout="wide")
st.title("📊 SUP 精準統計系統")

# 設定區：嚴格鎖定目標
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        # 只取第 2 至 第 7 個分頁
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            # 讀取日期 (第2行, 第2-8欄)
            header_dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            # 數據清洗：確保日期欄位是有效日期或格式
            valid_dates = []
            for d in header_dates:
                if isinstance(d, datetime):
                    valid_dates.append(d.strftime('%Y-%m-%d'))
                else:
                    valid_dates.append(str(d))

            for r in SUP_ROWS:
                # 只遍歷 B 到 H 欄 (2 到 8)
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        if TIME_PATTERN.match(t_str):
                            # 只有在日期是有效字串時才加入
                            all_records.append({
                                "Date": valid_dates[col_idx], 
                                "Time": t_str, 
                                "Count": 1
                            })
        
        if not all_records:
            return None
            
        df = pd.DataFrame(all_records)
        # 建立 Pivot Table
        result = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        
        # 強制只保留符合 YYYY-MM-DD 格式的日期欄位 (過濾掉 'Fri', 'Mon' 等雜訊)
        date_cols = [c for c in result.columns if re.match(r'\d{4}-\d{2}-\d{2}', str(c))]
        return result[date_cols]

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("統計完成！")
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_精準統計結果.csv", "text/csv")
    else:
        st.warning("⚠️ 未偵測到資料。")