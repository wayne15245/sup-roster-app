import streamlit as st
import pandas as pd
import openpyxl
import re

st.set_page_config(page_title="SUP 智慧掃描系統", layout="wide")
st.title("📊 SUP 人數統計 (智慧掃描版)")

# 時間格式範本：符合 0000-0000 格式
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        # 掃描全部工作表 (或者您可以在這裡指定範圍，例如 workbook.worksheets[1:7])
        target_sheets = workbook.worksheets 
        all_records = []
        
        for sheet in target_sheets:
            # 排除掉 Pattern 分頁，只抓有日期欄位的頁面
            if sheet.title == "Pattern":
                continue
            
            # 取得該頁的日期 (假設日期在第 2 行，欄位 B 到 H)
            dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            # --- 智慧掃描：不再指定行數，直接遍歷第 1 到 100 行 ---
            for r in range(1, 100):
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        # 只要內容符合 0000-0000 格式，就納入統計
                        if TIME_PATTERN.match(t_str):
                            # 將該日期加入記錄
                            date_val = str(dates[col_idx]) if col_idx < len(dates) else "未知"
                            all_records.append({"Date": date_val, "Time": t_str, "Count": 1})
        
        if not all_records:
            st.warning("⚠️ 掃描了 100 行還是沒找到資料。請確認資料是否確實存在於 B 到 H 欄 (第 2-8 欄) 中？")
            return None
            
        df = pd.DataFrame(all_records)
        return df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"❌ 系統錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("成功抓取資料！")
        st.dataframe(result)
        csv = result.to_csv().encode('utf-8-sig')
        st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")