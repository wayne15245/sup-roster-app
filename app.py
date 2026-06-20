import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="SUP 偵錯與統計系統", layout="wide")
st.title("📊 SUP 統計系統 (強制除錯版)")

TARGET_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]

def process_roster(uploaded_file):
    try:
        # 使用 data_only=True 讀取公式計算後的結果
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # 1. 檢查 Pattern 分頁
        if 'Pattern' not in workbook.sheetnames:
            st.error("找不到 'Pattern' 分頁！")
            return None
        
        pattern_ws = workbook['Pattern']
        st.subheader("🔍 正在偵錯 J 欄數據...")
        
        sup_rows = []
        for r in TARGET_ROWS:
            # 直接讀取 J 欄 (第 10 欄)
            cell_val = pattern_ws.cell(row=r, column=10).value
            st.write(f"第 {r} 行, J 欄讀取值: **{cell_val}** (類型: {type(cell_val)})")
            
            # 如果讀到 1 或其他數字，就加入清單
            if cell_val is not None and str(cell_val).strip() not in ["0", "None", ""]:
                sup_rows.append(r)
        
        st.write(f"✅ 系統最終鎖定的 SUP 行數: {sup_rows}")
        
        if not sup_rows:
            st.warning("⚠️ 雖然掃描了 J 欄，但沒找到任何有效標記 (讀取到的可能都是 None 或 0)。")
            return None

        # 2. 處理指定分頁 (第 2 至 第 7 頁)
        # 注意: sheets[1:7] 取出的是 index 1, 2, 3, 4, 5, 6 (即第 2 到第 7 個頁面)
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            st.write(f"正在處理: {sheet.title}")
            dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            for r in sup_rows:
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        # 簡單的格式過濾：包含 '-' 且長度大於 8
                        if "-" in t_str and len(t_str) >= 8:
                            all_records.append({"Date": str(dates[col_idx]), "Time": t_str, "Count": 1})
        
        if not all_records:
            st.warning("⚠️ 未讀取到資料。")
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