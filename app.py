import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="SUP 強制偵錯版", layout="wide")
st.title("📊 SUP 人數統計 (偵錯模式)")

# 定義行號 (Row Numbers)
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
# 10 對應 J 欄
J_COL = 10 

def process_roster(uploaded_file):
    try:
        # data_only=True 讀取數值
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # 顯示所有工作表名稱，確認是否有抓對
        st.write(f"📂 檔案中的工作表: {workbook.sheetnames}")
        
        # 處理第 2 至 第 7 個分頁 (Sheet2 到 Sheet7)
        # 注意: 如果你的檔案 Sheet 順序不同，這裡可能會抓錯
        target_sheets = workbook.worksheets[1:7]
        all_records = []
        
        for sheet in target_sheets:
            st.write(f"--- 正在檢查分頁: {sheet.title} ---")
            dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            for r in SUP_ROWS:
                # 偵錯：顯示它在目標位置讀到了什麼
                cell_val = sheet.cell(row=r, column=J_COL).value
                st.write(f"檢查第 {r} 行, 第 {J_COL} 欄 (J{r}): 內容為 '{cell_val}'")
                
                # 遍歷 B 欄 (2) 到 H 欄 (8)
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        if "-" in t_str and len(t_str) >= 8:
                            all_records.append({"Date": str(dates[col_idx]), "Time": t_str, "Count": 1})
        
        if not all_records:
            st.warning("⚠️ 統計為空。請參考上面的『檢查紀錄』，如果讀到的都是 None，代表 Python 沒讀到 Excel 裡的格子。")
            return None
            
        df = pd.DataFrame(all_records)
        return df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"❌ 錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.dataframe(result)