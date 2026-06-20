import streamlit as st
import pandas as pd
import openpyxl
import datetime

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統 (智慧解碼版)")

# 24 個 Team 的名單
TEAMS = [
    "Team C", "Team D", "Team E", "Team F", "Team G", "Team H", "Team J",
    "Team L1", "Team L2", "Team L3", "Team L4", "Team L5", "Team L6",
    "Team W", "Team X", "Team Y", "Team Z",
    "Team N1", "Team N2", "Team N3", "Team N4", "Team N5", "Team N6", "Team N7"
]

TEAM_MAP = {t.replace(" ", ""): t for t in TEAMS}
IGNORE_WORDS = ["DO", "AL", "PH", "SL", "NIL", "OFF", "REST", "V", ""]

def process_roster(uploaded_file):
    # 讀取 Excel
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    all_records = []

    # 只讀取第 2 到第 7 頁 (Index 1 到 6)
    worksheets_to_process = workbook.worksheets[1:7]

    for sheet in worksheets_to_process:
        
        # 🔑 核心修復：建立「合併儲存格」自動對應字典
        # 這樣就能解決「時間在第5行，但SUP打勾在第6行」抓不到資料的問題
        merged_dict = {}
        for merged_range in sheet.merged_cells.ranges:
            top_left_val = sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value
            for r in range(merged_range.min_row, merged_range.max_row + 1):
                for c in range(merged_range.min_col, merged_range.max_col + 1):
                    merged_dict[(r, c)] = top_left_val

        # 輔助讀取函式：自動判斷是不是合併儲存格
        def get_cell_val(r, c):
            if (r, c) in merged_dict:
                return merged_dict[(r, c)]
            return sheet.cell(row=r, column=c).value

        # 1. 取得表頭的日期 (第 2 行, B~H 欄)
        dates = []
        for c in range(2, 9):
            val = get_cell_val(2, c)
            if isinstance(val, datetime.datetime):
                dates.append(val.strftime("%Y-%m-%d"))
            else:
                dates.append(str(val).strip() if val else "未知日期")

        current_team = None

        # 2. 逐行向下掃描
        for row_idx in range(1, sheet.max_row + 1):
            
            # 檢查 A 欄判斷目前的 Team
            cell_a = get_cell_val(row_idx, 1)
            if cell_a:
                cell_val = str(cell_a).strip().replace(" ", "")
                if cell_val.startswith("Team"):
                    if cell_val in TEAM_MAP:
                        current_team = TEAM_MAP[cell_val]
                    else:
                        current_team = None 
            
            # 3. 如果在指定的 Team 區塊內
            if current_team:
                # 檢查 I(9) 欄和 J(10) 欄是否有 SUP 標記
                sup_i = get_cell_val(row_idx, 9)
                sup_j = get_cell_val(row_idx, 10)

                is_sup = False
                if (sup_i and str(sup_i).strip() not in ["None", ""]) or \
                   (sup_j and str(sup_j).strip() not in ["None", ""]):
                    is_sup = True

                # 4. 抓取該 SUP 的一週時間
                if is_sup:
                    for i, d_col in enumerate(range(2, 9)):
                        time_val = get_cell_val(row_idx, d_col)
                        if time_val:
                            # 統一轉大寫並去除空格，用來比對休假字眼
                            t_str = str(time_val).replace(" ", "").replace("\n", "").strip().upper()
                            
                            if t_str not in IGNORE_WORDS and len(t_str) > 2:
                                # 實際顯示的文字 (保留原本格式)
                                display_time = str(time_val).replace(" ", "").replace("\n", "").strip()
                                all_records.append({
                                    "Date": dates[i],
                                    "Time": display_time,
                                    "Count": 1
                                })
    
    # 🛡️ 防呆機制：如果真的什麼都沒抓到，回傳空的表格，避免程式崩潰報錯
    if not all_records: 
        return pd.DataFrame()
        
    # 5. 製作統計矩陣
    df = pd.DataFrame(all_records)
    pivot_df = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
    pivot_df.index.name = "Time"
    
    return pivot_df

uploaded_file = st.file_uploader("請上傳 Roster Excel (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        try:
            result = process_roster(uploaded_file)
            
            # 檢查是否有資料
            if not result.empty:
                st.success("🎉 計算完成！")
                st.dataframe(result)
                
                # 下載 CSV
                csv = result.to_csv().encode('utf-8-sig')
                st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
            else:
                st.warning("⚠️ 檔案讀取成功，但未找到任何符合條件的 SUP 數據。請確認：\n1. Excel 第 2~7 頁是否有資料\n2. I 或 J 欄是否有正確填寫標記\n3. 時間格子內是否有內容。")
        except Exception as e:
            st.error(f"程式發生例外錯誤: {e}")