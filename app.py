import streamlit as st
import pandas as pd
import openpyxl
import datetime

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統 (精準抓取版)")

# 包含空格的 24 個 Team
TEAMS = [
    "Team C", "Team D", "Team E", "Team F", "Team G", "Team H",
    "Team L1", "Team L2", "Team L3", "Team L4", "Team L5", "Team L6",
    "Team W", "Team X", "Team Y", "Team Z",
    "Team N1", "Team N2", "Team N3", "Team N4", "Team N5", "Team N6", "Team N7",
    "Team J"
]

# 建立一個無空格版本的對照表，防止 Excel 裡有人打錯字 (例如把 Team C 打成 TeamC)
TEAM_MAP = {t.replace(" ", ""): t for t in TEAMS}

def process_roster(uploaded_file):
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    all_records = []

    for sheet in workbook.worksheets:
        # 1. 取得表頭的日期 (第 2 行, B~H 欄)
        dates = []
        for c in range(2, 9):
            val = sheet.cell(row=2, column=c).value
            # 如果 Excel 將其視為時間物件，強制轉為純文字日期 (YYYY-MM-DD)
            if isinstance(val, datetime.datetime):
                dates.append(val.strftime("%Y-%m-%d"))
            else:
                dates.append(str(val).strip())

        current_team = None

        # 2. 逐行向下掃描整個分頁
        for row_idx in range(1, sheet.max_row + 1):
            
            # 檢查 A, B 欄是否有 Team 的名稱
            for col_idx in [1, 2]:
                cell_val = str(sheet.cell(row=row_idx, column=col_idx).value).strip().replace(" ", "")
                if cell_val.startswith("Team"):
                    # 如果是我們要的 24 個 Team 之一，鎖定它
                    if cell_val in TEAM_MAP:
                        current_team = TEAM_MAP[cell_val]
                    else:
                        current_team = None # 遇到不要的 Team (例如 Team A)，直接忽略其後的資料
                    break

            # 3. 如果目前正在我們需要的 Team 區塊內
            if current_team:
                # 檢查 I 欄 (第 9 欄) 和 J 欄 (第 10 欄) 是否有字
                sup_i = sheet.cell(row=row_idx, column=9).value
                sup_j = sheet.cell(row=row_idx, column=10).value

                is_sup = False
                if (sup_i and str(sup_i).strip() not in ["None", ""]) or \
                   (sup_j and str(sup_j).strip() not in ["None", ""]):
                    is_sup = True

                # 如果確認這行是 SUP，就去抓他 B~H 欄 (第 2~8 欄) 的上班時間
                if is_sup:
                    for i, d_col in enumerate(range(2, 9)):
                        time_val = sheet.cell(row=row_idx, column=d_col).value
                        if time_val:
                            t_str = str(time_val).replace(" ", "").strip()
                            # 排除掉休假或空白格
                            if t_str not in ["None", "", "OFF", "REST", "AL", "SL"]:
                                all_records.append({
                                    "Date": dates[i],
                                    "Time": t_str,
                                    "Count": 1
                                })
    
    if not all_records: return None
        
    # 4. 統計與排版 (上方為日期，左側為時間)
    df = pd.DataFrame(all_records)
    pivot_df = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
    
    # 確保 CSV 左上角第一個格子顯示為 "Time"
    pivot_df.index.name = "Time"
    return pivot_df

uploaded_file = st.file_uploader("請上傳 Roster Excel (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        try:
            result = process_roster(uploaded_file)
            if result is not None:
                st.dataframe(result)
                csv = result.to_csv().encode('utf-8-sig')
                st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
            else:
                st.warning("未找到任何符合條件的 SUP 數據，請確認 Excel 格式。")
        except Exception as e:
            st.error(f"錯誤: {e}")