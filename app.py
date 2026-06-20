import streamlit as st
import pandas as pd
import openpyxl
import datetime

st.set_page_config(page_title="SUP 統計系統", layout="wide")
st.title("📊 SUP 人數統計系統")

# 你要求的 24 個 Team (已包含 C-J, L1-L6, W-Z, N1-N7)
TEAMS = [
    "Team C", "Team D", "Team E", "Team F", "Team G", "Team H", "Team J",
    "Team L1", "Team L2", "Team L3", "Team L4", "Team L5", "Team L6",
    "Team W", "Team X", "Team Y", "Team Z",
    "Team N1", "Team N2", "Team N3", "Team N4", "Team N5", "Team N6", "Team N7"
]

# 建立無空格對照表，防呆 (防止 Excel 裡有人少打空格)
TEAM_MAP = {t.replace(" ", ""): t for t in TEAMS}

# 排除的休假與非上班代碼
IGNORE_WORDS = ["DO", "AL", "PH", "SL", "NIL", "OFF", "REST", "V", ""]

def process_roster(uploaded_file):
    workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
    all_records = []

    # 你的要求：不理會第一頁，只讀取第 2 至 7 頁 (Python 索引為 1 到 6)
    # 使用切片 [1:7] 確保只抓這 6 個分頁
    worksheets_to_process = workbook.worksheets[1:7]

    for sheet in worksheets_to_process:
        # 1. 取得表頭的日期 (第 2 行, B~H 欄)
        dates = []
        for c in range(2, 9):
            val = sheet.cell(row=2, column=c).value
            if isinstance(val, datetime.datetime):
                dates.append(val.strftime("%Y-%m-%d"))
            else:
                dates.append(str(val).strip() if val else "未知日期")

        current_team = None

        # 2. 逐行向下掃描
        for row_idx in range(1, sheet.max_row + 1):
            
            # 檢查 A 欄是否有 Team 的名稱
            cell_a = sheet.cell(row=row_idx, column=1).value
            if cell_a:
                cell_val = str(cell_a).strip().replace(" ", "")
                # 如果開頭是 Team，判定是否為我們要的 24 隊
                if cell_val.startswith("Team"):
                    if cell_val in TEAM_MAP:
                        current_team = TEAM_MAP[cell_val]
                    else:
                        current_team = None # 遇到不相干的 Team，暫停紀錄
            
            # 3. 如果我們目前正在需要的 Team 區塊內
            if current_team:
                # 檢查 I 欄 (第 9 欄) 和 J 欄 (第 10 欄) 是否有值
                sup_i = sheet.cell(row=row_idx, column=9).value
                sup_j = sheet.cell(row=row_idx, column=10).value

                is_sup = False
                if (sup_i and str(sup_i).strip() not in ["None", ""]) or \
                   (sup_j and str(sup_j).strip() not in ["None", ""]):
                    is_sup = True

                # 4. 如果是 SUP，提取他星期一至日 (B~H) 的時間
                if is_sup:
                    for i, d_col in enumerate(range(2, 9)):
                        time_val = sheet.cell(row=row_idx, column=d_col).value
                        if time_val:
                            # 移除換行符號與多餘空格
                            t_str = str(time_val).replace(" ", "").replace("\n", "").strip().upper()
                            
                            # 過濾掉休假代碼，只統計實際上班時間
                            if t_str not in IGNORE_WORDS and len(t_str) > 2:
                                # 為了顯示美觀，將大寫還原回正常時間格式 (若有需要)
                                display_time = str(time_val).replace(" ", "").replace("\n", "").strip()
                                all_records.append({
                                    "Date": dates[i],
                                    "Time": display_time,
                                    "Count": 1
                                })
    
    if not all_records: return None
        
    # 5. 轉換為矩陣格式 (上方為日期，左側為時間)
    df = pd.DataFrame(all_records)
    pivot_df = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
    pivot_df.index.name = "Time"
    
    return pivot_df

uploaded_file = st.file_uploader("請上傳 Roster Excel (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("開始計算"):
        try:
            result = process_roster(uploaded_file)
            if result is not None:
                st.success("計算完成！")
                st.dataframe(result)
                
                # 下載檔案，使用 utf-8-sig 確保 Excel 中文不亂碼
                csv = result.to_csv().encode('utf-8-sig')
                st.download_button("下載統計表 CSV", csv, "SUP_統計結果.csv", "text/csv")
            else:
                st.warning("未找到任何符合條件的 SUP 數據，請確認 Excel 格式。")
        except Exception as e:
            st.error(f"錯誤: {e}")