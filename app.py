import streamlit as st
import pandas as pd
import openpyxl
import re

st.set_page_config(page_title="SUP 除錯系統", layout="wide")
st.title("📊 SUP 除錯與統計")

# 設定區
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]
HEADER_ROW = 3 # 依照您之前的調整，設為第3行
TIME_PATTERN = re.compile(r'^\d{4}-\d{4}$')

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])

# 將判斷邏輯拆開，增加除錯訊息
if uploaded_file:
    st.success(f"✅ 已偵測到檔案: {uploaded_file.name}")
    
    if st.button("開始計算"):
        st.write("⏳ 開始計算中...")
        try:
            workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
            st.write(f"📂 成功讀取 Excel，共有 {len(workbook.worksheets)} 個分頁")
            
            # 檢查分頁數量
            if len(workbook.worksheets) < 2:
                st.error("❌ 檔案中找不到足夠的分頁 (需要至少 2 個以上)")
            else:
                target_sheets = workbook.worksheets[1:7]
                st.write(f"⚙️ 處理分頁: {[s.title for s in target_sheets]}")
                
                all_records = []
                for sheet in target_sheets:
                    # 讀取標頭 (B3 到 H3)
                    headers = [sheet.cell(row=HEADER_ROW, column=c).value for c in range(2, 9)]
                    
                    for r in SUP_ROWS:
                        for col_idx in range(7):
                            header_val = str(headers[col_idx]) if headers[col_idx] else "未知"
                            val = sheet.cell(row=r, column=col_idx + 2).value
                            
                            if val:
                                t_str = str(val).strip()
                                if TIME_PATTERN.match(t_str):
                                    all_records.append({
                                        "Date/Day": header_val,
                                        "Time": t_str,
                                        "Count": 1
                                    })
                
                if all_records:
                    st.write(f"✅ 成功抓取到 {len(all_records)} 筆資料")
                    df = pd.DataFrame(all_records)
                    result = df.pivot_table(index="Time", columns="Date/Day", values="Count", aggfunc="sum", fill_value=0)
                    
                    st.dataframe(result)
                    csv = result.to_csv().encode('utf-8-sig')
                    st.download_button("下載 CSV", csv, "SUP_統計結果.csv", "text/csv")
                else:
                    st.warning("⚠️ 掃描完成，但找不到任何符合 0000-0000 格式的時間資料。")
                    
        except Exception as e:
            st.error(f"❌ 程式發生錯誤: {e}")
            st.exception(e) # 顯示完整的錯誤堆疊，幫你找到問題根源
else:
    st.info("請上傳檔案以開始")