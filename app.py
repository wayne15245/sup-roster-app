import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="SUP 診斷工具", layout="wide")
st.title("🔍 SUP 檔案結構診斷工具")

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])

if uploaded_file:
    if st.button("查看檔案內容 (診斷模式)"):
        try:
            workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
            # 讀取第一個分頁 (你可以根據需要調整分頁索引)
            sheet = workbook.worksheets[1] 
            st.write(f"正在讀取分頁: {sheet.title}")
            
            # 讀取前 50 行
            data = []
            for row in sheet.iter_rows(min_row=1, max_row=50, values_only=True):
                data.append(row)
            
            df = pd.DataFrame(data)
            st.dataframe(df) # 這裡會把 Excel 結構直接畫出來
            st.info("💡 請觀察上方的表格：你的『時間數據』在哪一行？請記下該行號，並把這些行號更新到下方的參數中。")
            
        except Exception as e:
            st.error(f"❌ 讀取失敗: {e}")

    st.divider()
    st.subheader("設定參數 (請根據上方的表格進行修改)")
    # 這裡讓你可以直接修改，測試完後把這些數字記下來即可
    header_row = st.number_input("日期標題所在的行數 (由1開始):", value=3)
    rows_input = st.text_input("數據所在的行數 (用逗號隔開):", value="5, 11, 19, 25, 33, 39, 47, 53")
    
    if st.button("確認數據是否正確"):
        # 這裡會根據你輸入的設定，嘗試撈取看看
        st.write("測試讀取結果...")
        # (這裡可以放你之前的統計邏輯，填入變數即可)