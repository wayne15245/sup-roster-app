import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="SUP 強制診斷系統", layout="wide")
st.title("📊 SUP 統計系統 (全診斷模式)")

# 設定區
SUP_ROWS = [5, 11, 19, 25, 33, 39, 47, 53]

def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # --- 診斷步驟 1：印出所有分頁名稱 ---
        all_sheets = workbook.sheetnames
        st.write(f"📂 系統偵測到共有 {len(all_sheets)} 個分頁，名稱如下：")
        st.write(all_sheets)
        
        # 我們假設資料在第 2 到第 7 個分頁 (index 1 到 6)
        # 如果您的資料分頁其實不在這裡，請看上面的列表名稱調整
        target_sheets = workbook.worksheets[1:7]
        st.write(f"🔧 系統目前鎖定的分頁範圍是：{[s.title for s in target_sheets]}")
        
        all_records = []
        for sheet in target_sheets:
            dates = [sheet.cell(row=2, column=c).value for c in range(2, 9)]
            
            # 隨機檢查一下有沒有抓到資料，如果都讀不到，印出提示
            found_any = False
            for r in SUP_ROWS:
                for col_idx, col in enumerate(range(2, 9)):
                    val = sheet.cell(row=r, column=col).value
                    if val:
                        t_str = str(val).strip()
                        if "-" in t_str and len(t_str) >= 8:
                            all_records.append({"Date": str(dates[col_idx]), "Time": t_str, "Count": 1})
                            found_any = True
            if not found_any:
                st.write(f"⚠️ 注意：分頁 '{sheet.title}' 在指定行數 ({SUP_ROWS}) 沒讀到任何時間格式資料。")
        
        if not all_records:
            st.warning("⚠️ 統計結果依然為空。請參考上面的分頁列表，確認資料是否真的在這些分頁裡。")
            return None
            
        df = pd.DataFrame(all_records)
        return df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)

    except Exception as e:
        st.error(f"❌ 程式發生錯誤: {e}")
        return None

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])
if uploaded_file and st.button("開始計算"):
    result = process_roster(uploaded_file)
    if result is not None:
        st.success("統計完成！")
        st.dataframe(result)