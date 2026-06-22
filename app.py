import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="SUP Manpower Calculator", layout="wide")
st.title("📊 SUP 人手統計系統")

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])

# ====== 固定設定 ======
TARGET_SHEETS = [1, 2, 3, 4, 5, 6]  # 第2至7頁（0-based）
TIME_BLOCKS = {
    4: [5, 11],
    18: [19, 25],
    32: [33, 39],
    46: [47, 53]
}

def normalize_time(val):
    if val is None:
        return None
    return str(val).strip()

# ====== 主邏輯 ======
if uploaded_file:

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)

    final_result = {}

    for sheet_index in TARGET_SHEETS:
        sheet = wb.worksheets[sheet_index]

        # ===== 日期 (Mon-Sun)
        dates = []
        for col in range(2, 9):  # B:H
            value = sheet.cell(row=3, column=col).value
            dates.append(value)

        # ===== 每個 block 處理
        for time_row, sup_rows in TIME_BLOCKS.items():

            for col_idx in range(2, 9):  # B:H

                time_value = normalize_time(
                    sheet.cell(row=time_row, column=col_idx).value
                )

                if not time_value:
                    continue

                # 初始化
                if time_value not in final_result:
                    final_result[time_value] = [0] * 7

                count = 0

                for sup_row in sup_rows:
                    cell_value = sheet.cell(row=sup_row, column=10).value  # J column
                    if cell_value not in [None, ""]:
                        count += 1

                final_result[time_value][col_idx - 2] += count

    # ===== 轉 DataFrame =====
    if final_result:

        df = pd.DataFrame.from_dict(
            final_result,
            orient="index",
            columns=["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        )

        df = df.reset_index()
        df.columns = ["Time", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

        # 排序
        df = df.sort_values(by="Time")

        st.success("✅ 計算完成")

        st.dataframe(df, use_container_width=True)

        # ===== 下載功能 =====
        output_file = "SUP_Result.xlsx"
        df.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="⬇️ 下載 Excel",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.error("❌ 無法解析資料，請確認Excel格式")

else:
    st.info("請上傳Excel開始")