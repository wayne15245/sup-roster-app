import streamlit as st
import pandas as pd
import openpyxl
import re

st.set_page_config(page_title="SUP Manpower Calculator", layout="wide")
st.title("📊 SUP 人手統計系統（智能版）")

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])

# ====== 設定 ======
TARGET_SHEETS = [1, 2, 3, 4, 5, 6]  # 第2–7頁 (0-based index)

TIME_BLOCKS = {
    4: [5, 11],
    18: [19, 25],
    32: [33, 39],
    46: [47, 53]
}

# ====== 智能時間抽取 ======
TIME_EXTRACT_PATTERN = re.compile(r'(\d{4}-\d{4})')

def extract_time(val):
    if not val:
        return None

    val = str(val)
    match = TIME_EXTRACT_PATTERN.search(val)

    if match:
        return match.group(1)

    return None


# ====== 主程式 ======
if uploaded_file:

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    final_result = {}

    # =========================
    # ✅ 第2–7頁（Block Structure）
    # =========================
    for sheet_index in TARGET_SHEETS:
        sheet = wb.worksheets[sheet_index]

        for time_row, sup_rows in TIME_BLOCKS.items():

            for col_idx in range(2, 9):  # B:H

                raw_value = sheet.cell(row=time_row, column=col_idx).value
                time_value = extract_time(raw_value)

                if not time_value:
                    continue

                if time_value not in final_result:
                    final_result[time_value] = [0] * 7

                count = 0

                for sup_row in sup_rows:
                    cell_value = sheet.cell(row=sup_row, column=10).value  # J欄

                    if cell_value not in [None, ""]:
                        count += 1

                final_result[time_value][col_idx - 2] += count


    # =========================
    # ✅ 第8頁（Row Structure + OT支援🔥）
    # =========================
    try:
        sheet = wb.worksheets[7]

        for row in range(5, 18):  # J5:J17

            sup_value = sheet.cell(row=row, column=10).value

            if sup_value in [None, ""]:
                continue

            for col_idx in range(2, 9):  # B:H

                raw_value = sheet.cell(row=row, column=col_idx).value
                time_value = extract_time(raw_value)

                if not time_value:
                    continue

                if time_value not in final_result:
                    final_result[time_value] = [0] * 7

                final_result[time_value][col_idx - 2] += 1

    except:
        st.warning("⚠️ 第8頁讀取失敗（可以忽略）")


    # =========================
    # ✅ 輸出結果
    # =========================
    if final_result:

        df = pd.DataFrame.from_dict(
            final_result,
            orient="index",
            columns=["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        )

        df = df.reset_index()
        df.columns = ["Time", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

        # 排序（按時間字串）
        df = df.sort_values(by="Time")

        # 加總
        df["Total"] = df.iloc[:, 1:].sum(axis=1)

        st.success("✅ 計算完成")
        st.dataframe(df, use_container_width=True)

        # 📊 圖表
        st.subheader("📈 人手分佈")
        st.bar_chart(df.set_index("Time")[["Total"]])

        # 下載
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
        st.error("❌ 無數據，請檢查Excel格式")

else:
    st.info("請上傳Excel開始")
``