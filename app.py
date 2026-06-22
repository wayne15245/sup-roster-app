import streamlit as st
import pandas as pd
import openpyxl
import re

# ===== UI =====
st.set_page_config(page_title="SUP Manpower Calculator", layout="wide")
st.title("📊 SUP 人手統計系統（Final Pro Fix）")

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])

# ===== 設定 =====
TARGET_SHEETS = [1, 2, 3, 4, 5, 6]

TIME_BLOCKS = {
    4: [5, 11],
    18: [19, 25],
    32: [33, 39],
    46: [47, 53]
}

# ===== 時間抽取 =====
TIME_PATTERN = re.compile(r'(\d{4}-\d{4})')

def extract_time(val):
    if not val:
        return None
    match = TIME_PATTERN.search(str(val))
    return match.group(1) if match else None


# ===== 防假SUP =====
def is_valid_sup(val):
    if val is None:
        return False

    val = str(val).strip()

    if val == "":
        return False

    if val == "0":
        return False

    if val.replace(" ", "") == "":
        return False

    return True


# ===== Summary mapping =====
SUMMARY_COLS = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]

SUMMARY_MAP = {
    "0730-1930": ["0730-1630", "0730-1930", "0830-1630", "0900-1800"],
    "1200-0100": ["1200-2100", "1300-0100"],
    "1530-0530": ["1500-0100", "1530-0330", "1630-0330", "1730-0330", "1830-0330"],
    "2130-0930": ["2030-0530", "2130-0530", "2230-0730", "2300-0800", "2330-0830"]
}


# ===== 主流程 =====
if uploaded_file:

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    final_result = {}

    # =========================
    # ✅ 第2–7頁（Block）
    # =========================
    for sheet_index in TARGET_SHEETS:
        sheet = wb.worksheets[sheet_index]

        for time_row, sup_rows in TIME_BLOCKS.items():

            for col_idx in range(2, 9):

                time_value = extract_time(sheet.cell(row=time_row, column=col_idx).value)

                if not time_value:
                    continue

                if time_value not in final_result:
                    final_result[time_value] = [0]*7

                count = 0

                for sup_row in sup_rows:
                    cell_value = sheet.cell(row=sup_row, column=10).value

                    if is_valid_sup(cell_value):
                        count += 1

                final_result[time_value][col_idx-2] += count


    # =========================
    # ✅ 第8頁（J5–J17）
    # =========================
    try:
        sheet = wb.worksheets[7]

        for row in range(5, 18):

            if not is_valid_sup(sheet.cell(row=row, column=10).value):
                continue

            for col_idx in range(2, 9):

                time_value = extract_time(sheet.cell(row=row, column=col_idx).value)

                if not time_value:
                    continue

                if time_value not in final_result:
                    final_result[time_value] = [0]*7

                final_result[time_value][col_idx-2] += 1

    except:
        st.warning("⚠️ 第8頁讀取失敗")


    # =========================
    # ✅ 第9頁（J4–J34）
    # =========================
    try:
        sheet = wb.worksheets[8]

        for row in range(4, 35):

            if not is_valid_sup(sheet.cell(row=row, column=10).value):
                continue

            for col_idx in range(2, 9):

                time_value = extract_time(sheet.cell(row=row, column=col_idx).value)

                if not time_value:
                    continue

                if time_value not in final_result:
                    final_result[time_value] = [0]*7

                final_result[time_value][col_idx-2] += 1

    except:
        st.warning("⚠️ 第9頁讀取失敗")


    # =========================
    # ✅ Debug（測完可刪）
    # =========================
    if final_result:
        st.json(final_result)


    # =========================
    # ✅ DataFrame
    # =========================
    if final_result:

        df = pd.DataFrame.from_dict(
            final_result,
            orient="index",
            columns=SUMMARY_COLS
        )

        df = df.reset_index()
        df.columns = ["Time"] + SUMMARY_COLS

        df = df.sort_values(by="Time")

        # ✅ 原始 Total
        df["Total"] = df[SUMMARY_COLS].sum(axis=1)


        # =========================
        # ✅ Summary rows（已修正）
        # =========================
        summary_rows = []

        for summary_time, time_list in SUMMARY_MAP.items():

            sub_df = df[df["Time"].isin(time_list)]

            if sub_df.empty:
                continue

            sums = sub_df[SUMMARY_COLS].sum()

            row = {"Time": summary_time}

            for col in SUMMARY_COLS:
                row[col] = sums[col]

            # ✅ 修正：只加數字欄
            row["Total"] = sum(row[col] for col in SUMMARY_COLS)

            summary_rows.append(row)

        if summary_rows:
            df = pd.concat([df, pd.DataFrame(summary_rows)], ignore_index=True)


        # =========================
        # ✅ FINAL TOTAL（已修正）
        # =========================
        total_row = {"Time": "TOTAL"}

        for col in SUMMARY_COLS:
            total_row[col] = df[col].sum()

        total_row["Total"] = sum(total_row[col] for col in SUMMARY_COLS)

        df.loc[len(df)] = total_row


        # =========================
        # ✅ UI
        # =========================
        st.success("✅ 計算完成")
        st.dataframe(df, use_container_width=True)

        st.subheader("📈 Total人手")
        st.bar_chart(df.set_index("Time")[["Total"]])


        # =========================
        # ✅ Excel下載
        # =========================
        output_file = "SUP_Result.xlsx"
        df.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                "⬇️ 下載 Excel",
                f,
                file_name=output_file
            )

    else:
        st.error("❌ 無數據")

else:
    st.info("請上載Excel")