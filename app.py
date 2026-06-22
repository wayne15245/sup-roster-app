import streamlit as st
import pandas as pd
import openpyxl
import re
from openpyxl.styles import PatternFill, Font

# ===== UI =====
st.set_page_config(page_title="SUP Manpower Calculator", layout="wide")
st.title("📊 SUP 人手統計系統（Final Clean Report）")

uploaded_file = st.file_uploader("請上傳 Roster Excel", type=["xlsm", "xlsx"])

# ===== 設定 =====
TARGET_SHEETS = [1, 2, 3, 4, 5, 6]
SUMMARY_COLS = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]

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
    if val == "" or val == "0":
        return False
    if val.replace(" ", "") == "":
        return False
    return True

# ===== 分類 =====
def get_shift_group(time_str):
    try:
        start = int(time_str.split("-")[0])
        if 700 <= start <= 959:
            return "0730-1930"
        elif 1000 <= start <= 1459:
            return "1200-0100"
        elif 1500 <= start <= 2059:
            return "1530-0530"
        else:
            return "2130-0930"
    except:
        return None

# ===== 顏色 =====
def get_excel_fill(group):
    if group == "0730-1930":
        return PatternFill(start_color="FFFF00", fill_type="solid")
    elif group == "1200-0100":
        return PatternFill(start_color="FFC000", fill_type="solid")
    elif group == "1530-0530":
        return PatternFill(start_color="92D050", fill_type="solid")
    elif group == "2130-0930":
        return PatternFill(start_color="00B0F0", fill_type="solid")
    elif group == "Daily_Total":
        return PatternFill(start_color="FFC000", fill_type="solid")
    return None


if uploaded_file:

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    final_result = {}

    # ===== 第2–7頁 =====
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
                    if is_valid_sup(sheet.cell(row=sup_row, column=10).value):
                        count += 1

                final_result[time_value][col_idx-2] += count

    # ===== 第8頁 =====
    try:
        sheet = wb.worksheets[7]
        for row in range(5, 18):
            if not is_valid_sup(sheet.cell(row=row, column=10).value):
                continue
            for col_idx in range(2, 9):
                time_value = extract_time(sheet.cell(row=row, column=col_idx).value)
                if time_value:
                    final_result.setdefault(time_value, [0]*7)
                    final_result[time_value][col_idx-2] += 1
    except:
        st.warning("⚠️ 第8頁失敗")

    # ===== 第9頁 =====
    try:
        sheet = wb.worksheets[8]
        for row in range(4, 35):
            if not is_valid_sup(sheet.cell(row=row, column=10).value):
                continue
            for col_idx in range(2, 9):
                time_value = extract_time(sheet.cell(row=row, column=col_idx).value)
                if time_value:
                    final_result.setdefault(time_value, [0]*7)
                    final_result[time_value][col_idx-2] += 1
    except:
        st.warning("⚠️ 第9頁失敗")

    # ===== DF =====
    if final_result:

        df = pd.DataFrame.from_dict(final_result, orient="index", columns=SUMMARY_COLS)
        df = df.reset_index()
        df.columns = ["Time"] + SUMMARY_COLS
        df = df.sort_values(by="Time")
        df["Total"] = df[SUMMARY_COLS].sum(axis=1)

        # ===== Summary =====
        summary_data = {
            "0730-1930":[0]*7,
            "1200-0100":[0]*7,
            "1530-0530":[0]*7,
            "2130-0930":[0]*7
        }

        for t, vals in final_result.items():
            group = get_shift_group(t)
            if group:
                for i in range(7):
                    summary_data[group][i] += vals[i]

        summary_rows = []
        for g, vals in summary_data.items():
            row = {"Time": g}
            for i, col in enumerate(SUMMARY_COLS):
                row[col] = vals[i]
            row["Total"] = sum(vals)
            summary_rows.append(row)

        df = pd.concat([df, pd.DataFrame(summary_rows)], ignore_index=True)

        # ===== Daily Total =====
        daily_total = {"Time": "Daily_Total"}
        for i, col in enumerate(SUMMARY_COLS):
            daily_total[col] = sum(v[i] for v in final_result.values())
        daily_total["Total"] = sum(daily_total[col] for col in SUMMARY_COLS)
        df.loc[len(df)] = daily_total

        # ===== 最少人 =====
        filtered = df[~df["Time"].isin(["0730-1930","1200-0100","1530-0530","2130-0930","Daily_Total"])]
        min_row = filtered.loc[filtered["Total"].idxmin()]
        st.write(f"📉 最少人：{min_row['Time']} ({min_row['Total']}人)")

        st.dataframe(df, use_container_width=True)

        # ===== Excel輸出 =====
        output = "SUP_Result.xlsx"

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
            ws = writer.active

            ws.freeze_panes = "A2"

            bold = Font(bold=True)
            summary_labels = ["0730-1930","1200-0100","1530-0530","2130-0930","Daily_Total"]

            for r in range(2, len(df)+2):
                val = ws.cell(row=r, column=1).value

                # ✅ 只底部summary加色+粗體
                if val in summary_labels and r > len(df) - 5:
                    fill = get_excel_fill(val)
                    for c in range(1, 10):
                        cell = ws.cell(row=r, column=c)
                        cell.font = bold
                        if fill:
                            cell.fill = fill

            # ✅ header加粗
            for c in range(1, 10):
                ws.cell(row=1, column=c).font = bold

            # ✅ 自動欄寬
            for col in ws.columns:
                length = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = length + 2

        with open(output, "rb") as f:
            st.download_button("⬇️ Download Excel", f, file_name=output)

    else:
        st.error("❌ 無數據")

else:
    st.info("請上載Excel")