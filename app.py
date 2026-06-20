def process_roster(uploaded_file):
    try:
        workbook = openpyxl.load_workbook(uploaded_file, data_only=True)
        all_records = []
        
        # 為了除錯，顯示目前讀到的頁數
        st.write(f"正在讀取檔案，共有 {len(workbook.worksheets)} 個分頁。")
        
        # 改為讀取所有分頁，或者是您指定的區間
        for sheet in workbook.worksheets: 
            
            # 1. 取得日期列 (第 2 行)
            dates = []
            for c in range(2, 9): 
                val = get_merged_cell_value(sheet, 2, c)
                dates.append(val.strftime("%Y-%m-%d") if isinstance(val, datetime.datetime) else "未知日期")

            # 2. 針對目標行進行精準掃描
            for r in TARGET_ROWS:
                # 確認 Team (A 欄)
                team_cell = str(get_merged_cell_value(sheet, r, 1)).strip()
                
                if team_cell in TEAM_MAPPING:
                    # 檢查 J 欄 (第 10 欄)
                    j_val = get_merged_cell_value(sheet, r, 10)
                    
                    # 修正後的判定邏輯：
                    # 將結果轉為字串檢查，排除 None, 空白, 以及 "0"
                    val_str = str(j_val).strip() if j_val is not None else ""
                    
                    if val_str != "" and val_str != "0" and val_str != "None":
                        
                        # 讀取當週時間 (B~H 欄)
                        for i, col in enumerate(range(2, 9)):
                            time_val = get_merged_cell_value(sheet, r, col)
                            if time_val:
                                t_str = str(time_val).replace(" ", "").upper()
                                if t_str not in IGNORE_WORDS and len(t_str) > 2:
                                    all_records.append({
                                        "Date": dates[i],
                                        "Time": t_str,
                                        "Count": 1
                                    })
                                    
        if not all_records: 
            st.warning("未讀取到符合條件的 SUP 資料。請檢查 J 欄是否有輸入標記（非 0 或空白）。")
            return None
            
        df = pd.DataFrame(all_records)
        pivot = df.pivot_table(index="Time", columns="Date", values="Count", aggfunc="sum", fill_value=0)
        return pivot

    except Exception as e:
        st.error(f"分析錯誤: {e}")
        return None