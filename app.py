def process_docx(file_bytes, target_year_roc, target_month, holiday_input, dept):
    """專門優化：解決框框亂碼並精準填入班級與日期表格"""
    doc = docx.Document(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}

    # 1. 更新頁首標題與班級
    for p in doc.paragraphs:
        full_p_text = p.text
        
        # --- 處理月份標題 (解決框框問題) ---
        if re.search(r'\d{2,3}\s*年\s*\d{1,2}\s*月', full_p_text):
            # 找到原本的格式內容，只更換數字
            new_text = re.sub(r'(\d{2,3})(\s*年\s*)(\d{1,2})(\s*月)', 
                               f"{target_year_roc}\\2{target_month:02d}\\4", full_p_text)
            p.text = "" # 清空舊的
            run = p.add_run(new_text)
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
            run.font.size = Pt(12)

        # --- 處理班級/單位 (精準填入) ---
        if "班級" in full_p_text:
            label = "班級：" if "：" in full_p_text else "班級 : "
            display_dept = dept if dept != "不指定" else "____"
            p.text = "" # 清空舊的
            run = p.add_run(f"{label}{display_dept}")
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
            run.font.size = Pt(12)

    # 2. 處理表格 (日期在上、星期在下)
    # 建立一個計數器來支援雙月表單 (第一面填目標月, 第二面填次月)
    date_block_idx = 0

    for table in doc.tables:
        for i, row in enumerate(table.rows):
            try:
                first_cell_txt = row.cells[0].text.strip()
                # 偵測日期列起點
                if any(x in first_cell_txt for x in ["日期", "項目", "/"]):
                    
                    # 決定這個區塊的月份
                    curr_m = target_month + date_block_idx
                    curr_y = target_year_roc
                    if curr_m > 12:
                        curr_m -= 12
                        curr_y += 1
                    
                    workdays, _, _ = get_target_info(curr_y, curr_m, holiday_input)
                    
                    date_row = row
                    weekday_row = table.rows[i+1] if i+1 < len(table.rows) else None
                    
                    # 尋找填寫起始欄位 (避開左側標題)
                    start_col = 1
                    for c_idx in range(len(date_row.cells)):
                        if date_row.cells[c_idx].text.strip().isdigit():
                            start_col = c_idx
                            break
                    
                    # 填充工作日
                    for col_idx in range(start_col, len(date_row.cells)):
                        workday_idx = col_idx - start_col
                        if workday_idx < len(workdays):
                            d = workdays[workday_idx]
                            safe_set_word_cell(date_row.cells[col_idx], str(d.day))
                            if weekday_row:
                                safe_set_word_cell(weekday_row.cells[col_idx], week_map[d.weekday()])
                        else:
                            # 清空多餘格子
                            safe_set_word_cell(date_row.cells[col_idx], "")
                            if weekday_row:
                                safe_set_word_cell(weekday_row.cells[col_idx], "")
                    
                    date_block_idx += 1 # 準備下一個區塊
            except:
                continue

    out_sim = io.BytesIO()
    doc.save(out_sim)
    return out_sim.getvalue()
