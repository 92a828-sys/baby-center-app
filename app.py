import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.cell.cell import MergedCell
import docx
from docx.shared import Pt
from docx.oxml.ns import qn
import os
import calendar
import re
import io
import zipfile
from datetime import date

# ================= 🎨 介面設定 =================
st.set_page_config(page_title="廣慈托嬰中心-行政自動化系統", page_icon="👶", layout="wide")
st.title("📊 托育報表日期自動更新系統")
st.markdown("支援範圍：幼生監測表、冰箱記錄表 (Excel) 及 環境自主檢核表 (Word)")

with st.sidebar:
    st.header("📅 全域設定")
    target_year_roc = st.number_input("設定目標民國年份", value=115)
    target_month = st.number_input("設定目標月份", min_value=1, max_value=12, value=3)
    
    st.subheader("🛑 國定假日/停托日")
    holiday_input = st.text_input("輸入日期 (例如: 3/28, 4/4)", help="用逗號隔開，程式會自動跳過這些日子")
    
    st.divider()
    st.info("💡 系統會自動識別：\n1. .xlsx -> 監測表/冰箱表\n2. .docx -> 環境檢核表")

# ================= 🛠️ 通用邏輯 =================

def get_workdays_list(year_roc, month, holiday_str):
    """取得該月工作日列表 (date 物件)"""
    year_ad = year_roc + 1911
    _, last_day = calendar.monthrange(year_ad, month)
    holidays = []
    if holiday_str:
        for item in holiday_str.replace('，', ',').split(','):
            try:
                m, d = item.strip().split('/')
                holidays.append((int(m), int(d)))
            except: continue
    
    workdays = []
    for d in range(1, last_day + 1):
        curr = date(year_ad, month, d)
        if curr.weekday() < 5 and (month, d) not in holidays:
            workdays.append(curr)
    return workdays

# ================= 📂 Excel 處理邏輯 (監測表/冰箱表) =================

def process_excel(file_bytes, year_roc, month, workdays):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}
    font_eng = Font(name='Times New Roman', size=12)
    font_chi = Font(name='標楷體', size=12)
    align_center = Alignment(horizontal='center', vertical='center')

    for ws in wb.worksheets:
        # 1. 標題替換
        for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=2):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = re.sub(r'\d{2,3}\s*[年./-]\s*\d{1,2}\s*月?', f"{year_roc}年{month:02d}月", cell.value)

        # 2. 內容填充
        ws_content = "".join([str(cell.value) for row in ws.iter_rows(max_row=10, max_col=5) for cell in row if cell.value])
        
        if "冰箱" in ws_content:
            curr_row = 6
            for d in workdays:
                ws.cell(row=curr_row, column=1).value = f"{d.month}/{d.day}"
                ws.cell(row=curr_row, column=2).value = week_map[d.weekday()]
                ws.cell(row=curr_row+1, column=1).value = None
                ws.cell(row=curr_row+1, column=2).value = None
                curr_row += 2
            while curr_row <= 70:
                if not isinstance(ws.cell(row=curr_row, column=1), MergedCell):
                    ws.cell(row=curr_row, column=1).value = None
                    ws.cell(row=curr_row, column=2).value = None
                curr_row += 1

        elif "監測" in ws_content or "星期" in ws_content:
            date_row = 3
            start_col = 4
            for r in range(1, 6):
                row_vals = [str(ws.cell(row=r, column=c).value) for c in range(1, 6)]
                if any("星期" in v for v in row_vals) or any("日期" in v for v in row_vals):
                    date_row = r
                    break
            for i, d in enumerate(workdays):
                col = start_col + i
                c_day = ws.cell(row=date_row, column=col)
                if not isinstance(c_day, MergedCell):
                    c_day.value = d.day
                    c_day.font = font_eng
                    c_day.alignment = align_center
                c_week = ws.cell(row=date_row+1, column=col)
                if not isinstance(c_week, MergedCell):
                    c_week.value = week_map[d.weekday()]
                    c_week.font = font_chi
                    c_week.alignment = align_center
            for col in range(start_col + len(workdays), 40):
                c_check = ws.cell(row=date_row, column=col)
                if c_check.value and ("餵藥" in str(c_check.value) or "統計" in str(c_check.value)): break
                if not isinstance(c_check, MergedCell): c_check.value = None
                c_next = ws.cell(row=date_row+1, column=col)
                if not isinstance(c_next, MergedCell): c_next.value = None

    out_sim = io.BytesIO()
    wb.save(out_sim)
    return out_sim.getvalue()

# ================= 📄 Word 處理邏輯 (環境檢核表) =================

def safe_set_word_cell(cell, text):
    cell.text = ""
    if not text: return
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    run.font.size = Pt(12)
    paragraph.alignment = 1

def process_docx(file_bytes, year_roc, month, workdays):
    doc = docx.Document(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}
    
    # 1. 更新大標題
    for p in doc.paragraphs:
        pattern = r'11[0-9]\s*年\s*(?:1[0-2]|0?[1-9])?\s*月'
        if re.search(pattern, p.text):
            new_text = re.sub(pattern, f"{year_roc} 年 {month:02d} 月", p.text)
            p.text = ""
            run = p.add_run(new_text)
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
            run.font.size = Pt(12)

    # 2. 更新表格
    for table in doc.tables:
        date_row = None
        weekday_row = None
        start_col = 2
        for i, row in enumerate(table.rows):
            row_text = "".join([c.text.strip() for c in row.cells])
            if "日期" in row_text or "項目" in row_text:
                date_row = row
                if i + 1 < len(table.rows): weekday_row = table.rows[i + 1]
                for c_idx, cell in enumerate(date_row.cells):
                    if cell.text.strip().isdigit():
                        start_col = c_idx
                        break
                break
        
        if date_row and weekday_row:
            for i in range(start_col, len(date_row.cells)):
                idx = i - start_col
                if idx < len(workdays):
                    d = workdays[idx]
                    safe_set_word_cell(date_row.cells[i], str(d.day))
                    safe_set_word_cell(weekday_row.cells[i], week_map[d.weekday()])
                else:
                    safe_set_word_cell(date_row.cells[i], "")
                    safe_set_word_cell(weekday_row.cells[i], "")

    out_sim = io.BytesIO()
    doc.save(out_sim)
    return out_sim.getvalue()

# ================= 🚀 執行介面 =================

uploaded_files = st.file_uploader("📂 上傳報表 (Excel 或 Word)", type=["xlsx", "docx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始批次更新"):
        workdays = get_workdays_list(target_year_roc, target_month, holiday_input)
        
        if not workdays:
            st.error("❌ 找不到工作日，請檢查月份設定。")
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for uploaded_file in uploaded_files:
                    fname = uploaded_file.name
                    f_bytes = uploaded_file.read()
                    
                    try:
                        if fname.endswith(".xlsx"):
                            processed_data = process_excel(f_bytes, target_year_roc, target_month, workdays)
                        elif fname.endswith(".docx"):
                            processed_data = process_docx(f_bytes, target_year_roc, target_month, workdays)
                        
                        new_name = f"更新_{target_year_roc}年{target_month}月_{fname}"
                        zf.writestr(new_name, processed_data)
                        st.write(f"✅ 已處理: {fname}")
                    except Exception as e:
                        st.error(f"❌ 處理 {fname} 時出錯: {e}")
            
            st.success(f"🎉 全部處理完成！共 {len(uploaded_files)} 個檔案。")
            st.download_button(
                label="📥 下載更新後的報表 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"廣慈報表整合包_{target_year_roc}年{target_month}月.zip",
                mime="application/zip"
            )
