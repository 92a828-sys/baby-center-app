import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.cell.cell import MergedCell
import docx
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import calendar
import re
import io
import zipfile
from datetime import date

# ================= 🎨 介面設定 =================
st.set_page_config(page_title="廣慈托嬰中心-行政自動化系統", page_icon="👶", layout="wide")
st.title("📊 托育報表日期自動更新系統")

with st.sidebar:
    st.header("🏢 單位設定")
    dept_options = ["不指定", "IC1", "IC2", "NIDO", "廚房", "保健室", "行政"]
    target_dept = st.selectbox("請選擇所屬班級/單位", options=dept_options)
    
    st.header("📅 全域日期設定")
    target_year_roc = st.number_input("設定目標民國年份", value=115)
    target_month = st.number_input("設定目標月份", min_value=1, max_value=12, value=3)
    
    st.subheader("🛑 國定假日/停托日")
    holiday_input = st.text_input("輸入日期 (例如: 3/28, 4/4)", help="用逗號隔開數字或月/日")
    
    st.divider()
    st.info("💡 提示：若幾月變成框框，請確保電腦有安裝『標楷體』。")

# ================= 🛠️ 通用邏輯 =================

def get_target_info(year_roc, month, holiday_str):
    year_ad = year_roc + 1911
    _, last_day = calendar.monthrange(year_ad, month)
    holidays_days = []
    if holiday_str:
        for item in holiday_str.replace('，', ',').split(','):
            item = item.strip()
            if '/' in item:
                try:
                    m, d = map(int, item.split('/'))
                    if m == month: holidays_days.append(d)
                except: continue
            else:
                try: holidays_days.append(int(item))
                except: continue
    workdays = []
    for d in range(1, last_day + 1):
        curr = date(year_ad, month, d)
        if curr.weekday() < 5 and d not in holidays_days:
            workdays.append(curr)
    return workdays, holidays_days, last_day

# ================= 📄 Word 核心處理邏輯 (解決框框亂碼) =================

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    for shd in tcPr.findall(qn('w:shd')):
        tcPr.remove(shd)
    new_shd = OxmlElement('w:shd')
    new_shd.set(qn('w:val'), 'clear')
    new_shd.set(qn('w:fill'), color_hex)
    tcPr.append(new_shd)

def safe_set_word_cell(cell, text, color="FFFFFF"):
    cell.text = ""
    set_cell_shading(cell, color)
    if not text: return
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    run.font.size = Pt(12)
    p.alignment = 1

def process_docx(file_bytes, target_year_roc, target_month, holiday_input, dept):
    doc = docx.Document(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}

    # 1. 更新標題 (防止框框亂碼的更新法)
    for p in doc.paragraphs:
        combined_text = "".join([run.text for run in p.runs])
        
        # 處理日期標題
        if re.search(r'\d{2,3}\s*年\s*\d{1,2}\s*月', combined_text):
            new_text = re.sub(r'(\d{2,3})(\s*年\s*)(\d{1,2})(\s*月)', 
                               f"{target_year_roc}\\2{target_month:02d}\\4", combined_text)
            # 清空並重新寫入，確保字體正確
            p.text = ""
            run = p.add_run(new_text)
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")

        # 處理班級單位
        if dept != "不指定" and "班級" in combined_text:
            label = "班級：" if "：" in combined_text else "班級 : "
            p.text = ""
            run = p.add_run(f"{label}{dept}")
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")

    # 2. 表格處理 (支援雙月填充)
    date_block_count = 0
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            # 偵測是否為日期列
            row_txt = "".join([c.text for c in row.cells])
            if any(x in row_txt[:10] for x in ["日期", "項目", "/"]):
                # 計算當前區塊月份
                calc_month = target_month + date_block_count
                calc_year = target_year_roc
                if calc_month > 12:
                    calc_month -= 12
                    calc_year += 1
                
                workdays, _, _ = get_target_info(calc_year, calc_month, holiday_input)
                date_row = row
                weekday_row = table.rows[i+1] if i+1 < len(table.rows) else None
                
                # 尋找數字起點
                start_col = 1
                for c_idx, cell in enumerate(date_row.cells):
                    if cell.text.strip().isdigit():
                        start_col = c_idx
                        break
                
                # 填入內容
                for col_i in range(start_col, len(date_row.cells)):
                    idx = col_i - start_col
                    if idx < len(workdays):
                        d = workdays[idx]
                        safe_set_word_cell(date_row.cells[col_i], str(d.day))
                        if weekday_row:
                            safe_set_word_cell(weekday_row.cells[col_i], week_map[d.weekday()])
                    else:
                        safe_set_word_cell(date_row.cells[col_i], "")
                        if weekday_row:
                            safe_set_word_cell(weekday_row.cells[col_i], "")
                
                date_block_count += 1 # 準備下一個日期區塊 (針對雙面表單)

    out_sim = io.BytesIO()
    doc.save(out_sim)
    return out_sim.getvalue()

# ================= 📂 Excel 邏輯 (略，保持原有功能) =================
def process_excel(file_bytes, year_roc, month, workdays, dept):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=10):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = re.sub(r'\d{2,3}\s*[年./-]\s*\d{1,2}\s*月?', f"{year_roc}年{month:02d}月", cell.value)
                    if dept != "不指定" and "班級" in cell.value: cell.value = f"班級：{dept}"
        ws_content = "".join([str(cell.value) for row in ws.iter_rows(max_row=5, max_col=5) for cell in row if cell.value])
        if "冰箱" in ws_content:
            curr_row = 6
            for d in workdays:
                ws.cell(row=curr_row, column=1).value = f"{d.month}/{d.day}"
                ws.cell(row=curr_row, column=2).value = week_map[d.weekday()]
                curr_row += 2
            while curr_row <= 70:
                if not isinstance(ws.cell(row=curr_row, column=1), MergedCell):
                    ws.cell(row=curr_row, column=1).value = None
                    ws.cell(row=curr_row, column=2).value = None
                curr_row += 1
    out_sim = io.BytesIO()
    wb.save(out_sim)
    return out_sim.getvalue()

# ================= 🚀 執行介面 =================
uploaded_files = st.file_uploader("📂 上傳報表", type=["xlsx", "docx"], accept_multiple_files=True)
if uploaded_files:
    if st.button("🚀 開始批次更新"):
        main_workdays, _, _ = get_target_info(target_year_roc, target_month, holiday_input)
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for uploaded_file in uploaded_files:
                f_bytes = uploaded_file.read()
                try:
                    if uploaded_file.name.endswith(".xlsx"):
                        processed = process_excel(f_bytes, target_year_roc, target_month, main_workdays, target_dept)
                    else:
                        processed = process_docx(f_bytes, target_year_roc, target_month, holiday_input, target_dept)
                    prefix = f"{target_dept}_" if target_dept != "不指定" else ""
                    zf.writestr(f"{prefix}更新_{uploaded_file.name}", processed)
                    st.write(f"✅ 已完成: {uploaded_file.name}")
                except Exception as e:
                    st.error(f"❌ {uploaded_file.name} 錯誤: {e}")
        st.success("🎉 全部處理完成！")
        st.download_button("📥 下載 ZIP 檔案", data=zip_buffer.getvalue(), file_name="更新報表.zip")
