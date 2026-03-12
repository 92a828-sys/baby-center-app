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

# ================= 🎨 1. 介面設定 =================
st.set_page_config(page_title="廣慈托嬰中心-行政自動化系統", page_icon="👶", layout="wide")
st.title("📊 托育報表日期自動更新系統 (終極修正版)")

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
    st.info("💡 說明：系統會自動將標題修正為標楷體，解決月份變框框的問題。")

# ================= 🛠️ 2. 通用邏輯 =================

def get_target_info(year_roc, month, holiday_str):
    year_ad = year_roc + 1911
    # 處理跨年月份
    if month > 12:
        month -= 12
        year_ad += 1
    elif month < 1:
        month = 1
        
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

# ================= 📄 3. Word 處理邏輯 (核心修正) =================

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
    run.font.size = Pt(11)
    p.alignment = 1

def process_docx(file_bytes, target_year_roc, target_month, holiday_input, dept):
    doc = docx.Document(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}
    
    # --- A. 更新標題與班級 ---
    for p in doc.paragraphs:
        full_text = p.text
        # 1. 處理月份標題 (防止框框)
        if re.search(r'\d{2,3}\s*年\s*\d{1,2}\s*月', full_text):
            # 取得原始間隔符號
            new_title = re.sub(r'(\d{2,3})(\s*年\s*)(\d{1,2})(\s*月)', 
                               f"{target_year_roc}\\2{target_month:02d}\\4", full_text)
            p.text = "" # 清空重寫，避免字體損壞
            run = p.add_run(new_title)
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
            run.font.size = Pt(12)

        # 2. 處理班級單位
        if "班級" in full_text:
            label = "班級：" if "：" in full_text else "班級 : "
            display_dept = dept if dept != "不指定" else "____"
            p.text = "" 
            run = p.add_run(f"{label}{display_dept}")
            run.font.name = "Times New Roman"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
            run.font.size = Pt(12)

    # --- B. 處理表格內容 ---
    date_block_idx = 0
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            try:
                first_cell_txt = row.cells[0].text.strip()
                # 偵測日期列 (自主環境檢核表專用)
                if any(x in first_cell_txt for x in ["日期", "項目", "/"]):
                    
                    # 計算該區塊對應月份
                    curr_m = target_month + date_block_idx
                    workdays, _, _ = get_target_info(target_year_roc, curr_m, holiday_input)
                    
                    date_row = row
                    weekday_row = table.rows[i+1] if i+1 < len(table.rows) else None
                    
                    # 尋找起始欄 (避開左側文字)
                    start_col = 1
                    for c_idx in range(len(date_row.cells)):
                        if date_row.cells[c_idx].text.strip().isdigit():
                            start_col = c_idx
                            break
                    
                    # 填充日期與星期
                    for col_idx in range(start_col, len(date_row.cells)):
                        w_idx = col_idx - start_col
                        if w_idx < len(workdays):
                            d = workdays[w_idx]
                            safe_set_word_cell(date_row.cells[col_idx], str(d.day))
                            if weekday_row:
                                safe_set_word_cell(weekday_row.cells[col_idx], week_map[d.weekday()])
                        else:
                            # 清空多餘格子
                            safe_set_word_cell(date_row.cells[col_idx], "")
                            if weekday_row:
                                safe_set_word_cell(weekday_row.cells[col_idx], "")
                    
                    date_block_idx += 1
            except:
                continue

    out_sim = io.BytesIO()
    doc.save(out_sim)
    return out_sim.getvalue()

# ================= 📂 4. Excel 處理邏輯 =================

def process_excel(file_bytes, year_roc, month, workdays, dept):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}
    
    for ws in wb.worksheets:
        # 標題更新
        for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=10):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = re.sub(r'\d{2,3}\s*[年./-]\s*\d{1,2}\s*月?', f"{year_roc}年{month:02d}月", cell.value)
                    if dept != "不指定" and "班級" in cell.value:
                        cell.value = f"班級：{dept}"

        # 冰箱表或監測表內容更新
        ws_content = "".join([str(cell.value) for row in ws.iter_rows(max_row=5, max_col=5) for cell in row if cell.value])
        
        if "冰箱" in ws_content:
            curr_row = 6
            for d in workdays:
                ws.cell(row=curr_row, column=1).value = f"{d.month}/{d.day}"
                ws.cell(row=curr_row, column=2).value = week_map[d.weekday()]
                curr_row += 2
            # 清空剩餘列
            while curr_row <= 70:
                if not isinstance(ws.cell(row=curr_row, column=1), MergedCell):
                    ws.cell(row=curr_row, column=1).value = None
                    ws.cell(row=curr_row, column=2).value = None
                curr_row += 1

    out_sim = io.BytesIO()
    wb.save(out_sim)
    return out_sim.getvalue()

# ================= 🚀 5. 執行介面 =================

uploaded_files = st.file_uploader("📂 上傳報表 (Excel 或 Word)", type=["xlsx", "docx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始批次更新"):
        workdays, _, _ = get_target_info(target_year_roc, target_month, holiday_input)
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for uploaded_file in uploaded_files:
                fname = uploaded_file.name
                f_bytes = uploaded_file.read()
                try:
                    if fname.endswith(".xlsx"):
                        processed = process_excel(f_bytes, target_year_roc, target_month, workdays, target_dept)
                    elif fname.endswith(".docx"):
                        processed = process_docx(f_bytes, target_year_roc, target_month, holiday_input, target_dept)
                    
                    prefix = f"{target_dept}_" if target_dept != "不指定" else ""
                    new_name = f"{prefix}更新_{fname}"
                    zf.writestr(new_name, processed)
                    st.write(f"✅ 已完成: {new_name}")
                except Exception as e:
                    st.error(f"❌ {fname} 處理出錯: {e}")
        
        st.success("🎉 全部處理完成！")
        st.download_button(
            label="📥 下載更新後的報表 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"廣慈報表_{target_year_roc}年{target_month}月.zip",
            mime="application/zip"
        )
