import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.cell.cell import MergedCell
import docx
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import calendar
import re
import io
import zipfile
from datetime import date

# ================= 🎨 1. 介面設定 =================
st.set_page_config(page_title="廣慈托嬰中心-行政自動化", page_icon="👶", layout="wide")
st.title("📊 廣慈報表自動化系統 (Excel + Word)")

with st.sidebar:
    st.header("🏢 單位設定")
    dept_options = ["不指定", "IC1", "IC2", "NIDO", "廚房", "保健室", "行政"]
    target_dept = st.selectbox("請選擇班級/單位", options=dept_options)
    
    st.header("📅 日期設定")
    target_year_roc = st.number_input("民國年份", value=115)
    target_month = st.number_input("月份", min_value=1, max_value=12, value=3)
    
    st.subheader("🛑 停托日")
    holiday_input = st.text_input("輸入日期 (如: 3/28, 4/4)")
    
    st.divider()
    st.info("💡 支援檔案：\n1. Word (.docx): 自主環境表\n2. Excel (.xlsx): 冰箱表、監測表")

# ================= 🛠️ 2. 通用日期邏輯 =================

def get_workdays(year_roc, month, holiday_str):
    year_ad = year_roc + 1911
    _, last_day = calendar.monthrange(year_ad, month)
    holidays = []
    if holiday_str:
        for item in holiday_str.replace('，', ',').split(','):
            item = item.strip()
            if '/' in item:
                try:
                    m, d = map(int, item.split('/'))
                    if m == month: holidays.append(d)
                except: continue
            else:
                try: holidays.append(int(item))
                except: continue
    workdays = []
    for d in range(1, last_day + 1):
        curr = date(year_ad, month, d)
        if curr.weekday() < 5 and d not in holidays:
            workdays.append(curr)
    return workdays

# ================= 📄 3. Word 處理邏輯 (解決方框亂碼) =================

def set_docx_font(run, size=12):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    run.font.size = Pt(size)

def process_docx(file_bytes, year, month, holidays, dept):
    doc = docx.Document(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}
    
    for p in doc.paragraphs:
        # 暴力修復標題亂碼：只要偵測到年份格式，整行重寫
        if "年" in p.text and ("月" in p.text or "11" in p.text):
            p.text = "" 
            run = p.add_run(f"{year} 年 {month:02d} 月")
            set_docx_font(run, 14)
            
        if "班級" in p.text:
            p.text = "" 
            run = p.add_run(f"班級：{dept if dept != '不指定' else '____'}")
            set_docx_font(run, 12)

    workdays = get_workdays(year, month, holidays)
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            if "日期" in row.cells[0].text or "項目" in row.cells[0].text:
                date_row = row
                weekday_row = table.rows[i+1] if i+1 < len(table.rows) else None
                for col_idx in range(1, len(date_row.cells)):
                    w_idx = col_idx - 1
                    if w_idx < len(workdays):
                        d = workdays[w_idx]
                        # 日期列
                        date_row.cells[col_idx].text = ""
                        dr = date_row.cells[col_idx].paragraphs[0].add_run(str(d.day))
                        set_docx_font(dr, 11)
                        # 星期列
                        if weekday_row:
                            weekday_row.cells[col_idx].text = ""
                            wr = weekday_row.cells[col_idx].paragraphs[0].add_run(week_map[d.weekday()])
                            set_docx_font(wr, 11)
                    else:
                        date_row.cells[col_idx].text = ""
                        if weekday_row: weekday_row.cells[col_idx].text = ""
                break
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ================= 📂 4. Excel 處理邏輯 =================

def process_excel(file_bytes, year, month, dept):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    workdays = get_workdays(year, month, "") # Excel 內部會判斷假日
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}
    
    for ws in wb.worksheets:
        # 1. 更新標題日期與班級
        for row in ws.iter_rows(min_row=1, max_row=3, min_col=1, max_col=8):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if "年" in cell.value and "月" in cell.value:
                        cell.value = f"{year}年{month:02d}月"
                    if "班級" in cell.value:
                        cell.value = f"班級：{dept}"

        # 2. 自動判斷填寫位置 (針對冰箱表)
        first_col_val = str(ws.cell(row=5, column=1).value)
        if "日期" in first_col_val or "冰箱" in str(ws.cell(row=1, column=1).value):
            curr_row = 6
            for d in workdays:
                ws.cell(row=curr_row, column=1).value = f"{d.month}/{d.day}"
                ws.cell(row=curr_row, column=2).value = week_map[d.weekday()]
                curr_row += 2
            # 清空後續
            while curr_row <= 70:
                if not isinstance(ws.cell(row=curr_row, column=1), MergedCell):
                    ws.cell(row=curr_row, column=1).value = None
                    ws.cell(row=curr_row, column=2).value = None
                curr_row += 1
                
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ================= 🚀 5. 執行介面 =================

# 確保 type 包含 xlsx
uploaded_files = st.file_uploader("📂 上傳 Word 或 Excel 報表", type=["docx", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 批次更新所有檔案"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for f in uploaded_files:
                f_ext = f.name.split('.')[-1].lower()
                try:
                    if f_ext == "docx":
                        res = process_docx(f.read(), target_year_roc, target_month, holiday_input, target_dept)
                    elif f_ext == "xlsx":
                        res = process_excel(f.read(), target_year_roc, target_month, target_dept)
                    
                    prefix = f"{target_dept}_" if target_dept != "不指定" else ""
                    zf.writestr(f"{prefix}更新_{f.name}", res)
                    st.write(f"✅ 已處理: {f.name}")
                except Exception as e:
                    st.error(f"❌ {f.name} 發生錯誤: {e}")
        
        st.download_button("📥 下載全部更新檔案 (ZIP)", data=zip_buffer.getvalue(), file_name=f"廣慈報表_{target_month}月.zip")
