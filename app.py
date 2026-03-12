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
st.title("📊 托育報表日期自動更新系統 (徹底修正版)")

with st.sidebar:
    st.header("🏢 單位設定")
    dept_options = ["不指定", "IC1", "IC2", "NIDO", "廚房", "保健室", "行政"]
    target_dept = st.selectbox("請選擇所屬班級/單位", options=dept_options)
    
    st.header("📅 全域日期設定")
    target_year_roc = st.number_input("設定目標民國年份", value=115)
    target_month = st.number_input("設定目標月份", min_value=1, max_value=12, value=3)
    
    st.subheader("🛑 國定假日/停托日")
    holiday_input = st.text_input("輸入日期 (例如: 3/28, 4/4)")
    
    st.divider()
    st.warning("💡 提示：此版本會自動清除亂碼符號並強制統一字體。")

# ================= 🛠️ 2. 通用邏輯 =================

def get_target_info(year_roc, month, holiday_str):
    year_ad = year_roc + 1911
    if month > 12: month, year_ad = month-12, year_ad+1
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
    return workdays

# ================= 📄 3. Word 處理邏輯 (解決方框亂碼) =================

def set_font_style(run, size=12):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    run.font.size = Pt(size)

def safe_set_word_cell(cell, text):
    cell.text = ""
    if not text: return
    p = cell.paragraphs[0]
    p.alignment = 1 # 居中
    run = p.add_run(str(text))
    set_font_style(run, 11)

def process_docx(file_bytes, target_year_roc, target_month, holiday_input, dept):
    doc = docx.Document(io.BytesIO(file_bytes))
    week_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五'}
    
    # --- A. 處理標題與班級 ---
    for p in doc.paragraphs:
        # 清洗特殊不可見字元 (解決 115„月 的問題)
        clean_p_text = p.text.replace('\x84', '').replace('\u00a0', ' ')
        
        # 更新年份月份
        if re.search(r'\d{2,3}\s*年\s*\d{1,2}\s*月', clean_p_text) or "年" in clean_p_text:
            # 強制改寫為乾淨的格式
            p.text = "" 
            run = p.add_run(f"{target_year_roc} 年 {target_month:02d} 月")
            set_font_style(run, 14) # 標題稍微大一點

        # 更新班級
        if "班級" in clean_p_text:
            label = "班級："
            display_dept = dept if dept != "不指定" else "____"
            p.text = "" 
            run = p.add_run(f"{label}{display_dept}")
            set_font_style(run, 12)

    # --- B. 處理表格 (日期一列、星期一列) ---
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            # 偵測「日期」列
            if "日期" in row.cells[0].text:
                workdays = get_target_info(target_year_roc, target_month, holiday_input)
                
                date_row = row
                # 尋找星期列 (向下找直到發現格內有星期的特徵或是固定下一列)
                # 根據您的截圖，星期在日期列的「下方」
                weekday_row = table.rows[i+1] if i+1 < len(table.rows) else None
                
                # 決定填寫的起點欄位 (通常是第2欄)
                start_col = 1
                
                # 填充資料
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
                break # 處理完一組日期就跳出

    out_sim = io.BytesIO()
    doc.save(out_sim)
    return out_sim.getvalue()

# ================= 🚀 4. 執行介面 =================

uploaded_files = st.file_uploader("📂 上傳報表 (Word)", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始批次更新"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for uploaded_file in uploaded_files:
                f_bytes = uploaded_file.read()
                try:
                    processed = process_docx(f_bytes, target_year_roc, target_month, holiday_input, target_dept)
                    prefix = f"{target_dept}_" if target_dept != "不指定" else ""
                    new_name = f"{prefix}更新_{uploaded_file.name}"
                    zf.writestr(new_name, processed)
                    st.write(f"✅ 已完成: {uploaded_file.name}")
                except Exception as e:
                    st.error(f"❌ {uploaded_file.name} 錯誤: {e}")
        
        st.success("🎉 全部處理完成！")
        st.download_button("📥 下載更新報表 (ZIP)", data=zip_buffer.getvalue(), file_name="更新報表.zip")
