import streamlit as st

import io

import re

from datetime import date

from dateutil.relativedelta import relativedelta

from docx import Document

from docx.oxml.ns import qn

from docx.shared import Pt

import zipfile



# ================= 📅 設定區 =================

DATES_TO_FILL = [

    ("115.02.24", 2026, 2, 24),

    ("115.03.25", 2026, 3, 25),

    ("115.04.22", 2026, 4, 22),

    ("115.05.21", 2026, 5, 21),

    ("115.06.24", 2026, 6, 24),

]



# ================= 核心功能函式 =================



def set_cell_style_dual_font(cell, text):

    cell.text = ""

    paragraph = cell.paragraphs[0]

    run = paragraph.add_run(text)

    run.font.name = "Times New Roman"

    run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")

    run.font.size = Pt(12)



def find_birthday_in_doc(doc):

    pattern = r"生日.*?(\d{2,3})[./年](\d{1,2})[./月](\d{1,2})"

    # 搜尋段落與表格 (含巢狀)

    for para in doc.paragraphs:

        match = re.search(pattern, para.text)

        if match:

            y, m, d = match.groups()

            return date(int(y) + 1911, int(m), int(d))

    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                match = re.search(pattern, cell.text)

                if match:

                    y, m, d = match.groups()

                    return date(int(y) + 1911, int(m), int(d))

    return None



def calculate_age(birth_date_obj, target_date_obj):

    diff = relativedelta(target_date_obj, birth_date_obj)

    return f"{diff.years}Y{diff.months:02d}m{diff.days:02d}d"



def find_measurement_table(doc):

    all_tables = []

    for tbl in doc.tables:

        all_tables.append(tbl)

        for row in tbl.rows:

            for cell in row.cells:

                if cell.tables:

                    all_tables.extend(cell.tables)

    for tbl in all_tables:

        if len(tbl.rows) > 0:

            header_text = "".join([c.text for c in tbl.rows[0].cells]).replace(" ", "").replace("\n", "")

            if ("量測日期" in header_text) or ("身高" in header_text and "體重" in header_text):

                return tbl

    return None



def get_fillable_row(table, start_index):

    for i in range(start_index, len(table.rows)):

        row = table.rows[i]

        if row.cells:

            cell_text = row.cells[0].text.strip().lower()

            if "ymd" in cell_text or cell_text == "":

                return row, i

    return None, -1



# ================= Streamlit 介面 =================



st.set_page_config(page_title="自動填寫體位測量表", page_icon="📝")

st.title("📝 幼兒體位測量表自動填寫")

st.info("請上傳 Word 檔案 (.docx)，系統將自動偵測生日並填入 115 學年度測量日期與年齡。")



uploaded_files = st.file_uploader("選擇 Word 檔案", type="docx", accept_multiple_files=True)



if uploaded_files:

    if st.button("開始處理檔案"):

        processed_files = []

        progress_bar = st.progress(0)

        

        for idx, uploaded_file in enumerate(uploaded_files):

            # 讀取檔案

            doc = Document(uploaded_file)

            birth_date = find_birthday_in_doc(doc)

            

            if birth_date:

                table = find_measurement_table(doc)

                if table:

                    current_search_idx = 1

                    for date_str, y, m, d in DATES_TO_FILL:

                        target_date = date(y, m, d)

                        age_str = calculate_age(birth_date, target_date)

                        target_row, found_idx = get_fillable_row(table, current_search_idx)

                        

                        if target_row:

                            current_search_idx = found_idx + 1

                        else:

                            target_row = table.add_row()

                            current_search_idx = len(table.rows)

                        

                        if len(target_row.cells) >= 2:

                            set_cell_style_dual_font(target_row.cells[0], date_str)

                            set_cell_style_dual_font(target_row.cells[1], age_str)

                    

                    # 儲存到 BytesIO

                    output_stream = io.BytesIO()

                    doc.save(output_stream)

                    processed_files.append((uploaded_file.name, output_stream.getvalue()))

                else:

                    st.warning(f"⚠️ {uploaded_file.name}: 找不到對應表格。")

            else:

                st.warning(f"⚠️ {uploaded_file.name}: 找不到生日資訊。")

            

            progress_bar.progress((idx + 1) / len(uploaded_files))



        if processed_files:

            st.success(f"✅ 成功處理 {len(processed_files)} 個檔案！")

            

            # 建立 ZIP 下載

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w") as zf:

                for name, content in processed_files:

                    zf.writestr(name, content)

            

            st.download_button(

                label="📥 下載所有處理後的檔案 (ZIP)",

                data=zip_buffer.getvalue(),

                file_name="processed_documents.zip",

                mime="application/zip"

            )
