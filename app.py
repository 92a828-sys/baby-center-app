import streamlit as st
import io
import re
import zipfile
from datetime import date
from dateutil.relativedelta import relativedelta

# 嘗試匯入必要的 Word 處理套件
try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Pt
except ImportError:
    st.error("找不到 python-docx 套件，請確保 requirements.txt 中已加入此套件。")

# ================= 📅 核心設定區 =================
# 定義 115 學年度的五個量測基準日
DATES_TO_FILL = [
    ("115.02.24", 2026, 2, 24),
    ("115.03.25", 2026, 3, 25),
    ("115.04.22", 2026, 4, 22),
    ("115.05.21", 2026, 5, 21),
    ("115.06.24", 2026, 6, 24),
]

# ================= 核心邏輯函式 =================

def set_cell_style_dual_font(cell, text):
    """設定雙字型：英數 Times New Roman，中文 標楷體，12pt"""
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(text)
    run.font.name = "Times New Roman"
    # 強制設定東亞字型為標楷體
    run._element.rPr.rFonts.set(qn('w:eastAsia'), "標楷體")
    run.font.size = Pt(12)

def find_birthday_in_doc(doc):
    """搜尋文件中隱藏的生日資訊 (含段落、表格、巢狀表格)"""
    pattern = r"生日.*?(\d{2,3})[./年](\d{1,2})[./月](\d{1,2})"
    
    # 掃描段落
    for para in doc.paragraphs:
        match = re.search(pattern, para.text)
        if match:
            y, m, d = match.groups()
            return date(int(y) + 1911, int(m), int(d))
            
    # 掃描所有表格 (包含巢狀結構)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                match = re.search(pattern, cell.text)
                if match:
                    y, m, d = match.groups()
                    return date(int(y) + 1911, int(m), int(d))
                # 遞迴檢查格子內的巢狀表格
                for nested_table in cell.tables:
                    for n_row in nested_table.rows:
                        for n_cell in n_row.cells:
                            match = re.search(pattern, n_cell.text)
                            if match:
                                y, m, d = match.groups()
                                return date(int(y) + 1911, int(m), int(d))
    return None

def calculate_age(birth_date, target_date):
    """計算精確年齡 yYmmMddD 格式"""
    diff = relativedelta(target_date, birth_date)
    return f"{diff.years}Y{diff.months:02d}m{diff.days:02d}d"

def find_measurement_table(doc):
    """終極深層掃描：尋找包含「量測日期」或「身高/體重」的表格"""
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
    """尋找可填寫的預留列 (檢查首格是否含 ymd 或為空)"""
    for i in range(start_index, len(table.rows)):
        row = table.rows[i]
        if row.cells:
            cell_text = row.cells[0].text.strip().lower()
            if "ymd" in cell_text or cell_text == "":
                return row, i
    return None, -1

# ================= Streamlit 網頁介面 =================

st.set_page_config(page_title="自動填寫體位表", page_icon="📖", layout="wide")

# 側邊欄：顯示設定與說明
with st.sidebar:
    st.header("⚙️ 設定資訊")
    st.write("目前設定填寫日期：")
    for d_str, _, _, _ in DATES_TO_FILL:
        st.write(f"- {d_str}")
    st.divider()
    st.write("💡 **提示**：上傳多個檔案後，系統會自動處理並壓縮成 ZIP 供您下載。")

st.title("📖 幼兒體位測量表 — 自動化填寫工具")
st.markdown("""
本工具會自動掃描 Word 檔中的**學生生日**，計算與**指定量測日**的年齡差，並填入文件中的**體位測量表格**。
""")

# 檔案上傳
uploaded_files = st.file_uploader("請拖曳或選擇 .docx 檔案", type="docx", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始批次處理", use_container_width=True):
        processed_zip_buffer = io.BytesIO()
        success_count = 0
        error_logs = []

        # 進度條
        progress_bar = st.progress(0)
        
        with zipfile.ZipFile(processed_zip_buffer, "w") as zf:
            for idx, uploaded_file in enumerate(uploaded_files):
                try:
                    # 讀取檔案至記憶體
                    file_content = uploaded_file.read()
                    doc = Document(io.BytesIO(file_content))
                    
                    # 1. 找生日
                    birth_date = find_birthday_in_doc(doc)
                    if not birth_date:
                        error_logs.append(f"❌ {uploaded_file.name}: 找不到生日資訊。")
                        continue
                    
                    # 2. 找表格
                    table = find_measurement_table(doc)
                    if not table:
                        error_logs.append(f"❌ {uploaded_file.name}: 找不到量測表格。")
                        continue
                    
                    # 3. 填入資料
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
                        
                        # 填寫前兩格 (日期、年齡)
                        if len(target_row.cells) >= 2:
                            set_cell_style_dual_font(target_row.cells[0], date_str)
                            set_cell_style_dual_font(target_row.cells[1], age_str)
                    
                    # 4. 儲存結果到 ZIP
                    temp_out = io.BytesIO()
                    doc.save(temp_out)
                    zf.writestr(uploaded_file.name, temp_out.getvalue())
                    success_count += 1
                    
                except Exception as e:
                    error_logs.append(f"💥 {uploaded_file.name} 處理時發生未預期錯誤: {str(e)}")
                
                # 更新進度
                progress_bar.progress((idx + 1) / len(uploaded_files))

        # 顯示結果
        if success_count > 0:
            st.success(f"🎉 處理完成！共成功處理 {success_count} 個檔案。")
            st.download_button(
                label="📥 下載處理完成的檔案 (ZIP)",
                data=processed_zip_buffer.getvalue(),
                file_name=f"processed_files_{date.today()}.zip",
                mime="application/zip",
                use_container_width=True
            )
        
        if error_logs:
            with st.expander("⚠️ 查看錯誤紀錄"):
                for log in error_logs:
                    st.write(log)
