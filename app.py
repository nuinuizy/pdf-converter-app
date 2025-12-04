import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import time
from docx import Document

# --- 1. Config ---
st.set_page_config(page_title="PDF2Word Exact", page_icon="📏", layout="centered")

st.markdown("""
    <style>
        .block-container { padding-top: 2rem; padding-bottom: 2rem; }
        .stButton>button { 
            width: 100%; 
            background-color: #000000; 
            color: white; 
            font-weight: bold; 
            border-radius: 8px; 
            height: 50px;
        }
        div[data-testid="column"] { gap: 0.5rem; }
    </style>
""", unsafe_allow_html=True)

# --- 2. Logic ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        # ซ่อมสระ ำ อย่างเดียว ไม่แตะ Layout
        def fix_sara_am(text):
            if not text or " ำ" not in text: return text
            return text.replace(" ำ", "ำ").replace(" ำ", "ำ")

        for para in doc.paragraphs:
            for run in para.runs:
                run.text = fix_sara_am(run.text)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.text = fix_sara_am(run.text)
        doc.save(docx_path)
        return True
    except: return False

def convert_pdf_to_docx(uploaded_file, start_page, end_page, status_box, progress_bar, join_lines):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f: f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            status_box.info("📏 เริ่มต้นกระบวนการแม่นยำสูง...")
            progress_bar.progress(10)
            
            cv = Converter(pdf_path)
            if end_page is None: end_page = len(cv.pages)
            
            status_box.info(f"📄 กำลังแปลงหน้า {start_page}-{end_page}...")
            progress_bar.progress(30)
            
            # --- SETTINGS ตามสั่ง ---
            settings = {
                "multi_processing": False, 
                "parse_images": True,
            }
            
            # [จุดตัดสินใจ]
            if join_lines:
                # ถ้า User สั่งเชื่อมบรรทัด (Flow Text)
                settings["connected_text"] = True 
            else:
                # [ค่า Default] เอาแบบบรรทัดต่อบรรทัด (Exact Line)
                settings["connected_text"] = False 
            
            cv.convert(docx_path, start=start_page-1, end=end_page, **settings)
            cv.close()
            
            progress_bar.progress(80)
            status_box.info("🔧 ซ่อมสระภาษาไทย...")
            repair_thai_docx(docx_path)
            progress_bar.progress(100)
            
            with open(docx_path, "rb") as f: docx_data = f.read()
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"Error: {e}")
            return None, None

# --- 3. UI ---

c1, c2 = st.columns([3, 1])
c1.markdown("### 📏 PDF to Word `Exact`")
c2.markdown("<div style='text-align: right; color: gray; font-size: 0.8em; padding-top: 10px;'>v3.4 Exact Lines</div>", unsafe_allow_html=True)

st.divider()

uploaded_file = st.file_uploader("Upload PDF file", type="pdf", label_visibility="collapsed")

if uploaded_file:
    # นับหน้า
    try:
        from pypdf import PdfReader
        reader = PdfReader(uploaded_file)
        total_pages = len(reader.pages)
    except: total_pages = 50
    
    st.write(f"เอกสารมี **{total_pages}** หน้า")
    
    col_mode, col_opt = st.columns([1, 1])
    
    with col_mode:
        mode = st.radio("เลือกขอบเขต:", ["ทั้งหมด (All)", "เลือกหน้า (Custom)"])
        
    with col_opt:
        # Checkbox นี้สำคัญ! ผมตั้งค่าเริ่มต้นเป็น "ไม่ติ๊ก" (False) 
        # เพื่อให้มันเคาะบรรทัดตามต้นฉบับเป๊ะๆ ตามที่คุณขอ
        join_lines = st.checkbox("🔗 เชื่อมประโยค (Merge Lines)", value=False, help="ถ้าติ๊ก: จะพยายามรวมบรรทัดให้เป็นย่อหน้าเดียว\nถ้าไม่ติ๊ก: จะขึ้นบรรทัดใหม่ตามต้นฉบับเป๊ะๆ")
    
    start_p, end_p = 1, None
    if mode == "เลือกหน้า (Custom)":
        c_s, c_e = st.columns(2)
        with c_s: start_p = st.number_input("หน้าแรก", 1, total_pages, 1)
        with c_e: end_p = st.number_input("ถึงหน้า", start_p, total_pages, min(start_p+4, total_pages))
    
    st.markdown("---")
    
    if st.button("🚀 แปลงไฟล์ (Convert)"):
        status_box = st.empty()
        progress_bar = st.empty()
        start_time = time.time()
        
        docx_data, docx_name = convert_pdf_to_docx(uploaded_file, start_p, end_p, status_box, progress_bar, join_lines)
        
        if docx_data:
            duration = time.time() - start_time
            status_box.success("✅ เสร็จสิ้น!")
            
            c1, c2 = st.columns([1, 1])
            with c1: st.caption(f"Time: {duration:.2f}s | Size: {len(docx_data)/1024:.1f} KB")
            with c2:
                st.download_button("📥 Download Word", docx_data, docx_name, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
