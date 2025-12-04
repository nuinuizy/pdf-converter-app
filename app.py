import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import time
from docx import Document

# --- 1. Config ---
st.set_page_config(page_title="PDF2Word Pro", page_icon="📑", layout="centered")

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

# --- 2. Logic (ซ่อมทุกซอกทุกมุม) ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        
        # ฟังก์ชันซ่อมสระ
        def fix_sara_am(text):
            if not text or " ำ" not in text: return text
            return text.replace(" ำ", "ำ").replace(" ำ", "ำ")

        # ฟังก์ชันวนลูปแก้ Paragraphs และ Tables
        def process_container(container):
            # แก้ในย่อหน้า (Paragraphs)
            for para in container.paragraphs:
                for run in para.runs:
                    run.text = fix_sara_am(run.text)
            # แก้ในตาราง (Tables)
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        process_container(cell) # Recursive เข้าไปในเซลล์

        # 1. ซ่อมเนื้อหาหลัก (Body)
        process_container(doc)

        # 2. ซ่อม Header & Footer (หัวและท้ายกระดาษ)
        # ต้องวนลูปทุก Section เพราะเอกสารอาจมีหลายส่วน
        for section in doc.sections:
            # ซ่อม Header
            process_container(section.header)
            process_container(section.first_page_header) # เผื่อหน้าแรกไม่เหมือนหน้าอื่น
            
            # ซ่อม Footer
            process_container(section.footer)
            process_container(section.first_page_footer)

        doc.save(docx_path)
        return True
    except Exception as e:
        print(f"Repair Error: {e}")
        return False

def convert_pdf_to_docx(uploaded_file, start_page, end_page, status_box, progress_bar, join_lines):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f: f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            status_box.info("📑 เริ่มกระบวนการแกะ Layout และ Header/Footer...")
            progress_bar.progress(10)
            
            cv = Converter(pdf_path)
            if end_page is None: end_page = len(cv.pages)
            
            status_box.info(f"📄 กำลังแปลงหน้า {start_page}-{end_page}...")
            progress_bar.progress(30)
            
            # --- SETTINGS ---
            settings = {
                "multi_processing": False, 
                "parse_images": True,       # สำคัญ! เพราะ Header ชอบมีโลโก้
                # "parse_section_header_footer": True # ปกติเป็น Default อยู่แล้ว แต่เราเน้นให้มันทำงาน
            }
            
            # การจัดการย่อหน้า (ตามที่คุณขอเรื่องบรรทัด)
            if join_lines:
                settings["connected_text"] = True 
            else:
                settings["connected_text"] = False 
            
            cv.convert(docx_path, start=start_page-1, end=end_page, **settings)
            cv.close()
            
            progress_bar.progress(80)
            status_box.info("🔧 ซ่อมสระภาษาไทย (รวมหัว/ท้ายกระดาษ)...")
            repair_thai_docx(docx_path)
            progress_bar.progress(100)
            
            with open(docx_path, "rb") as f: docx_data = f.read()
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"Error: {e}")
            return None, None

# --- 3. UI ---

c1, c2 = st.columns([3, 1])
c1.markdown("### 📑 PDF to Word `Full Layout`")
c2.markdown("<div style='text-align: right; color: gray; font-size: 0.8em; padding-top: 10px;'>v3.5 Header/Footer</div>", unsafe_allow_html=True)

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
        # Checkbox เดิมที่ยังคงไว้ เพราะมีผลกับย่อหน้า
        join_lines = st.checkbox("🔗 เชื่อมประโยค (Merge Lines)", value=False, help="ไม่ติ๊ก = ยึดบรรทัดตามต้นฉบับเป๊ะๆ (แนะนำ)\nติ๊ก = ให้โปรแกรมช่วยจัดย่อหน้าใหม่")
    
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
