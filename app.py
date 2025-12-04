import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import time
from docx import Document
from docx.shared import Cm

# --- 1. Config ---
st.set_page_config(page_title="PDF2Word Pro", page_icon="💎", layout="centered")

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
        .stAlert { padding: 0.5rem; border-radius: 8px; }
        div[data-testid="column"] { gap: 0.5rem; }
    </style>
""", unsafe_allow_html=True)

# --- 2. Logic ---
def repair_and_format_docx(docx_path):
    try:
        doc = Document(docx_path)
        
        # 1. ปรับขอบกระดาษให้กว้างขึ้น (กันข้อความตกขอบ)
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(1.27)
            section.bottom_margin = Cm(1.27)
            section.left_margin = Cm(1.27)
            section.right_margin = Cm(1.27)

        # 2. Logic ซ่อมสระ ำ
        def fix_sara_am(text):
            if not text or " ำ" not in text: return text
            return text.replace(" ำ", "ำ").replace(" ำ", "ำ")

        for para in doc.paragraphs:
            for run in para.runs:
                run.text = fix_sara_am(run.text)
        
        # วนลูปแก้ในตารางด้วย
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.text = fix_sara_am(run.text)
                            
        doc.save(docx_path)
        return True
    except: return False

def convert_pdf_to_docx(uploaded_file, start_page, end_page, status_box, progress_bar, high_quality_table):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f: f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            status_box.info("⚙️ Initializing Engine...")
            progress_bar.progress(5)
            
            cv = Converter(pdf_path)
            if end_page is None: end_page = len(cv.pages)
            
            status_box.info(f"💎 กำลังแปลงหน้า {start_page}-{end_page} ด้วยความละเอียดสูง...")
            progress_bar.progress(20)
            
            # --- TWEAKED SETTINGS (จุดสำคัญ) ---
            settings = {
                "multi_processing": False, # กันค้าง
            }
            
            if high_quality_table:
                # โหมดเน้นตาราง: ใช้ lattice (เส้นขอบ) ช่วยแกะตาราง
                settings["parse_lattices_tables"] = True 
                # settings["connected_text"] = True # ช่วยเรื่องคำฉีก (แต่อาจทำให้อืด)
            
            # เริ่มแปลง
            cv.convert(docx_path, start=start_page-1, end=end_page, **settings)
            cv.close()
            
            progress_bar.progress(80)
            status_box.info("🔧 กำลังจัดหน้าและซ่อมสระ...")
            repair_and_format_docx(docx_path)
            progress_bar.progress(100)
            
            with open(docx_path, "rb") as f: docx_data = f.read()
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"Error: {e}")
            return None, None

# --- 3. UI ---

c1, c2 = st.columns([3, 1])
c1.markdown("### 💎 PDF to Word `Hi-Fi`")
c2.markdown("<div style='text-align: right; color: gray; font-size: 0.8em; padding-top: 10px;'>v3.1 High Fidelity</div>", unsafe_allow_html=True)

st.divider()

uploaded_file = st.file_uploader("Upload PDF file", type="pdf", label_visibility="collapsed")

if uploaded_file:
    # นับหน้าเร็วๆ
    try:
        from pypdf import PdfReader
        reader = PdfReader(uploaded_file)
        total_pages = len(reader.pages)
    except: total_pages = 50
    
    st.write(f"เอกสารมี **{total_pages}** หน้า")
    
    # Grid Layout
    col_mode, col_set = st.columns([1, 1])
    
    with col_mode:
        mode = st.radio("เลือกขอบเขต:", ["ทั้งหมด (All)", "เลือกหน้า (Custom)"])
        
    with col_set:
        # Checkbox ตัวช่วยเรื่อง Format
        hq_table = st.checkbox("📐 เน้นตาราง
