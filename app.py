import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import time
from docx import Document

# --- 1. Pro Config ---
st.set_page_config(page_title="PDF2Word Pro", page_icon="⚡", layout="centered")

st.markdown("""
    <style>
        .block-container { padding-top: 2rem; padding-bottom: 1rem; }
        .stButton>button { width: 100%; background-color: #FF4B4B; color: white; font-weight: bold; }
        div[data-testid="column"] { gap: 0.5rem; }
    </style>
""", unsafe_allow_html=True)

# --- 2. Logic (Optimized) ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        def fix_sara_am(text):
            if not text: return text
            text = text.replace(" ำ", "ำ") 
            text = text.replace(" ำ", "ำ")
            return text

        for para in doc.paragraphs:
            for run in para.runs:
                if run.text: run.text = fix_sara_am(run.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.text: run.text = fix_sara_am(run.text)
        doc.save(docx_path)
        return True
    except: return False

def convert_pdf_to_docx(uploaded_file, progress_bar, status_box):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f: f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            status_box.info("⚙️ Engine Starting...")
            progress_bar.progress(10)
            
            # Conversion
            cv = Converter(pdf_path)
            
            # --- จุดแก้ความช้า (FIXED) ---
            # multi_processing=False: บังคับใช้ Core เดียว ลดภาระเครื่องฟรี ทำให้ไม่ค้าง
            # cpu_count=1: ย้ำอีกทีว่าขอใช้แค่ 1 แรงงานพอ
            cv.convert(docx_path, multi_processing=False, cpu_count=1)
            
            cv.close()
            progress_bar.progress(60)
            status_box.info("🔧 Patching Thai Vowels...")
            
            # Patching
            repair_thai_docx(docx_path)
            progress_bar.progress(100)
            status_box.success("✅ Complete!")
            
            with open(docx_path, "rb") as f: docx_data = f.read()
            return docx_data, docx_name
        except Exception as e:
            status_box.error(f"Error: {e}")
            return None, None

# --- 3. UI Layout ---

c1, c2 = st.columns([3, 1])
c1.markdown("### ⚡ PDF to Word `Pro`")
c2.markdown("<div style='text-align: right; color: gray; font-size: 0.8em; padding-top: 10px;'>v2.1 Fast</div>", unsafe_allow_html=True)

st.divider()

uploaded_file = st.file_uploader("Upload PDF", type="pdf", label_visibility="collapsed")

if uploaded_file:
    # Action Area
    col_btn, col_prog = st.columns([1, 2])
    
    with col_btn:
        run_btn = st.button("🚀 START")
    
    with col_prog:
        status_box = st.empty()
        progress_bar = st.progress(0)

    if run_btn:
        start_time = time.time()
        docx_data, docx_name = convert_pdf_to_docx(uploaded_file, progress_bar, status_box)
        duration = time.time() - start_time
        
        if docx_data:
            st.divider()
            r1, r2 = st.columns([2, 2])
            with r1:
                st.caption(f"⏱️ Processing time: {duration:.2f}s | 📦 Size: {len(docx_data)/1024:.1f} KB")
            with r2:
                st.download_button(
                    label="📥 Download Result",
                    data=docx_data,
                    file_name=docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
else:
    # Idle State
    st.markdown("<div style='text-align: center; font-size: 2em;'>🚀 🚀 🚀</div>", unsafe_allow_html=True)
    st.markdown("<div style='text-align: center; color: #555; margin-top: 5px;'>Waiting for input file...</div>", unsafe_allow_html=True)
