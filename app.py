import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import time
from docx import Document

# --- 1. Compact Config & CSS (หัวใจสำคัญของหน้าเดียวจบ) ---
st.set_page_config(page_title="PDF2Word Pro", page_icon="⚡", layout="centered")

st.markdown("""
    <style>
        /* ลด Padding ของหน้าจอ */
        .block-container {
            padding-top: 2rem;
            padding-bottom: 1rem;
        }
        /* ปรับปุ่มให้เต็มช่อง */
        .stButton>button {
            width: 100%;
            background-color: #FF4B4B;
            color: white;
            font-weight: bold;
        }
        /* ลดระยะห่าง Elements */
        div[data-testid="column"] {
            gap: 0rem;
        }
    </style>
""", unsafe_allow_html=True)

# --- 2. Logic (เหมือนเดิม) ---
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
            status_box.info("⚙️ Initializing...")
            progress_bar.progress(10)
            
            # Conversion
            cv = Converter(pdf_path)
            cv.convert(docx_path)
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

# --- 3. Compact UI Layout ---

# Header Section (บรรทัดเดียวจบ)
c1, c2 = st.columns([3, 1])
c1.markdown("### ⚡ PDF to Word `Compact`")
c2.markdown("<div style='text-align: right; color: gray; font-size: 0.8em;'>v2.1 WS-Patch</div>", unsafe_allow_html=True)

st.divider()

# Input Section
uploaded_file = st.file_uploader("Upload PDF", type="pdf", label_visibility="collapsed")

if uploaded_file:
    # Action Area: ใช้ Grid เพื่อประหยัดที่
    # [ ปุ่ม Run ] [ Progress Bar ] [ Status Text ]
    col_btn, col_prog = st.columns([1, 2])
    
    with col_btn:
        run_btn = st.button("🚀 GO")
    
    with col_prog:
        # Placeholder ไว้รอรับ Status จะได้ไม่ดันหน้าจอลง
        status_box = st.empty()
        progress_bar = st.progress(0)

    if run_btn:
        start_time = time.time()
        docx_data, docx_name = convert_pdf_to_docx(uploaded_file, progress_bar, status_box)
        duration = time.time() - start_time
        
        if docx_data:
            st.divider()
            # Result Section: [ Metrics ] [ Download Button ]
            r1, r2 = st.columns([2, 2])
            
            with r1:
                st.caption(f"⏱️ Time: {duration:.2f}s | 📦 Size: {len(docx_data)/1024:.1f} KB")
            
            with r2:
                st.download_button(
                    label="📥 Download .DOCX",
                    data=docx_data,
                    file_name=docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
else:
    # Empty State แบบ Minimal
    st.markdown("<div style='text-align: center; color: #555; margin-top: 20px;'>Waiting for input file...</div>", unsafe_allow_html=True)
