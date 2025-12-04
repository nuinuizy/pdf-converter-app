import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import time
from docx import Document

# --- 1. Config ---
st.set_page_config(page_title="PDF2Word Pro", page_icon="⚡", layout="centered")

st.markdown("""
    <style>
        .block-container { padding-top: 2rem; padding-bottom: 2rem; }
        .stButton>button { 
            width: 100%; 
            background-color: #FF4B4B; 
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

def convert_pdf_to_docx(uploaded_file, status_box, turbo_mode):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f: f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            cv = Converter(pdf_path)
            
            # --- TURBO MODE LOGIC ---
            if turbo_mode:
                # ตัดรูปภาพออก เพื่อความเร็วสูงสุด
                settings = {"parse_images": False}
                cv.convert(docx_path, multi_processing=False, **settings)
            else:
                # แบบปกติ (ช้าหน่อย แต่ได้ครบ)
                cv.convert(docx_path, multi_processing=False)
                
            cv.close()
            
            status_box.info("🔧 กำลังซ่อมสระภาษาไทย...")
            repair_thai_docx(docx_path)
            
            with open(docx_path, "rb") as f: docx_data = f.read()
            return docx_data, docx_name
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            return None, None

# --- 3. UI (v2.5 Turbo) ---

c1, c2 = st.columns([3, 1])
c1.markdown("### ⚡ PDF to Word `Pro`")
c2.markdown("<div style='text-align: right; color: gray; font-size: 0.8em; padding-top: 10px;'>v2.5 Turbo</div>", unsafe_allow_html=True)

st.divider()

uploaded_file = st.file_uploader("Upload PDF file", type="pdf", label_visibility="collapsed")

if uploaded_file:
    # Checkbox เลือกโหมด
    turbo = st.checkbox("⚡ Turbo Mode (ตัดรูปภาพออกเพื่อให้เร็วขึ้น)", value=True)
    
    run_btn = st.button("🚀 เริ่มแปลงไฟล์ (Start)")
    status_box = st.empty()

    if run_btn:
        status_box.info("⏳ กำลังทำงาน... (44 หน้าอาจใช้เวลา 2-3 นาที)")
        start_time = time.time()
        
        docx_data, docx_name = convert_pdf_to_docx(uploaded_file, status_box, turbo)
        
        duration = time.time() - start_time
        
        if docx_data:
            status_box.success("✅ เสร็จเรียบร้อย!")
            st.divider()
            
            c_info, c_btn = st.columns([1.5, 2])
            with c_info:
                st.caption(f"⏱️ ใช้เวลา: {duration:.2f}s")
                st.caption(f"📦 ขนาด: {len(docx_data)/1024:.1f} KB")
            with c_btn:
                st.download_button(
                    label="📥 ดาวน์โหลด Word (.docx)",
                    data=docx_data,
                    file_name=docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
else:
    st.markdown(
        """
        <div style='text-align: center; color: #666; padding: 20px;'>
            <div style='font-size: 3em; margin-bottom: 10px;'>📄 ➡️ 📝</div>
            <div>อัปโหลดไฟล์ PDF เพื่อเริ่มต้น</div>
        </div>
        """, 
        unsafe_allow_html=True
    )
