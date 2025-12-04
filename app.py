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
        /* แต่ง Progress Bar ให้ดูดี */
        .stProgress > div > div > div > div {
            background-color: #FF4B4B;
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

def convert_pdf_to_docx(uploaded_file, status_box, progress_bar, turbo_mode):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f: f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            # 1. เริ่มต้นเครื่องยนต์
            status_box.info("⚙️ กำลังวิเคราะห์ไฟล์ PDF... (Initializing)")
            progress_bar.progress(10)
            
            cv = Converter(pdf_path)
            
            # 2. นับจำนวนหน้า (ลูกเล่นใหม่!)
            num_pages = len(cv.pages)
            if num_pages > 10:
                status_box.warning(f"⚠️ เจอไฟล์ใหญ่ {num_pages} หน้า! อาจใช้เวลา 2-5 นาที... กรุณารอห้ามปิดจอนะครับ")
            else:
                status_box.info(f"📄 เจอทั้งหมด {num_pages} หน้า กำลังเริ่มแปลงร่าง...")
            
            progress_bar.progress(20)
            
            # 3. แปลงไฟล์ (The Heavy Lifting)
            if turbo_mode:
                # ตัดรูปทิ้ง เร็วขึ้น
                cv.convert(docx_path, multi_processing=False, parse_images=False)
            else:
                # เอาครบ ช้าหน่อย
                cv.convert(docx_path, multi_processing=False)
            
            cv.close()
            
            # 4. ซ่อมสระ
            progress_bar.progress(80)
            status_box.info("🔧 กำลังซ่อมสระภาษาไทย (Fixing Thai Vowels)...")
            repair_thai_docx(docx_path)
            
            progress_bar.progress(100)
            
            with open(docx_path, "rb") as f: docx_data = f.read()
            return docx_data, docx_name, num_pages
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            return None, None, 0

# --- 3. UI (v2.6 Progress Bar) ---

c1, c2 = st.columns([3, 1])
c1.markdown("### ⚡ PDF to Word `Pro`")
c2.markdown("<div style='text-align: right; color: gray; font-size: 0.8em; padding-top: 10px;'>v2.6 Progress</div>", unsafe_allow_html=True)

st.divider()

uploaded_file = st.file_uploader("Upload PDF file", type="pdf", label_visibility="collapsed")

if uploaded_file:
    # Checkbox
    turbo = st.checkbox("⚡ Turbo Mode (ตัดรูปภาพออก = เร็วขึ้น 3 เท่า)", value=True)
    
    run_btn = st.button("🚀 เริ่มแปลงไฟล์ (Start)")
    
    # จองพื้นที่สำหรับ Bar และ Status
    status_box = st.empty()
    progress_bar = st.empty()

    if run_btn:
        start_time = time.time()
        
        # ใส่ progress_bar ลงไปในฟังก์ชันด้วย
        docx_data, docx_name, pages = convert_pdf_to_docx(uploaded_file, status_box, progress_bar, turbo)
        
        duration = time.time() - start_time
        
        if docx_data:
            status_box.success(f"✅ เสร็จเรียบร้อย! (แปลงไปทั้งหมด {pages} หน้า)")
            st.divider()
            
            c_info, c_btn = st.columns([1.5, 2])
            with c_info:
                st.caption(f"⏱️ เวลาที่ใช้: {duration:.2f}s")
                st.caption(f"📦 ขนาดไฟล์: {len(docx_data)/1024:.1f} KB")
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
            <div style='font-size: 3em; margin-bottom: 10px;'>📄 📊 📝</div>
            <div>อัปโหลดไฟล์ PDF เพื่อเริ่มงาน</div>
            <div style='font-size: 0.8em; color: #999;'>(มี Progress Bar บอกสถานะแล้วนะ)</div>
        </div>
        """, 
        unsafe_allow_html=True
    )
