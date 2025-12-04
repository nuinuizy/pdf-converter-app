import streamlit as st
from pdf2docx import Converter
import os
import tempfile

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="PDF to Word Converter", page_icon="📄")

def convert_pdf_to_docx(uploaded_file):
    # สร้าง Temporary directory เพื่อพักไฟล์
    with tempfile.TemporaryDirectory() as temp_dir:
        # 1. Save ไฟล์ PDF ที่อัพโหลดมาลงเครื่องชั่วคราว
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # 2. ตั้งชื่อไฟล์ปลายทาง (.docx)
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        # 3. เริ่มกระบวนการแปลง (Converter Logic)
        try:
            cv = Converter(pdf_path)
            # convert(docx_filename, start=0, end=None) แปลงทุกหน้า
            cv.convert(docx_path) 
            cv.close()
            
            # 4. อ่านไฟล์ Word ที่ได้เพื่อส่งกลับไปให้ User
            with open(docx_path, "rb") as f:
                docx_data = f.read()
            
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการแปลงไฟล์: {e}")
            return None, None

# --- UI Section ---
st.title("📄 PDF to Word Converter (Thai Supported)")
st.write("อัพโหลดไฟล์ PDF ของคุณ แล้วระบบจะแปลงเป็น Microsoft Word ให้ทันที")

# Widget สำหรับอัพโหลด
uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    # ปุ่มกดเพื่อเริ่มแปลง
    if st.button("🚀 แปลงเป็น Word"):
        with st.spinner('กำลังแปลงร่าง... รอแป๊บนะครับ (ไฟล์ใหญ่จะนานหน่อย)'):
            docx_data, docx_name = convert_pdf_to_docx(uploaded_file)
            
        if docx_data:
            st.success("✅ เรียบร้อย! ดาวน์โหลดได้เลยครับ")
            
            # ปุ่มดาวน์โหลด
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word",
                data=docx_data,
                file_name=docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# Footer หล่อๆ
st.markdown("---")
st.caption("Note: ภาษาไทยจะสมบูรณ์ 90-100% ขึ้นอยู่กับการฝัง Font ในต้นฉบับ PDF นะครับ")