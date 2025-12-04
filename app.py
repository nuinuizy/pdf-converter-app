import streamlit as st
from pdf2docx import Converter
import os
import tempfile
from docx import Document # พระเอกคนใหม่ มาช่วยซ่อมไฟล์

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="PDF to Word Converter", page_icon="📄")

# --- ฟังก์ชันซ่อมภาษาไทย (The Fixer) ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        # วนลูปเช็คทุกย่อหน้าในไฟล์
        for para in doc.paragraphs:
            # สูตรแก้คำผิดยอดฮิต
            if " ำ" in para.text:
                para.text = para.text.replace(" ำ", "ำ") # ลบช่องว่างหน้าสระอำ
            if "ํ" in para.text and "า" in para.text:
                para.text = para.text.replace("ํ า", "ำ").replace("ํ", "") # ถ้านิคหิตแยกกับสระอา ให้รวมร่าง
        
        doc.save(docx_path)
    except Exception as e:
        print(f"Repair skipped: {e}") 
        # ถ้าซ่อมไม่ได้ (เช่น Font ไม่รองรับ) ก็ปล่อยผ่านไป ดีกว่าโปรแกรมพัง

# --- ฟังก์ชันแปลงไฟล์หลัก ---
def convert_pdf_to_docx(uploaded_file):
    with tempfile.TemporaryDirectory() as temp_dir:
        # 1. Save PDF
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # 2. แปลงเป็น Word
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path) 
            cv.close()
            
            # 🔥 3. เรียกช่างมาซ่อมภาษาไทยก่อนส่ง!
            repair_thai_docx(docx_path)
            
            # 4. อ่านไฟล์ที่ซ่อมเสร็จแล้ว
            with open(docx_path, "rb") as f:
                docx_data = f.read()
            
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            return None, None

# --- UI หน้าเว็บ ---
st.title("📄 PDF to Word Converter (Thai Repair Ver.)")
st.write("อัพโหลด PDF มาเลย เดี๋ยวแปลง + ซ่อมสระลอยให้ด้วย!")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    if st.button("🚀 แปลงและซ่อมไฟล์"):
        with st.spinner('กำลังแปลงร่างและซ่อมสระ... ใจเย็นๆ นะครับ'):
            docx_data, docx_name = convert_pdf_to_docx(uploaded_file)
            
        if docx_data:
            st.success("✅ เรียบร้อย! ซ่อมสระอำให้แล้ว ลองโหลดดูครับ")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word",
                data=docx_data,
                file_name=docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

st.markdown("---")
st.caption("Tips: ถ้าสระยังเพี้ยนอยู่ แนะนำให้เปิดไฟล์ Word แล้วกด Ctrl+H เพื่อค้นหาและแทนที่คำผิดดูนะครับ")
