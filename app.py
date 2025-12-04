import streamlit as st
from pdf2docx import Converter
import os
import tempfile
from docx import Document

st.set_page_config(page_title="PDF to Word Converter", page_icon="📄")

# --- ฟังก์ชันซ่อมแบบ "ผ่าตัดเล็ก" (รักษา Format เดิม) ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        
        # วนลูปดูทุกย่อหน้า
        for para in doc.paragraphs:
            # เจาะดูทีละ "ก้อนข้อความ" (Run) ซึ่งเป็นหน่วยย่อยที่สุดที่เก็บ Format ไว้
            for run in para.runs:
                # ถ้าเจอ " ำ" (เว้นวรรค + สระอำ) ในก้อนนั้น
                if " ำ" in run.text:
                    # ให้ลบเว้นวรรคออก โดยไม่ไปยุ่งกับ Font หรือตัวหนา/เอียง
                    run.text = run.text.replace(" ำ", "ำ")
                
                # แถม: ถ้าเจอ สระอำ แยกร่างแบบ "นิคหิต" + "สระอา" ( ํ า )
                if "ํ า" in run.text:
                    run.text = run.text.replace("ํ า", "ำ")
                    
        doc.save(docx_path)
    except Exception as e:
        print(f"Repair skipped: {e}")

# --- ฟังก์ชันแปลงไฟล์ ---
def convert_pdf_to_docx(uploaded_file):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            # ใช้ pdf2docx ตามปกติ เพื่อคง Layout เดิมให้มากที่สุด
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
            
            # เรียกช่างมาซ่อมเฉพาะจุด
            repair_thai_docx(docx_path)
            
            with open(docx_path, "rb") as f:
                docx_data = f.read()
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            return None, None

# --- UI ---
st.title("📄 PDF to Word (Safe Mode 🛡️)")
st.write("โหมดปลอดภัย: รักษา Format เดิมไว้ 100% แก้เฉพาะสระอำ")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    if st.button("🚀 แปลงไฟล์"):
        with st.spinner('กำลังแปลงร่าง...'):
            docx_data, docx_name = convert_pdf_to_docx(uploaded_file)
            
        if docx_data:
            st.success("✅ เรียบร้อยครับ! Format เดิมน่าจะกลับมาแล้ว")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word",
                data=docx_data,
                file_name=docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
