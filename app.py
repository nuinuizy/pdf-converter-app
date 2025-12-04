import streamlit as st
from pdf2docx import Converter
import os
import tempfile
from docx import Document

st.set_page_config(page_title="PDF to Word Converter", page_icon="📄")

# --- ฟังก์ชันซ่อม "สระ ำ" (เน้นแก้ Space หน้า ำ) ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        
        # ฟังก์ชันย่อย: จัดการ " ำ" ตัวปัญหา
        def fix_sara_am(text):
            if not text: return text
            
            # --- จุดแก้ไขหลักตามที่คุณขอ ---
            # ถ้าเจอ " ำ" (เคาะวรรค + สระอำเต็มรูป) ให้เปลี่ยนเป็น "ำ" (ชิดตัวหน้า)
            # ทำซ้ำ 2 รอบ เผื่อเจอเคาะวรรคเบิ้ลมา (เช่น "  ำ")
            text = text.replace(" ำ", "ำ") 
            text = text.replace(" ำ", "ำ")
            
            return text

        # 1. วนลูปแก้ในย่อหน้า (Paragraphs)
        for para in doc.paragraphs:
            for run in para.runs:
                if run.text:
                    run.text = fix_sara_am(run.text)

        # 2. วนลูปแก้ในตาราง (Tables) - สำคัญมาก เอกสารราชการ/บัญชีมักอยู่ในนี้
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.text:
                                run.text = fix_sara_am(run.text)
                    
        doc.save(docx_path)
    except Exception as e:
        st.error(f"Repair Error: {e}")

# --- ฟังก์ชันแปลงไฟล์ ---
def convert_pdf_to_docx(uploaded_file):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            # ใช้ pdf2docx แปลงไฟล์
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
            
            # เรียกช่างมาซ่อม " ำ"
            repair_thai_docx(docx_path)
            
            with open(docx_path, "rb") as f:
                docx_data = f.read()
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            return None, None

# --- UI ---
st.title("📄 PDF to Word (รองรับภาษาไทย)")
st.caption("็Happy Everyday")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    if st.button("🚀 แปลงและซ่อม"):
        with st.spinner('กำลังจัดการเจ้าสระอำ...'):
            docx_data, docx_name = convert_pdf_to_docx(uploaded_file)
            
        if docx_data:
            st.success("✅ เสร็จเรียบร้อย! Have a good day")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word",
                data=docx_data,
                file_name=docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

