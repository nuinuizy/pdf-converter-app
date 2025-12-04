import streamlit as st
from pdf2docx import Converter
import os
import tempfile
from docx import Document

st.set_page_config(page_title="PDF to Word Converter", page_icon="📄")

# --- ฟังก์ชันซ่อม "สระ ำ" (The Sara Am Fixer) ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        
        # ฟังก์ชันย่อยสำหรับซ่อมข้อความใน Run
        def fix_text(text):
            if not text: return text
            # 1. แก้ " ำ" (เว้นวรรค + สระอำ) -> "ำ"
            text = text.replace(" ำ", "ำ")
            # 2. แก้ " ํ า" (นิคหิต + เว้นวรรค + สระอา) -> "ำ"
            text = text.replace("\u0e4d \u0e32", "\u0e33")
            # 3. แก้ " ํา" (นิคหิต + สระอา ติดกันแต่คนละตัว) -> "ำ" (ตัวเดียว)
            text = text.replace("\u0e4d\u0e32", "\u0e33")
            return text

        # 1. วนลูปแก้ในย่อหน้าปกติ (Paragraphs)
        for para in doc.paragraphs:
            for run in para.runs:
                if run.text:
                    run.text = fix_text(run.text)

        # 2. วนลูปแก้ในตาราง (Tables) - สำคัญมาก เพราะ pdf2docx ชอบสร้างตาราง
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.text:
                                run.text = fix_text(run.text)
                    
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
            # ใช้ pdf2docx แปลงไฟล์
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
            
            # เรียกช่างมาซ่อมสระ ำ
            repair_thai_docx(docx_path)
            
            with open(docx_path, "rb") as f:
                docx_data = f.read()
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            return None, None

# --- UI ---
st.title("📄 PDF to Word (Sara Am Fixed 🛠️)")
st.write("โหมดพิเศษ: เน้นแก้สระ ำ (Decomposed & Spaced Fix)")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    if st.button("🚀 แปลงไฟล์"):
        with st.spinner('กำลังแปลงร่างและซ่อมสระ...'):
            docx_data, docx_name = convert_pdf_to_docx(uploaded_file)
            
        if docx_data:
            st.success("✅ เรียบร้อย! สระ ำ น่าจะหายป่วยแล้วครับ")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word",
                data=docx_data,
                file_name=docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
