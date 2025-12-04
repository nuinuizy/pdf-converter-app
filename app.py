import streamlit as st
from pdf2docx import Converter
import os
import tempfile
from docx import Document
import re # พระเอกคนใหม่! เอาไว้สแกนหาคำผิดแบบละเอียด

st.set_page_config(page_title="PDF to Word Converter", page_icon="📄")

# --- ฟังก์ชันซ่อมภาษาไทย (Advanced Regex) ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        for para in doc.paragraphs:
            text = para.text
            
            # --- กฎการซ่อม 1: แก้สระอำแยกร่าง (เช่น "ก ำ" หรือ "ก  ำ" ให้เป็น "กำ") ---
            # ความหมาย: หาพยัญชนะไทย (ก-ฮ) ที่ตามด้วยช่องว่าง แล้วตามด้วย สระอำ
            text = re.sub(r'([ก-ฮ])\s+(ำ)', r'\1\2', text)

            # --- กฎการซ่อม 2: แก้สระอำที่มาแบบแยกชิ้น (นิคหิต + สระอา) ---
            # เช่น "ก" + "วงกลม" + "สระอา"
            text = re.sub(r'([ก-ฮ])\s*([ํ])\s*([า])', r'\1ำ', text)
            text = re.sub(r'([ํ])\s*([า])', r'ำ', text) # กรณีเหลือแค่เศษๆ

            # --- กฎการซ่อม 3: แก้สระบน/ล่าง ลอยห่างจากเพื่อน ---
            # เช่น "ที่" กลายเป็น "ท ี่" หรือ "ผู้" กลายเป็น "ผ ู้"
            # หาพยัญชนะ + ช่องว่าง + สระบนล่าง/วรรณยุกต์ -> จับมารวมกัน
            text = re.sub(r'([ก-ฮ])\s+([ัิีึืฺุู็่้๊๋์])', r'\1\2', text)
            
            # บันทึกข้อความที่ซ่อมแล้วกลับลงไป
            para.text = text
        
        doc.save(docx_path)
    except Exception as e:
        print(f"Repair skipped: {e}")

# --- ฟังก์ชันแปลงไฟล์ (เหมือนเดิม) ---
def convert_pdf_to_docx(uploaded_file):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
            
            # เรียกช่างซ่อมชุดใหญ่!
            repair_thai_docx(docx_path)
            
            with open(docx_path, "rb") as f:
                docx_data = f.read()
            return docx_data, docx_name
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            return None, None

# --- UI ---
st.title("📄 PDF to Word (Thai Super Fix 🔧)")
st.write("เวอร์ชันอัปเกรด: ซ่อมสระลอย สระแยกร่าง ด้วย Regex")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF", type="pdf")

if uploaded_file is not None:
    if st.button("🚀 แปลงและซ่อมไฟล์"):
        with st.spinner('กำลังแปลงร่าง... (ขั้นตอนนี้จะละเอียดหน่อยครับ)'):
            docx_data, docx_name = convert_pdf_to_docx(uploaded_file)
            
        if docx_data:
            st.success("✅ เสร็จแล้ว! ลองโหลดไปดูว่าหายไหมครับ")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Word",
                data=docx_data,
                file_name=docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
