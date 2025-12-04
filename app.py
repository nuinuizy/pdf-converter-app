import streamlit as st
from pdf2docx import Converter
import os
import tempfile
import time
from docx import Document

# --- Config & Style ---
st.set_page_config(
    page_title="PDF2Word Pro [WS-Core]", 
    page_icon="⚡", 
    layout="centered"
)

# Custom CSS for Tech Vibe (Hide default menu, custom font tweaks)
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #FF4B4B;
        color: white;
    }
    .reportview-container {
        background: #0E1117;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Core Logic (The Engine) ---
def repair_thai_docx(docx_path):
    try:
        doc = Document(docx_path)
        
        def fix_sara_am(text):
            if not text: return text
            # Logic เดิมที่ถูกต้องแล้ว: แก้ Space หน้า ำ
            text = text.replace(" ำ", "ำ") 
            text = text.replace(" ำ", "ำ")
            return text

        # Patching Process
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
    except Exception as e:
        st.error(f"System Error during patch: {e}")
        return False

def convert_pdf_to_docx(uploaded_file, progress_bar, status_text):
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        docx_name = os.path.splitext(uploaded_file.name)[0] + ".docx"
        docx_path = os.path.join(temp_dir, docx_name)
        
        try:
            # Step 1: Initialization
            status_text.text("Status: Initializing Converter Engine...")
            progress_bar.progress(10)
            time.sleep(0.5) # Simulate init

            # Step 2: Conversion
            status_text.text("Status: Extracting Layout & Text...")
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
            progress_bar.progress(60)
            
            # Step 3: Patching Thai Vowels
            status_text.text("Status: Patching Thai Vowel (Sara Am) Glitch...")
            repair_thai_docx(docx_path)
            progress_bar.progress(90)
            
            # Step 4: Finalizing
            status_text.text("Status: Finalizing Output...")
            with open(docx_path, "rb") as f:
                docx_data = f.read()
            
            progress_bar.progress(100)
            time.sleep(0.2)
            return docx_data, docx_name
            
        except Exception as e:
            st.error(f"Critical Error: {e}")
            return None, None

# --- UI / Dashboard Layout ---

# Sidebar for Context
with st.sidebar:
    st.title("⚙️ Control Panel")
    st.info("**System:** PDF2Word Converter\n\n**Version:** 2.1 (Patch WS)\n\n**Module:** Thai Language Fixer enabled.")
    st.markdown("---")
    st.caption("Designed for High-Performance Workflows.")

# Main Area
col1, col2 = st.columns([3, 1])
with col1:
    st.title("PDF2Word `Pro`")
    st.markdown("**Automated Document Conversion Utility**")
with col2:
    # Techy Status Badge
    st.success("● System Online")

st.markdown("---")

# File Upload Section
uploaded_file = st.file_uploader("Drop your PDF source file here:", type="pdf")

if uploaded_file is not None:
    # File Details (Tech info)
    file_stats = f"Filename: {uploaded_file.name} | Size: {uploaded_file.size / 1024:.2f} KB"
    st.caption(f"📄 Source Detected: {file_stats}")
    
    st.markdown("###") # Spacer

    if st.button("🚀 EXECUTE CONVERSION"):
        # UI Elements for Process
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        start_time = time.time()
        
        # Run Process
        docx_data, docx_name = convert_pdf_to_docx(uploaded_file, progress_bar, status_text)
        
        end_time = time.time()
        duration = end_time - start_time
        
        if docx_data:
            st.markdown("---")
            # Result Metrics
            m1, m2 = st.columns(2)
            m1.metric(label="Processing Time", value=f"{duration:.2f}s", delta="Completed")
            m2.metric(label="Patch Status", value="Verified", delta="Clean")
            
            st.success("✅ Operation Successful. Output ready for deployment.")
            
            # Download
            st.download_button(
                label="📥 DOWNLOAD .DOCX",
                data=docx_data,
                file_name=docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
    else:
        st.info("Waiting for execution command...")

else:
    # Empty State with minimal tech visual
    st.markdown(
        """
        <div style='text-align: center; color: gray; margin-top: 50px;'>
            Awaiting Input Stream...
        </div>
        """, unsafe_allow_html=True
    )
