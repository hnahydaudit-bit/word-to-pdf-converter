import streamlit as st
from docx import Document
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import zipfile
import io
from datetime import datetime

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Word to PDF Converter",
    page_icon="📄",
    layout="centered"
)

# ---------------- CSS ----------------
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

html, body, [data-testid="stApp"] {
    background-color: #f5f7fb;
}

/* MAIN CONTAINER */
.block-container {
    padding-top: 3.5rem;
    max-width: 720px;
}

/* HEADER */
.app-header {
    text-align: center;
    margin-bottom: 1.8rem;
}

.app-title {
    font-size: 2rem;
    font-weight: 600;
    color: #1a73e8;
    margin-bottom: 0.3rem;
    line-height: 1.25;
}

.app-subtitle {
    font-size: 0.95rem;
    color: #5f6368;
}

/* SUCCESS MESSAGE */
.stAlert {
    border-radius: 10px;
}

/* DOWNLOAD BUTTON */
.stDownloadButton button {
    background: #1a73e8;
    color: white;
    border-radius: 10px;
    padding: 0.7rem 1.8rem;
    font-size: 0.95rem;
    border: none;
}

.stDownloadButton button:hover {
    background: #1558c0;
}
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------
st.markdown("""
<div class="app-header">
    <div class="app-title">Word to PDF Converter</div>
    <div class="app-subtitle">
        Convert Word documents into sequentially numbered PDFs
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------- FILE UPLOADER ----------------
uploaded_files = st.file_uploader(
    "Upload Word documents (.docx)",
    type=["docx"],
    accept_multiple_files=True
)

# ---------------- PROCESS ----------------
if uploaded_files:
    zip_buffer = io.BytesIO()
    styles = getSampleStyleSheet()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for idx, file in enumerate(uploaded_files, start=1):
            doc = Document(file)

            pdf_buffer = io.BytesIO()
            pdf = SimpleDocTemplate(
                pdf_buffer,
                pagesize=A4,
                rightMargin=40,
                leftMargin=40,
                topMargin=40,
                bottomMargin=40
            )

            elements = []

            for para in doc.paragraphs:
                if para.text.strip():
                    elements.append(Paragraph(para.text, styles["Normal"]))
                    elements.append(Spacer(1, 12))
                else:
                    elements.append(Spacer(1, 12))

            pdf.build(elements)
            pdf_buffer.seek(0)
            zipf.writestr(f"{idx:03}.pdf", pdf_buffer.read())

    zip_buffer.seek(0)

    timestamp = datetime.now().strftime("%d%m%Y_%H%M")
    zip_filename = f"Word_to_PDF_{timestamp}.zip"

    st.success("Your PDFs are ready")

    st.download_button(
        label="Download ZIP",
        data=zip_buffer,
        file_name=zip_filename,
        mime="application/zip"
    )














