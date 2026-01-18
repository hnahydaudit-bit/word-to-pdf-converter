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
    background: linear-gradient(135deg, #f4f7fb 0%, #eef2ff 100%);
}

/* FIXED TOP SPACING */
.block-container {
    padding-top: 1.2rem;
    max-width: 850px;
}

/* HERO CARD (POSITION FIXED) */
.hero-box {
    background: white;
    border-radius: 20px;
    padding: 2rem 2rem 1.8rem 2rem;
    box-shadow: 0 15px 30px rgba(0,0,0,0.08);
    text-align: center;
    margin-top: 0;
    margin-bottom: 2rem;
}

.hero-title {
    font-size: 2rem;
    font-weight: 600;
    color: #1a73e8;
    margin-bottom: 0.4rem;
}

.hero-subtitle {
    color: #5f6368;
    font-size: 1rem;
}

/* DOWNLOAD BUTTON */
.stDownloadButton button {
    background: linear-gradient(135deg, #1a73e8, #4285f4);
    color: white;
    border-radius: 12px;
    padding: 0.8rem 2rem;
    font-size: 1.05rem;
    border: none;
}

.stDownloadButton button:hover {
    box-shadow: 0 8px 18px rgba(26,115,232,0.35);
}
</style>
""", unsafe_allow_html=True)

# ---------------- HERO ----------------
st.markdown("""
<div class="hero-box">
    <div class="hero-title">Word to PDF Converter</div>
    <div class="hero-subtitle">
        Convert Word documents into sequentially numbered PDFs
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------- UPLOAD ----------------
uploaded_files = st.file_uploader(
    "📤 Upload Word documents (.docx)",
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

    st.success("✅ Your PDFs are ready")

    st.download_button(
        label="⬇️ Download ZIP",
        data=zip_buffer,
        file_name=zip_filename,
        mime="application/zip"
    )










