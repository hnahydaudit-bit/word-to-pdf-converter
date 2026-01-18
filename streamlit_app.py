import streamlit as st
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
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

.block-container {
    padding-top: 1.5rem;
}

/* HERO CARD (MEDIUM SIZE) */
.hero-box {
    background: white;
    border-radius: 16px;
    padding: 1.8rem 1.8rem 1.6rem 1.8rem;
    box-shadow: 0 12px 28px rgba(0,0,0,0.08);
    text-align: center;
    margin-bottom: 1.8rem;
}

.hero-title {
    font-size: 1.8rem;
    font-weight: 600;
    color: #1a73e8;
    margin-bottom: 0.4rem;
}

.hero-subtitle {
    color: #5f6368;
    font-size: 0.95rem;
}

/* DOWNLOAD BUTTON */
.stDownloadButton button {
    background: linear-gradient(135deg, #1a73e8, #4285f4);
    color: white;
    border-radius: 12px;
    padding: 0.75rem 2rem;
    font-size: 1rem;
    border: none;
    transition: 0.2s;
}

.stDownloadButton button:hover {
    box-shadow: 0 8px 18px rgba(26,115,232,0.35);
    transform: translateY(-1px);
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

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for idx, file in enumerate(uploaded_files, start=1):
            doc = Document(file)

            pdf_buffer = io.BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=A4)
            width, height = A4
            y = height - 40

            for para in doc.paragraphs:
                if y < 40:
                    c.showPage()
                    y = height - 40
                c.drawString(40, y, para.text)
                y -= 14

            c.save()
            pdf_buffer.seek(0)

            zipf.writestr(f"{idx:03}.pdf", pdf_buffer.read())

    zip_buffer.seek(0)

    # --------- DYNAMIC ZIP FILE NAME ---------
    timestamp = datetime.now().strftime("%d%m%Y_%H%M")
    zip_filename = f"Word_to_PDF_{timestamp}.zip"

    st.success("✅ Your PDFs are ready")

    st.download_button(
        label="⬇️ Download ZIP",
        data=zip_buffer,
        file_name=zip_filename,
        mime="application/zip"
    )






