import streamlit as st
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import zipfile
import io

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
    padding-top: 2rem;
}

/* HERO CARD */
.hero-box {
    background: white;
    border-radius: 24px;
    padding: 3rem 2.5rem 2.5rem 2.5rem;
    box-shadow: 0 20px 40px rgba(0,0,0,0.08);
    text-align: center;
    margin-bottom: 2.5rem;
}

.hero-title {
    font-size: 2.4rem;
    font-weight: 650;
    color: #1a73e8;
    margin-bottom: 0.6rem;
}

.hero-subtitle {
    color: #5f6368;
    font-size: 1.05rem;
}

/* DOWNLOAD BUTTON */
.stDownloadButton button {
    background: linear-gradient(135deg, #1a73e8, #4285f4);
    color: white;
    border-radius: 14px;
    padding: 0.85rem 2.2rem;
    font-size: 1.1rem;
    border: none;
    transition: 0.2s;
}

.stDownloadButton button:hover {
    box-shadow: 0 10px 22px rgba(26,115,232,0.35);
    transform: translateY(-1px);
}
</style>
""", unsafe_allow_html=True)

# ---------------- HERO ----------------
st.markdown("""
<div class="hero-box">
    <div class="hero-title">Word to PDF Converter</div>
    <div class="hero-subtitle">
        Convert Word documents into sequentially numbered PDFs — instantly
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

    st.success("✅ Your PDFs are ready")

    st.download_button(
        label="⬇️ Download ZIP",
        data=zip_buffer,
        file_name="converted_pdfs.zip",
        mime="application/zip"
    )





