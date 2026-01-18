import streamlit as st
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import zipfile
import io

# ------------------ PAGE CONFIG ------------------
st.set_page_config(
    page_title="Word to PDF Converter",
    page_icon="📄",
    layout="centered"
)

# ------------------ GLOBAL CSS ------------------
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

html, body, [data-testid="stApp"] {
    background: linear-gradient(135deg, #f4f7fb 0%, #eef2ff 100%);
}

.block-container {
    padding-top: 1rem;
    padding-bottom: 2.5rem;
}

/* HERO SECTION */
.hero-box {
    background: linear-gradient(
        145deg,
        rgba(255,255,255,0.85),
        rgba(255,255,255,0.95)
    );
    backdrop-filter: blur(10px);
    border-radius: 22px;
    padding: 3rem 2.2rem 2.4rem 2.2rem;
    box-shadow: 
        0 20px 40px rgba(0,0,0,0.08),
        inset 0 1px 0 rgba(255,255,255,0.6);
    margin-bottom: 2.8rem;
}

.hero-title {
    font-size: 2.5rem;
    font-weight: 650;
    color: #1a73e8;
    text-align: center;
    margin-bottom: 0.6rem;
    letter-spacing: -0.3px;
}

.hero-subtitle {
    text-align: center;
    color: #5f6368;
    font-size: 1.08rem;
}

/* MAIN CARD */
.main-card {
    background: #ffffff;
    padding: 2.4rem 3rem;
    border-radius: 20px;
    box-shadow:
        0 16px 32px rgba(0,0,0,0.07);
    max-width: 780px;
    margin: auto;
}

/* DOWNLOAD BUTTON */
.stDownloadButton button {
    background: linear-gradient(135deg, #1a73e8, #4285f4);
    color: white;
    border-radius: 14px;
    padding: 0.85rem 2.2rem;
    font-size: 1.1rem;
    font-weight: 500;
    border: none;
    transition: all 0.2s ease-in-out;
}

.stDownloadButton button:hover {
    transform: translateY(-1px);
    box-shadow: 0 10px 22px rgba(26,115,232,0.35);
}
</style>
""", unsafe_allow_html=True)

# ------------------ HERO HEADER ------------------
st.markdown("""
<div class="hero-box">
    <div class="hero-title">Word to PDF Converter</div>
    <div class="hero-subtitle">
        Convert Word documents into sequentially numbered PDFs — instantly
    </div>
</div>
""", unsafe_allow_html=True)

# ------------------ MAIN CONTENT ------------------
st.markdown('<div class="main-card">', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "📤 Upload Word documents (.docx)",
    type=["docx"],
    accept_multiple_files=True
)

# ------------------ PROCESSING ------------------
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
    st.markdown("<br>", unsafe_allow_html=True)

    st.download_button(
        label="⬇️  Download ZIP",
        data=zip_buffer,
        file_name="converted_pdfs.zip",
        mime="application/zip"
    )

st.markdown('</div>', unsafe_allow_html=True)




