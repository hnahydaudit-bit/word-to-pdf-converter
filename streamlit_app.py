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

# ------------------ REMOVE DEFAULT STREAMLIT SPACING ------------------
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
}
</style>
""", unsafe_allow_html=True)

# ------------------ CUSTOM GOOGLE-STYLE UI ------------------
st.markdown("""
<style>
.main-card {
    background-color: #ffffff;
    padding: 2.6rem 3.2rem;
    border-radius: 18px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.06);
    max-width: 760px;
    margin: auto;
}

.title {
    font-size: 2.2rem;
    font-weight: 600;
    color: #1a73e8;
    text-align: center;
    margin-bottom: 0.4rem;
}

.subtitle {
    text-align: center;
    color: #5f6368;
    font-size: 1rem;
    margin-bottom: 2.2rem;
}

.stDownloadButton button {
    background-color: #1a73e8;
    color: white;
    border-radius: 12px;
    padding: 0.75rem 1.8rem;
    font-size: 1.05rem;
    font-weight: 500;
    border: none;
}

.stDownloadButton button:hover {
    background-color: #1558c0;
}
</style>
""", unsafe_allow_html=True)

# ------------------ MAIN UI CARD ------------------
st.markdown('<div class="main-card">', unsafe_allow_html=True)

st.markdown('<div class="title">Word to PDF Converter</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Convert Word documents into sequentially numbered PDFs — instantly</div>',
    unsafe_allow_html=True
)

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



