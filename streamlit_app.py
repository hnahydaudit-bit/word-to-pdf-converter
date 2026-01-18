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

# ------------------ CUSTOM CSS ------------------
st.markdown("""
<style>
body {
    background-color: #ffffff;
}

.main-card {
    background-color: #ffffff;
    padding: 2.5rem;
    border-radius: 16px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.05);
    max-width: 720px;
    margin: auto;
}

.title {
    font-size: 2rem;
    font-weight: 600;
    color: #1a73e8;
    text-align: center;
}

.subtitle {
    text-align: center;
    color: #5f6368;
    margin-bottom: 2rem;
}

.footer {
    text-align: center;
    font-size: 0.85rem;
    color: #80868b;
    margin-top: 2rem;
}

.stDownloadButton button {
    background-color: #1a73e8;
    color: white;
    border-radius: 8px;
    padding: 0.6rem 1.2rem;
    font-size: 1rem;
}

.stDownloadButton button:hover {
    background-color: #1558c0;
}
</style>
""", unsafe_allow_html=True)

# ------------------ UI CARD ------------------
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
        "⬇ Download ZIP",
        zip_buffer,
        file_name="converted_pdfs.zip",
        mime="application/zip"
    )

st.markdown(
    '<div class="footer">🔒 Files are processed securely and never stored.</div>',
    unsafe_allow_html=True
)

st.markdown('</div>', unsafe_allow_html=True)


