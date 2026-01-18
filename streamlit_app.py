import streamlit as st
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import zipfile
import io

st.set_page_config(page_title="Word to PDF Converter", layout="centered")

st.title("📄 Word to PDF Converter")
st.caption("Upload Word files → Get sequentially numbered PDFs")

uploaded_files = st.file_uploader(
    "Upload .docx files",
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

            pdf_name = f"{idx:03}.pdf"
            zipf.writestr(pdf_name, pdf_buffer.read())

    zip_buffer.seek(0)

    st.success("✅ Conversion complete")
    st.download_button(
        "⬇ Download ZIP",
        zip_buffer,
        file_name="converted_pdfs.zip",
        mime="application/zip"
    )

