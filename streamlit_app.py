# app.py
# Improved Word (.docx) to PDF converter
# Preserves alignment, spacing, and paragraphs better

import streamlit as st
from docx import Document
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
import io
import zipfile

st.set_page_config(page_title="Word to PDF Converter", layout="centered")

st.title("Word to PDF Converter")
st.write("Convert Word documents into properly formatted PDFs")

uploaded_files = st.file_uploader(
    "Upload Word documents (.docx)",
    type=["docx"],
    accept_multiple_files=True
)


def convert_docx_to_pdf(docx_file):
    # Read Word file
    document = Document(docx_file)

    # Create PDF in memory
    pdf_buffer = io.BytesIO()
    pdf = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=11,
        leading=14,
        spaceAfter=8
    )

    story = []

    # Loop through Word paragraphs
    for para in document.paragraphs:
        text = para.text.strip()

        if not text:
            story.append(Spacer(1, 10))
        else:
            story.append(Paragraph(text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), normal_style))

    pdf.build(story)
    pdf_buffer.seek(0)

    return pdf_buffer


if uploaded_files:
    pdf_files = []

    for file in uploaded_files:
        pdf_buffer = convert_docx_to_pdf(file)
        pdf_files.append((file.name.replace(".docx", ".pdf"), pdf_buffer.getvalue()))

    # Create ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for name, data in pdf_files:
            zip_file.writestr(name, data)

    zip_buffer.seek(0)

    st.success("Your PDFs are ready")
    st.download_button(
        label="Download PDFs",
        data=zip_buffer,
        file_name="converted_pdfs.zip",
        mime="application/zip"
    )








