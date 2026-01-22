import streamlit as st
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import pikepdf
from docx2pdf import convert
import tempfile
import os

st.set_page_config(page_title="PDF Utility Tool", layout="centered")

st.title("📄 PDF Utility Tool")
st.caption("All processing happens locally on your system")

option = st.selectbox(
    "Choose an option",
    [
        "Merge PDFs",
        "Split PDF",
        "Compress PDF",
        "Convert Word to PDF"
    ]
)

# ---------------- MERGE PDFs ----------------
if option == "Merge PDFs":
    st.subheader("Merge Multiple PDFs")

    uploaded_files = st.file_uploader(
        "Upload PDF files",
        type="pdf",
        accept_multiple_files=True
    )

    if st.button("Merge PDFs") and uploaded_files:
        merger = PdfMerger()

        for pdf in uploaded_files:
            merger.append(pdf)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            merger.write(tmp.name)
            merger.close()

            with open(tmp.name, "rb") as f:
                st.download_button(
                    "Download Merged PDF",
                    f,
                    file_name="merged.pdf",
                    mime="application/pdf"
                )

# ---------------- SPLIT PDF ----------------
elif option == "Split PDF":
    st.subheader("Split PDF by Page Range")

    uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

    if uploaded_file:
        start = st.number_input("Start page", min_value=1, value=1)
        end = st.number_input("End page", min_value=1, value=1)

        if st.button("Split PDF"):
            reader = PdfReader(uploaded_file)
            writer = PdfWriter()

            for i in range(int(start) - 1, int(end)):
                writer.add_page(reader.pages[i])

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                writer.write(tmp.name)

                with open(tmp.name, "rb") as f:
                    st.download_button(
                        "Download Split PDF",
                        f,
                        file_name="split.pdf",
                        mime="application/pdf"
                    )

# ---------------- COMPRESS PDF ----------------
elif option == "Compress PDF":
    st.subheader("Compress PDF (Best for text-only PDFs)")

    uploaded_file = st.file_uploader("Upload PDF", type="pdf")

    if uploaded_file and st.button("Compress PDF"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as input_tmp:
            input_tmp.write(uploaded_file.read())

        output_path = input_tmp.name.replace(".pdf", "_compressed.pdf")

        with pikepdf.open(input_tmp.name) as pdf:
            pdf.save(output_path, optimize_streams=True)

        with open(output_path, "rb") as f:
            st.download_button(
                "Download Compressed PDF",
                f,
                file_name="compressed.pdf",
                mime="application/pdf"
            )

# ---------------- WORD TO PDF ----------------
elif option == "Convert Word to PDF":
    st.subheader("Convert Word to PDF")

    uploaded_file = st.file_uploader("Upload Word file", type=["docx"])

    if uploaded_file and st.button("Convert"):
        with tempfile.TemporaryDirectory() as tmpdir:
            word_path = os.path.join(tmpdir, uploaded_file.name)
            pdf_path = word_path.replace(".docx", ".pdf")

            with open(word_path, "wb") as f:
                f.write(uploaded_file.read())

            convert(word_path, pdf_path)

            with open(pdf_path, "rb") as f:
                st.download_button(
                    "Download PDF",
                    f,
                    file_name="converted.pdf",
                    mime="application/pdf"
                )
















