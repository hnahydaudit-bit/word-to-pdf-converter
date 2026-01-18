import streamlit as st
import subprocess
import tempfile
import zipfile
from pathlib import Path

st.set_page_config(page_title="Word to PDF Converter", layout="centered")

st.title("📄 Word to PDF Converter")
st.write("Upload Word files and download PDFs as a ZIP.")

uploaded_files = st.file_uploader(
    "Upload .docx files",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # Save uploaded files
        for file in uploaded_files:
            (tmpdir / file.name).write_bytes(file.getbuffer())

        # Convert using LibreOffice
        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmpdir),
            ]
            + [str(tmpdir / f.name) for f in uploaded_files],
            check=True,
        )

        # Create ZIP
        zip_path = tmpdir / "converted_pdfs.zip"
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for i, pdf in enumerate(sorted(tmpdir.glob("*.pdf")), start=1):
                zipf.write(pdf, arcname=f"{i:03}.pdf")

        st.success("✅ Conversion successful!")

        st.download_button(
            label="⬇️ Download ZIP",
            data=zip_path.read_bytes(),
            file_name="word_to_pdf.zip",
            mime="application/zip",
        )
