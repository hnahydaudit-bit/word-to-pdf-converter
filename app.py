import streamlit as st
import subprocess
import tempfile
import zipfile
from pathlib import Path

st.set_page_config(page_title="Word to PDF Converter")

st.title("📄 Word to PDF Converter with Auto Numbering")
st.write("Upload Word documents (.docx). They will be converted to sequentially numbered PDFs.")

uploaded_files = st.file_uploader(
    "Upload Word files",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Convert to PDF"):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_dir = Path(tmpdir) / "input"
            output_dir = Path(tmpdir) / "output"

            input_dir.mkdir()
            output_dir.mkdir()

            # Save uploaded Word files
            for file in uploaded_files:
                with open(input_dir / file.name, "wb") as f:
                    f.write(file.read())

            # Convert Word → PDF using LibreOffice
            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(output_dir),
                    str(input_dir / "*.docx")
                ],
                check=True
            )

            # Create ZIP with sequential numbering
            zip_path = Path(tmpdir) / "numbered_pdfs.zip"
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for i, pdf in enumerate(sorted(output_dir.glob("*.pdf")), start=1):
                    zipf.write(pdf, f"{i:03}.pdf")

            # Download button
            with open(zip_path, "rb") as f:
                st.download_button(
                    "⬇ Download PDFs (ZIP)",
                    f,
                    file_name="numbered_pdfs.zip"
                )
