import streamlit as st
from docx import Document
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
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
st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    html, body, [data-testid="stApp"] {
        background: linear-gradient(135deg, #f4f7fb 0%, #eef2ff 100%);
    }

    .block-container {
        padding-top: 0.6rem;
        max-width: 850px;
    }

    /* SMALL HERO CARD */
    .hero-box {
        background: white;
        border-radius: 10px;
        padding: 0.6rem 1.2rem;
        box-shadow: 0 8px 16px rgba(0,0,0,0.06);
        text-align: center;
        margin-bottom: 1rem;
    }

    .hero-title {
        font-size: 1.35rem;
        font-weight: 600;
        color: #1a73e8;
        margin: 0;
        line-height: 1.2;
    }

    .hero-subtitle {
        color: #5f6368;
        font-size: 0.82rem;
        margin-top: 0.2rem;
        line-height: 1.3;
    }

    .stDownloadButton button {
        background: linear-gradient(135deg, #1a73e8, #4285f4);
        color: white;
        border-radius: 10px;
        padding: 0.7rem 1.8rem;
        font-size: 1rem;
        border: none;
    }

    .stDownloadButton button:hover {
        box-shadow: 0 6px 14px rgba(26,115,232,0.35);
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------- HERO ----------------
st.markdown(
    """
    <div class="hero-box">
        <div class="hero-title">Word to PDF Converter</div>
        <div class="hero-subtitle">
            Convert Word documents into sequentially numbered PDFs
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# ---------------- UPLOAD ----------------
uploaded_files = st.file_uploader(
    "📤 Upload Word documents (.docx)",
    type=["docx"],
    accept_multiple_files=True
)

# ---------------- PROCESS ----------------
if uploaded_files:
    zip_buffer = io.BytesIO()
    styles = getSampleStyleSheet()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for idx, file in enumerate(uploaded_files, start=1):
            doc = Document(file)

            pdf_buffer = io.BytesIO()
            pdf = SimpleDocTemplate(
                pdf_buffer,
                pagesize=A4,
                rightMargin=40,
                leftMargin=40,
                topMargin=40,
                bottomMargin=40
            )

            elements = []

            for para in doc.paragraphs:
                if para.text.strip():
                    elements.append(Paragraph(para.text, styles["Normal"]))
                    elements.append(Spacer(1, 12))
                else:
                    elements.append(Spacer(1, 12))

            pdf.build(elements)
            pdf_buffer.seek(0)
            zipf.writestr(f"{idx:03}.pdf", pdf_buffer.read())

    zip_buffer.seek(0)
    timestamp = datetime.now().strftime("%d%m%Y_%H%M")

    st.success("✅ Your PDFs are ready")

    st.download_button(
        label="⬇️ Download ZIP",
        data=zip_buffer,
        file_name=f"Word_to_PDF_{timestamp}.zip",
        mime="application/zip"
    )













