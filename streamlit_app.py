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
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

html, body, [data-testid="stApp"] {
    background: linear-gradient(135deg, #f4f7fb 0%, #eef2ff 100%);
}

/* PAGE WIDTH & TOP GAP */
.block-container {
    padding-top: 1rem;
    max-width: 850px;
}

/* HERO CARD — SIZE FIXED */
.hero-box {
    background: white;
    border-radius: 16px;
    padding: 1.3rem 1.8rem 1.2rem 1.8rem;  /* ⬅ reduced */
    box-shadow: 0 12px 26px rgba(0,0,0,0.07);
    text-align: center;
    margin-bottom: 1.6rem;
}

.hero-title {
    font-size: 1.7rem;   /* ⬅ reduced */
    font-weight: 600;
    color: #1a73e8;
    margin-bottom: 0.25rem;
}

.hero-subtitle {
    color: #5f6368;
    font-size: 0.95rem;  /* ⬅ reduced */
    line-height: 1.4;
}

/* DOWNLOAD BUTTON */
.stDownloadButton button {
    background: linear-gradient(135deg, #1a73e8, #4285f4);
    color: white;
    border-radius: 12px;
    padding: 0.75rem 1.9rem;
    font-size: 1.05rem;
    border: none;
}

.stDownloadButton button:hover {
    box-shadow: 0 8px 18px rgba(26,115,232,0.35);
}
</style>
""", unsafe_al_











