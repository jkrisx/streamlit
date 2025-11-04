import re, io, tempfile
from docx import Document
from openpyxl import load_workbook, Workbook
import streamlit as st

st.set_page_config(page_title="AL Converter - Streamlit", layout="wide")
st.title("Analisis Lendutan Converter — Streamlit")
st.markdown("Upload file .docx (bisa banyak), file database .db dan template Excel (.xlsx). Klik **Run** untuk menghasilkan Excel.")
