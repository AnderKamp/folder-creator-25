import streamlit as st
import pandas as pd
from docx import Document
import io
import zipfile
import re
import unicodedata
from pathlib import PurePosixPath

FORBIDDEN = r"\\/:*?\"<>|"
_forbidden_pattern = re.compile(f"[{re.escape(FORBIDDEN)}]")

def sanitize_name(name: str) -> str:
    name = unicodedata.normalize("NFKC", name).strip()
    name = _forbidden_pattern.sub("-", name)
    return name or "Untitled"

def make_unique(name: str, existing: set) -> str:
    base = name
    i = 2
    while name in existing:
        name = f"{base} ({i})"
        i += 1
    existing.add(name)
    return name

def _write_empty_folder(z: zipfile.ZipFile, folder_path: str):
    p = PurePosixPath(folder_path)
    dir_path = str(p) if str(p).endswith("/") else f"{p}/"
    z.writestr(dir_path + ".keep", "")

def build_zip_from_names(names):
    out_buf = io.BytesIO()
    existing = set()
    with zipfile.ZipFile(out_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            if not n:
                continue
            folder_name = make_unique(n, existing)
            _write_empty_folder(zout, folder_name)
    return out_buf.getvalue()

st.set_page_config(page_title="Folder Creator 25", page_icon="📁")
st.title("📁 Folder Creator 25")
tab1, tab2, tab3 = st.tabs(["From Filenames (.zip)", "From Word (.docx)", "From Spreadsheet (.csv/.xlsx)"])

with tab1:
    uploaded_zip = st.file_uploader("Upload a ZIP of files", type=["zip"])
    if uploaded_zip and st.button("Create Folders from Filenames"):
        with zipfile.ZipFile(uploaded_zip) as zin:
            names = [sanitize_name(PurePosixPath(f).stem) for f in zin.namelist() if not f.endswith("/")]
        out_bytes = build_zip_from_names(names)
        st.download_button("Download Folders.zip", out_bytes, "Folders.zip", "application/zip")

with tab2:
    uploaded_docx = st.file_uploader("Upload a DOCX", type=["docx"])
    if uploaded_docx and st.button("Create Folders from Word"):
        doc = Document(uploaded_docx)
        names = [sanitize_name(p.text) for p in doc.paragraphs if p.text.strip()]
        out_bytes = build_zip_from_names(names)
        st.download_button("Download Folders.zip", out_bytes, "Folders.zip", "application/zip")

with tab3:
    uploaded_sheet = st.file_uploader("Upload a CSV or XLSX", type=["csv", "xlsx"])
    if uploaded_sheet:
        try:
            if uploaded_sheet.name.endswith(".csv"):
                df = pd.read_csv(uploaded_sheet, dtype=str)
            else:
                df = pd.read_excel(uploaded_sheet, dtype=str)
            col = st.selectbox("Select column", df.columns)
            if st.button("Create Folders from Spreadsheet"):
                names = [sanitize_name(v) for v in df[col].dropna().astype(str) if v.strip()]
                out_bytes = build_zip_from_names(names)
                st.download_button("Download Folders.zip", out_bytes, "Folders.zip", "application/zip")
        except Exception as e:
            st.error(f"Error reading spreadsheet: {e}")
