import io
import re
import zipfile
import unicodedata
from pathlib import PurePosixPath

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="Folder Creator 25", page_icon="📁", layout="centered")
st.title("📁 Folder Creator 25")

# ---------- Helpers ----------
FORBIDDEN = r"\\/:*?\"<>|"
SYSTEM_DIR_PREFIXES = ("__MACOSX/",)
SYSTEM_BASENAMES = {".DS_Store", "Thumbs.db"}
_forbidden_pattern = re.compile(f"[{re.escape(FORBIDDEN)}]|[\x00-\x1f]")

def sanitize_name(name: str) -> str:
    name = unicodedata.normalize("NFKC", str(name)).strip()
    name = _forbidden_pattern.sub("-", name)
    name = name.rstrip(" .")
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

def build_zip_from_filenames_zip(uploaded_zip, include_files: bool, do_sanitize: bool):
    out_buf = io.BytesIO()
    existing = set()
    uploaded_zip.seek(0)
    with zipfile.ZipFile(out_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zout:
        with zipfile.ZipFile(uploaded_zip) as zin:
            for info in zin.infolist():
                if info.is_dir():
                    continue
                if any(info.filename.startswith(pref) for pref in SYSTEM_DIR_PREFIXES):
                    continue
                basename = PurePosixPath(info.filename).name
                if basename in SYSTEM_BASENAMES:
                    continue
                stem = PurePosixPath(basename).stem
                name = sanitize_name(stem) if do_sanitize else (stem.strip() or "Untitled")
                folder_name = make_unique(name, existing)
                _write_empty_folder(zout, folder_name)
                if include_files:
                    data = zin.read(info.filename)
                    dest = str(PurePosixPath(folder_name) / basename)
                    zout.writestr(dest, data)
    return out_buf.getvalue()

def build_zip_from_uploaded_files(files, include_files: bool, do_sanitize: bool):
    out_buf = io.BytesIO()
    existing = set()
    with zipfile.ZipFile(out_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zout:
        for f in files:
            basename = PurePosixPath(f.name).name
            stem = PurePosixPath(basename).stem
            name = sanitize_name(stem) if do_sanitize else (stem.strip() or "Untitled")
            folder_name = make_unique(name, existing)
            _write_empty_folder(zout, folder_name)
            if include_files:
                data = f.read()
                f.seek(0)  # reset for safety
                dest = str(PurePosixPath(folder_name) / basename)
                zout.writestr(dest, data)
    return out_buf.getvalue()

# ---------- UI ----------
tab1, tab2, tab3 = st.tabs(["From Filenames", "From Word (.docx)", "From Spreadsheet (.csv/.xlsx)"])

# --- Tab 1: Filenames ---
with tab1:
    mode = st.radio("Choose how to provide files:", ["Select multiple files", "Upload a ZIP of files"], horizontal=True)
    col1, col2 = st.columns(2)
    with col1:
        opt_sanitize = st.checkbox("Sanitize names", value=True)
    with col2:
        opt_include = st.checkbox("Include original files in each folder", value=True)

    if mode == "Select multiple files":
        files = st.file_uploader("Select files (you can choose many)", accept_multiple_files=True, type=None, key="multi_files")
        st.caption("Tip: You can drag & drop many files here at once.")
        if files and st.button("Create Folders (From Selected Files)", type="primary"):
            out_bytes = build_zip_from_uploaded_files(files, opt_include, opt_sanitize)
            st.success("Folders ZIP created.")
            st.download_button("⬇️ Download Folders.zip", out_bytes, "Folders.zip", "application/zip")
    else:
        uploaded_zip = st.file_uploader("Upload a ZIP of files", type=["zip"], key="zip_upl")
        if uploaded_zip and st.button("Create Folders (From ZIP)", type="primary"):
            out_bytes = build_zip_from_filenames_zip(uploaded_zip, opt_include, opt_sanitize)
            st.success("Folders ZIP created.")
            st.download_button("⬇️ Download Folders.zip", out_bytes, "Folders.zip", "application/zip")

# --- Tab 2: Word (.docx) ---
with tab2:
    uploaded_docx = st.file_uploader("Upload DOCX (one folder name per line)", type=["docx"], key="docx_upl")
    opt_sanitize_w = st.checkbox("Sanitize names (Word)", value=True, key="san_w")
    if uploaded_docx is not None and st.button("Create Folders (From Word)"):
        doc = Document(uploaded_docx)
        names = [sanitize_name(p.text) if opt_sanitize_w else (p.text.strip() or "Untitled")
                 for p in doc.paragraphs if p.text.strip()]
        out_bytes = build_zip_from_names(names)
        st.success("Folders ZIP created.")
        st.download_button("⬇️ Download Folders.zip", out_bytes, "Folders.zip", "application/zip")

# --- Tab 3: Spreadsheet (.csv/.xlsx) ---
with tab3:
    uploaded_sheet = st.file_uploader("Upload CSV or XLSX", type=["csv", "xlsx"], key="sheet_upl")
    opt_sanitize_s = st.checkbox("Sanitize names (Spreadsheet)", value=True, key="san_s")
    if uploaded_sheet is not None:
        try:
            uploaded_sheet.seek(0)
            if uploaded_sheet.name.lower().endswith(".csv"):
                df_preview = pd.read_csv(uploaded_sheet, dtype=str, keep_default_na=False, nrows=200)
            else:
                df_preview = pd.read_excel(uploaded_sheet, dtype=str, engine="openpyxl").fillna("")
            col = st.selectbox("Select the column containing folder names", list(df_preview.columns))
            st.dataframe(df_preview.head(10))
            if st.button("Create Folders (From Spreadsheet)"):
                uploaded_sheet.seek(0)
                if uploaded_sheet.name.lower().endswith(".csv"):
                    df = pd.read_csv(uploaded_sheet, dtype=str, keep_default_na=False)
                else:
                    df = pd.read_excel(uploaded_sheet, dtype=str, engine="openpyxl").fillna("")
                values = [str(v).strip() for v in df[col].tolist()]
                names = [(sanitize_name(v) if opt_sanitize_s else (v or "Untitled")) for v in values if v]
                out_bytes = build_zip_from_names(names)
                st.success("Folders ZIP created.")
                st.download_button("⬇️ Download Folders.zip", out_bytes, "Folders.zip", "application/zip")
        except Exception as e:
            st.error(f"Could not read the file: {e}")
