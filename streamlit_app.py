import io
import re
import zipfile
import unicodedata
from pathlib import PurePosixPath

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="Folder Creator 25", page_icon="📁", layout="centered")

# ----- Simple header (we can swap to your exact branding once you share colors/logo) -----
st.markdown("""<style>
:root { --fc-accent: #c23a2b; }
.block-container { padding-top: 2rem; }
h1, .stTabs [data-baseweb="tab"] p { font-weight: 700; }
.stTabs [data-baseweb="tab-highlight"] { background: var(--fc-accent); }
.stTabs [data-baseweb="tab"] { padding-top: .75rem; padding-bottom: .75rem; }
.st-emotion-cache-1wmy9hl p { margin-bottom: .25rem; }
</style>
""", unsafe_allow_html=True)

st.title("📁 Folder Creator 25")

# ---------- Helpers ----------
FORBIDDEN = r"\\/:*?\"<>|"
SYSTEM_DIR_PREFIXES = ("__MACOSX/",)
SYSTEM_BASENAMES = {".DS_Store", "Thumbs.db"}
_forbidden_pattern = re.compile(f"[{re.escape(FORBIDDEN)}]|[\x00-\x1f]")

def sanitize_name(name: str) -> str:
    # Always keep this internal for filesystem safety
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
            folder_name = make_unique(sanitize_name(n), existing)
            _write_empty_folder(zout, folder_name)
    return out_buf.getvalue()

def build_zip_from_uploaded_files(files, include_files: bool):
    out_buf = io.BytesIO()
    existing = set()
    with zipfile.ZipFile(out_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zout:
        for f in files:
            basename = PurePosixPath(f.name).name
            stem = PurePosixPath(basename).stem
            folder_name = make_unique(sanitize_name(stem), existing)
            _write_empty_folder(zout, folder_name)
            if include_files:
                data = f.read()
                f.seek(0)
                dest = str(PurePosixPath(folder_name) / basename)
                zout.writestr(dest, data)
    return out_buf.getvalue()

def build_zip_from_filenames_zip(uploaded_zip, include_files: bool):
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
                folder_name = make_unique(sanitize_name(stem), existing)
                _write_empty_folder(zout, folder_name)
                if include_files:
                    data = zin.read(info.filename)
                    dest = str(PurePosixPath(folder_name) / basename)
                    zout.writestr(dest, data)
    return out_buf.getvalue()

# ---------- UI ----------
tab1, tab2, tab3 = st.tabs(["From Filenames", "From Word (.docx)", "From Spreadsheet (.csv/.xlsx)"])

# --- Tab 1: Filenames ---
with tab1:
    st.caption("Choose files directly or upload a ZIP. You'll download a ZIP of folders.")
    mode = st.radio("Choose how to provide files:", ["Select multiple files", "Upload a ZIP of files"], horizontal=True)
    include_files = st.checkbox("Include original files in each folder", value=True)

    if mode == "Select multiple files":
        files = st.file_uploader("Select files (you can choose many)", accept_multiple_files=True, type=None, key="multi_files")
        if files and st.button("Create Folders (From Selected Files)", type="primary"):
            out_bytes = build_zip_from_uploaded_files(files, include_files)
            st.success("Folders ZIP created.")
            st.download_button("⬇️ Download Folders.zip", out_bytes, "Folders.zip", "application/zip")
    else:
        uploaded_zip = st.file_uploader("Upload a ZIP of files", type=["zip"], key="zip_upl")
        if uploaded_zip and st.button("Create Folders (From ZIP)", type="primary"):
            out_bytes = build_zip_from_filenames_zip(uploaded_zip, include_files)
            st.success("Folders ZIP created.")
            st.download_button("⬇️ Download Folders.zip", out_bytes, "Folders.zip", "application/zip")

# --- Tab 2: Word (.docx) ---
with tab2:
    uploaded_docx = st.file_uploader("Upload DOCX (one folder name per line)", type=["docx"], key="docx_upl")
    if uploaded_docx is not None and st.button("Create Folders (From Word)"):
        doc = Document(uploaded_docx)
        names = [p.text for p in doc.paragraphs if p.text.strip()]
        out_bytes = build_zip_from_names(names)
        st.success("Folders ZIP created.")
        st.download_button("⬇️ Download Folders.zip", out_bytes, "Folders.zip", "application/zip")

# --- Tab 3: Spreadsheet (.csv/.xlsx) ---
with tab3:
    uploaded_sheet = st.file_uploader("Upload CSV or XLSX", type=["csv", "xlsx"], key="sheet_upl")
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
                values = [str(v).strip() for v in df[col].tolist() if str(v).strip()]
                out_bytes = build_zip_from_names(values)
                st.success("Folders ZIP created.")
                st.download_button("⬇️ Download Folders.zip", out_bytes, "Folders.zip", "application/zip")
        except Exception as e:
            st.error(f"Could not read the file: {e}")
