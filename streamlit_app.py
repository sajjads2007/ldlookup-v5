import io
import re
from pathlib import Path
from typing import List, Optional

import pandas as pd
import requests
import streamlit as st

# Optional readers
import pdfplumber              # PDF text
from docx import Document      # DOCX text

# -----------------------------
# Page + theme
# -----------------------------
OXFORD = "#0B132B"
st.set_page_config(page_title="LD Lookup (V5)", page_icon="🧾", layout="wide")

# Minimal PWA hooks (served as static files from repo root on Streamlit Cloud)
st.markdown(
    """
    <meta name="theme-color" content="#0B132B">
    <link rel="manifest" href="manifest.json">
    <script>
    // register SW if present (best-effort)
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('service-worker.js').catch(()=>{});
      });
    }
    </script>
    """,
    unsafe_allow_html=True,
)

# Light CSS for mobile/responsive + oxford blue accents
st.markdown(
    f"""
    <style>
      .topbar {{
        background: {OXFORD};
        color: #ffffff;
        padding: .9rem 1.1rem;
        border-radius: 12px;
        margin-bottom: .75rem;
      }}
      .topbar h1 {{ margin: 0; font-size: 1.15rem; }}
      .runrow {{ position: sticky; top: 0; z-index: 5; background: white; padding: .5rem 0 .35rem 0; }}
      .thumb {{
        height: 72px; width: auto; object-fit: contain; border-radius: 6px;
        border: 1px solid #e6e6e6; background: #fff;
        transition: transform .05s ease-in-out;
      }}
      .thumb:hover {{ transform: scale(1.02); }}
      .tbl table {{
        border-collapse: collapse; width: 100%;
      }}
      .tbl th, .tbl td {{
        border: 1px solid #e6e6e6; padding: 8px; vertical-align: middle; font-size: 0.9rem;
      }}
      .tbl th {{
        background: #f6f7fb; text-align: left; color: {OXFORD};
      }}
      @media (max-width: 640px) {{
        .thumb {{ height: 64px; }}
        .tbl th, .tbl td {{ font-size: .88rem; }}
      }}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class="topbar">
      <h1>LD Lookup — Version 5.0</h1>
      <div style="opacity:.85;font-size:.95rem">Find L-numbers in files/text, show thumbnails, and export to Excel/PDF.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Config
# -----------------------------
IMG_WIDTH = 640
IMG_QUALITY = 75
CDN_TEMPLATE = "https://cdn-tp2.mozu.com/28945-m4/cms/files/{L}.jpg?w={w}&q={q}"
LNUM_RE = re.compile(r"\bL\d+\b", flags=re.IGNORECASE)

# -----------------------------
# Helpers
# -----------------------------
def extract_lnumbers_from_text(text: str) -> List[str]:
    if not text:
        return []
    candidates = [m.group(0).upper() for m in LNUM_RE.finditer(text)]
    # dedupe preserve order
    seen, out = set(), []
    for c in candidates:
        if c not in seen:
            seen.add(c); out.append(c)
    return out

def extract_lnumbers_from_dataframe(df: pd.DataFrame) -> List[str]:
    lnumbers, seen = [], set()
    for col in df.columns:
        for val in df[col].astype(str).tolist():
            for l in extract_lnumbers_from_text(val):
                if l not in seen:
                    seen.add(l); lnumbers.append(l)
    return lnumbers

def build_image_url(l_number: str, width=IMG_WIDTH, quality=IMG_QUALITY) -> str:
    l = str(l_number).strip().upper()
    return CDN_TEMPLATE.format(L=l, w=width, q=quality)

def read_any_file(file) -> str:
    """
    Return all readable text from uploaded file.
    Supports: CSV, XLSX/XLS, TXT, DOCX, PDF.
    For legacy .doc, ask user to convert to .docx.
    """
    name = file.name
    suffix = Path(name).suffix.lower()

    if suffix in [".xlsx", ".xls", ".csv"]:
        try:
            if suffix == ".csv":
                df = pd.read_csv(file, dtype=str, on_bad_lines="skip")
            else:
                # xlsx via openpyxl; xls via xlrd (pinned in requirements)
                df = pd.read_excel(file, dtype=str)
            return "\n".join(df.astype(str).fillna("").agg(" ".join, axis=1).tolist())
        except Exception as e:
            return f"READ_ERROR: {e}"

    if suffix == ".txt":
        return file.read().decode("utf-8", errors="ignore")

    if suffix == ".docx":
        try:
            doc = Document(file)
            parts = [p.text for p in doc.paragraphs]
            # also parse tables if present
            for tbl in doc.tables:
                for row in tbl.rows:
                    parts.append(" ".join(cell.text for cell in row.cells))
            return "\n".join(parts)
        except Exception as e:
            return f"READ_ERROR: {e}"

    if suffix == ".pdf":
        try:
            with pdfplumber.open(file) as pdf:
                texts = []
                for page in pdf.pages:
                    texts.append(page.extract_text() or "")
                return "\n".join(texts)
        except Exception as e:
            return f"READ_ERROR: {e}"

    if suffix == ".doc":
        return ("READ_ERROR: Legacy .doc detected. Please convert to .docx or PDF and upload again.")

    return "READ_ERROR: Unsupported file type."

def make_excel_bytes(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df_xlsx = df.copy()

        def excel_img(url, alt_text=""):
            safe_url = str(url).replace('"', '""')
            safe_alt = str(alt_text).replace('"', '""')
            return f'=IMAGE("{safe_url}","{safe_alt}")'

        df_xlsx["Photo"] = [excel_img(u, l) for u, l in zip(df_xlsx["ImageURL"], df_xlsx["LNumber"])]
        df_xlsx[["LNumber", "ImageURL", "Photo"]].to_excel(writer, index=False, sheet_name="Results")
        ws = writer.sheets["Results"]
        ws.set_column(0, 0, 14)  # LNumber
        ws.set_column(1, 1, 60)  # URL
        ws.set_column(2, 2, 18)  # Photo
    out.seek(0)
    return out.read()

def make_pdf_bytes(df: pd.DataFrame, include_images: bool = False, thumb_h: int = 72) -> bytes:
    """
    Lightweight PDF export. If include_images=True, downloads images and embeds thumbnails.
    """
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.utils import ImageReader

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("LD Lookup — Results", styles["Title"]))
    story.append(Paragraph("Version 5.0", styles["Normal"]))
    story.append(Spacer(1, 10))

    data = [["LNumber", "ImageURL", "Photo (thumb)"]]
    for _, r in df.iterrows():
        lnum = str(r["LNumber"])
        url  = str(r["ImageURL"])
        cell_img = ""
        if include_images:
            try:
                resp = requests.get(url, timeout=5)
                if resp.status_code // 100 == 2:
                    ir = ImageReader(io.BytesIO(resp.content))
                    img = Image(ir, hAlign="LEFT")
                    # keep aspect ratio based on height
                    iw, ih = ir.getSize()
                    scale = thumb_h / float(ih)
                    img.drawHeight = thumb_h
                    img.drawWidth = iw * scale
                    cell_img = img
                else:
                    cell_img = ""
            except Exception:
                cell_img = ""
        data.append([lnum, url, cell_img])

    tbl = Table(data, colWidths=[90, 360, 120])
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor(OXFORD)),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("GRID", (0,0), (-1,-1), 0.5, colors.lightgrey),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))
    story.append(tbl)
    doc.build(story)
    buf.seek(0)
    return buf.read()

# Placeholder (kept off by default). Wire to a known endpoint if available.
def maybe_lookup_name_for_lnumber(l_number: str) -> Optional[str]:
    return None

# -----------------------------
# Inputs
# -----------------------------
with st.expander("Input options", expanded=True):
    c1, c2 = st.columns([1.3, 1])
    with c1:
        uploaded = st.file_uploader(
            "Upload CSV / XLSX / XLS / TXT / PDF / DOCX",
            type=["csv", "xlsx", "xls", "txt", "pdf", "docx"],
            key="file_input_v5"
        )
        st.caption("Tip: For old .doc, convert to .docx first.")
    with c2:
        pasted = st.text_area("Or paste any text", height=180)

with st.container():
    st.markdown('<div class="runrow"></div>', unsafe_allow_html=True)
    run = st.button("🚀 Run L-Number Lookup", use_container_width=True, type="primary")

if not run:
    st.stop()

# -----------------------------
# Collect L-numbers
# -----------------------------
lnumbers: List[str] = []

if uploaded is not None:
    text = read_any_file(uploaded)
    if text.startswith("READ_ERROR:"):
        st.error(text)
        st.stop()
    lnumbers = extract_lnumbers_from_text(text)

if not lnumbers and pasted.strip():
    lnumbers = extract_lnumbers_from_text(pasted)

if not lnumbers:
    st.warning("No L-numbers detected. Upload a file or paste text containing items like **L1304179**.")
    st.stop()

# Build results dataframe
df = pd.DataFrame({"LNumber": lnumbers})
df["LNumber"] = df["LNumber"].str.upper().str.strip()
df["ImageURL"] = df["LNumber"].apply(build_image_url)

# -----------------------------
# Results table (HTML so thumbnails render + click to full size)
# -----------------------------
rows = []
for _, r in df.iterrows():
    url = r["ImageURL"]
    lnum = r["LNumber"]
    thumb = f'<a href="{url}" target="_blank" title="Open full size"><img class="thumb" src="{url}" alt="{lnum}"/></a>'
    rows.append(f"<tr><td>{lnum}</td><td><a href='{url}' target='_blank'>{url}</a></td><td>{thumb}</td></tr>")

st.subheader("Results")
st.markdown(
    f"""
    <div class="tbl">
    <table>
      <thead><tr><th>LNumber</th><th>ImageURL</th><th>Photo (thumb)</th></tr></thead>
      <tbody>{''.join(rows)}</tbody>
    </table>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Exports
# -----------------------------
st.divider()
st.subheader("Export")
c1, c2, c3 = st.columns([1,1,2])
with c1:
    st.download_button(
        "⬇️ Excel (with IMAGE formula)",
        data=make_excel_bytes(df),
        file_name="ld_lookup_v5.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with c2:
    include_pics = st.toggle("Embed images in PDF (slower)", value=False)
    st.download_button(
        "⬇️ PDF",
        data=make_pdf_bytes(df, include_images=include_pics),
        file_name="ld_lookup_v5.pdf",
        mime="application/pdf"
    )

st.caption("Fast by default. PDF embedding downloads images; toggle it only if you need pictures inside the PDF.")
