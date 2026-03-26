"""
app.py  –  PPT → Markdown Converter (Streamlit)
------------------------------------------------
Upload a .pptx file, preview the generated Markdown,
and download the result — all in the browser.

Deploy to Streamlit Cloud:
    1. Push this repo to GitHub
    2. Go to share.streamlit.io → New app → select repo
    3. Set "Main file path" to  app.py
"""

import io
import os
import streamlit as st
from pptx import Presentation

# ─────────────────────────────────────────────
# Page config (must be first Streamlit call)
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="PPT → Markdown Converter",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# Custom CSS – dark gradient theme
# ─────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* ── Global ──────────────────────────────── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .stApp {
        background: #f5f7fa;
        min-height: 100vh;
    }

    /* ── Hero banner ─────────────────────────── */
    .hero {
        text-align: center;
        padding: 2.5rem 1rem 1.5rem;
    }
    .hero h1 {
        font-size: 2.8rem;
        font-weight: 700;
        background: linear-gradient(90deg, #7c3aed, #2563eb, #0891b2);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.4rem;
    }
    .hero p {
        color: #475569;
        font-size: 1.05rem;
        margin-top: 0;
    }

    /* ── Upload card ─────────────────────────── */
    .upload-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 2rem;
        margin: 1rem auto;
        max-width: 640px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.07);
    }

    /* ── Slide cards ─────────────────────────── */
    .slide-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 0.8rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .slide-card h3 {
        color: #7c3aed;
        margin: 0 0 0.5rem;
        font-size: 1rem;
    }
    .slide-card ul {
        color: #334155;
        margin: 0;
        padding-left: 1.2rem;
    }
    .slide-card ul li {
        margin-bottom: 0.2rem;
        font-size: 0.92rem;
    }

    /* ── Stats badge row ─────────────────────── */
    .stat-row {
        display: flex;
        gap: 1rem;
        justify-content: center;
        flex-wrap: wrap;
        margin: 1.2rem 0;
    }
    .stat-badge {
        background: #ede9fe;
        border: 1px solid #c4b5fd;
        border-radius: 999px;
        padding: 0.4rem 1.1rem;
        color: #6d28d9;
        font-size: 0.88rem;
        font-weight: 600;
    }

    /* ── Buttons ─────────────────────────────── */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #7c3aed, #2563eb) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.6rem 1.6rem !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        width: 100%;
        transition: opacity 0.2s;
    }
    .stDownloadButton > button:hover {
        opacity: 0.88 !important;
    }

    /* ── File uploader ───────────────────────── */
    [data-testid="stFileUploadDropzone"] {
        background: #f8fafc !important;
        border: 2px dashed #a78bfa !important;
        border-radius: 12px !important;
        color: #475569 !important;
    }

    /* ── Tabs ────────────────────────────────── */
    .stTabs [data-baseweb="tab-list"] {
        background: #e2e8f0;
        border-radius: 10px;
        gap: 4px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        color: #475569 !important;
        border-radius: 8px;
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        background: #ffffff !important;
        color: #7c3aed !important;
        box-shadow: 0 1px 4px rgba(0,0,0,0.10);
    }

    /* ── Code / text area ────────────────────── */
    .stTextArea textarea {
        background: #f8fafc !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
        font-family: 'Fira Mono', monospace;
        font-size: 0.85rem;
    }

    /* ── Hide default Streamlit chrome ───────── */
    #MainMenu, footer { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# Core conversion logic
# ─────────────────────────────────────────────

def _table_to_md(rows: list[list[str]]) -> str:
    """Convert a list of row-lists into a Markdown pipe table string."""
    if not rows:
        return ""
    def escape(cell: str) -> str:
        return cell.replace("|", "\\|").replace("\n", " ")
    header = "| " + " | ".join(escape(c) for c in rows[0]) + " |"
    sep    = "| " + " | ".join(["---"] * len(rows[0])) + " |"
    body   = ["| " + " | ".join(escape(c) for c in row) + " |" for row in rows[1:]]
    return "\n".join([header, sep] + body)


def parse_pptx(file_bytes: bytes) -> list[dict]:
    """
    Return a list of slide dicts.
    Each slide: {index, title, content}
    content is an ordered list of items:
      {"type": "text",  "lines": [str, ...]}
      {"type": "table", "rows":  [[str, ...], ...]}
    """
    prs = Presentation(io.BytesIO(file_bytes))
    slides = []
    for i, slide in enumerate(prs.slides, start=1):
        title = None
        content: list[dict] = []

        for shape in slide.shapes:
            # ── Tables ──────────────────────────────
            if shape.has_table:
                rows_data = [
                    [cell.text.strip() for cell in row.cells]
                    for row in shape.table.rows
                ]
                # Skip fully-empty tables
                if any(any(c for c in row) for row in rows_data):
                    content.append({"type": "table", "rows": rows_data})
                continue

            # ── Text shapes ─────────────────────────
            if not hasattr(shape, "text"):
                continue
            text = shape.text.strip()
            if not text:
                continue

            if title is None:
                title = text          # First non-empty text → slide title
            else:
                lines = [sub.strip() for sub in text.split("\n") if sub.strip()]
                if lines:
                    content.append({"type": "text", "lines": lines})

        slides.append({"index": i, "title": title or "Untitled", "content": content})
    return slides


def slides_to_markdown(slides: list[dict], pptx_name: str = "Presentation") -> str:
    """Convert parsed slides to a Markdown string (text + tables)."""
    stem = os.path.splitext(pptx_name)[0]
    lines = [f"# {stem}\n"]
    for s in slides:
        lines.append(f"## Slide {s['index']}: {s['title']}")
        for item in s["content"]:
            if item["type"] == "text":
                for line in item["lines"]:
                    lines.append(f"- {line}")
            elif item["type"] == "table":
                lines.append("")                       # blank line before table
                lines.append(_table_to_md(item["rows"]))
                lines.append("")                       # blank line after table
        lines.append("")
    return "\n".join(lines)

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────

st.markdown(
    """
    <div class="hero">
        <h1>📑 PPT → Markdown</h1>
        <p>Upload a PowerPoint deck and instantly get a clean Markdown file.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Upload ────────────────────────────────────
st.markdown('<div class="upload-card">', unsafe_allow_html=True)
uploaded = st.file_uploader(
    "Drop your .pptx file here or click to browse",
    type=["pptx"],
    label_visibility="collapsed",
)
st.markdown("</div>", unsafe_allow_html=True)

# ── Process ───────────────────────────────────
if uploaded is not None:
    file_bytes = uploaded.read()
    pptx_name  = uploaded.name

    with st.spinner("Converting slides…"):
        slides     = parse_pptx(file_bytes)
        md_text    = slides_to_markdown(slides, pptx_name)

    slide_count  = len(slides)
    bullet_count = sum(
        len(item["lines"]) for s in slides
        for item in s["content"] if item["type"] == "text"
    )
    table_count  = sum(
        1 for s in slides
        for item in s["content"] if item["type"] == "table"
    )
    word_count   = len(md_text.split())
    md_filename  = os.path.splitext(pptx_name)[0] + ".md"

    # Stats row
    table_badge = f'<span class="stat-badge">📊 {table_count} table{"s" if table_count != 1 else ""}</span>' if table_count else ""
    st.markdown(
        f"""
        <div class="stat-row">
            <span class="stat-badge">🗂 {slide_count} slides</span>
            <span class="stat-badge">• {bullet_count} bullet points</span>
            {table_badge}
            <span class="stat-badge">📝 ~{word_count} words</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Download button (prominent)
    col_dl, col_gap = st.columns([1, 2])
    with col_dl:
        st.download_button(
            label="⬇️  Download Markdown",
            data=md_text.encode("utf-8"),
            file_name=md_filename,
            mime="text/markdown",
        )

    st.divider()

    # Tabbed preview
    tab_preview, tab_raw, tab_slides = st.tabs(
        ["🖼 Rendered Preview", "📄 Raw Markdown", "🗂 Slide-by-Slide"]
    )

    with tab_preview:
        st.markdown(md_text)

    with tab_raw:
        st.text_area(
            "Markdown source",
            value=md_text,
            height=520,
            label_visibility="collapsed",
        )

    with tab_slides:
        for s in slides:
            # Build inner HTML preserving content order
            inner_html = ""
            for item in s["content"]:
                if item["type"] == "text":
                    lis = "".join(f"<li>{ln}</li>" for ln in item["lines"])
                    inner_html += f"<ul style='color:#334155;padding-left:1.2rem;margin:0.4rem 0'>{lis}</ul>"
                elif item["type"] == "table":
                    rows = item["rows"]
                    if not rows:
                        continue
                    thead = "<tr>" + "".join(
                        f"<th style='padding:6px 12px;border-bottom:2px solid #7c3aed;color:#7c3aed;text-align:left'>{c}</th>"
                        for c in rows[0]
                    ) + "</tr>"
                    tbody = "".join(
                        "<tr>" + "".join(
                            f"<td style='padding:5px 12px;border-bottom:1px solid #e2e8f0;color:#334155'>{c}</td>"
                            for c in row
                        ) + "</tr>"
                        for row in rows[1:]
                    )
                    inner_html += (
                        f"<table style='width:100%;border-collapse:collapse;margin:0.6rem 0;font-size:0.88rem'>"
                        f"<thead>{thead}</thead><tbody>{tbody}</tbody></table>"
                    )
            if not inner_html:
                inner_html = "<em style='color:#94a3b8'>No text or table content</em>"
            st.markdown(
                f"""
                <div class="slide-card">
                    <h3>Slide {s['index']}: {s['title']}</h3>
                    {inner_html}
                </div>
                """,
                unsafe_allow_html=True,
            )

else:
    st.info("👆 Upload a `.pptx` file above to get started.", icon="💡")
