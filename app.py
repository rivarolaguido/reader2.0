import streamlit as st
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
from urllib.parse import urlparse

# ─── Page Config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="URL Text Transcriber",
    page_icon="📄",
    layout="centered"
)

# ─── Hide ALL Streamlit UI ───────────────────────────────────────────────────
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .stDeployButton {display: none;}
        .stToolbar {display: none;}
        .block-container {
            max-width: 650px;
            padding-top: 50px;
        }
        body {
            background-color: #ffffff;
            font-family: 'Segoe UI', sans-serif;
        }
        .stTextInput > div > div > input {
            border-radius: 8px;
            border: 1.5px solid #d0d7de;
            padding: 12px;
            font-size: 15px;
        }
        .stButton > button {
            background-color: #1a73e8;
            color: white;
            font-size: 15px;
            font-weight: 600;
            padding: 12px 0;
            border-radius: 8px;
            border: none;
            width: 100%;
        }
        .stDownloadButton > button {
            background-color: #34a853;
            color: white;
            font-size: 15px;
            font-weight: 600;
            padding: 12px 0;
            border-radius: 8px;
            border: none;
            width: 100%;
        }
    </style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
#  CORE: Extract ONLY Blog Text
# ════════════════════════════════════════════════════════════════════════════

def extract_blog_text(url):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/114.0.0.0 Safari/537.36"
        )
    }

    response = requests.get(url, headers=headers, timeout=15)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "lxml")

    # ── Step 1: Get Page Title ──────────────────────────────────────────────
    title_tag  = soup.find("title")
    page_title = title_tag.get_text(strip=True) if title_tag else "Untitled"

    # ── Step 2: Remove ALL non-content tags ────────────────────────────────
    noise_tags = [
        "script", "style", "img", "picture", "video", "audio",
        "source", "track", "iframe", "embed", "object", "canvas",
        "svg", "figure", "figcaption", "nav", "header", "footer",
        "aside", "form", "input", "button", "select", "textarea",
        "label", "noscript", "link", "meta", "head", "map", "area",
        "dialog", "menu", "menuitem", "toolbar", "template"
    ]
    for tag in soup(noise_tags):
        tag.decompose()

    # ── Step 3: Remove UI elements by class/id keywords ────────────────────
    # FIX: Check element is not None and is a Tag before calling .get()
    ui_keywords = [
        "nav", "navbar", "navigation", "menu", "sidebar", "header",
        "footer", "banner", "cookie", "popup", "modal", "overlay",
        "ad", "ads", "advertisement", "promo", "social", "share",
        "comment", "related", "recommended", "subscribe", "newsletter",
        "breadcrumb", "pagination", "widget", "toolbar", "tooltip",
        "dropdown", "notification", "tag-list", "label", "meta",
        "author", "byline", "dateline", "timestamp", "category",
        "search", "login", "signup", "cart", "checkout", "rating",
        "review", "feedback", "table-of-contents", "toc"
    ]

    for element in soup.find_all(True):
        try:
            # ✅ FIX: Guard against NoneType and NavigableString
            if element is None:
                continue
            if not hasattr(element, 'get'):
                continue

            el_id    = element.get("id") or ""
            el_class = element.get("class") or []

            # el_class can be a list or a string — normalize to string
            if isinstance(el_class, list):
                el_class = " ".join(el_class)

            combined = (el_id + " " + el_class).lower()

            if any(kw in combined for kw in ui_keywords):
                element.decompose()

        except Exception:
            # Skip any element that causes issues
            continue

    # ── Step 4: Find the main blog content container ────────────────────────
    main_content = (
        soup.find("article")                                        or
        soup.find("main")                                           or
        soup.find(attrs={"role": "main"})                           or
        soup.find("div", class_=re.compile(
            r'(post|blog|article|content|entry|body|text|story)',
            re.IGNORECASE
        ))                                                          or
        soup.find("div", id=re.compile(
            r'(post|blog|article|content|entry|body|text|story)',
            re.IGNORECASE
        ))                                                          or
        soup.body
    )

    # ── Fallback if nothing found ───────────────────────────────────────────
    if main_content is None:
        main_content = soup

    # ── Step 5: Extract clean text blocks ───────────────────────────────────
    content = []
    seen    = set()

    tags_to_extract = ["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "blockquote"]

    for element in main_content.find_all(tags_to_extract):
        try:
            text = element.get_text(separator=" ", strip=True)
            text = re.sub(r'\s+', ' ', text).strip()

            if not text:                           # Empty
                continue
            if len(text) < 20:                     # Too short
                continue
            if text in seen:                       # Duplicate
                continue
            if re.match(r'^[\W\d\s]+$', text):    # Symbols/numbers only
                continue
            if text.startswith("©"):              # Copyright
                continue
            if re.match(r'^https?://', text):     # Raw URL
                continue

            seen.add(text)
            content.append({"tag": element.name, "text": text})

        except Exception:
            continue

    return page_title, content


# ════════════════════════════════════════════════════════════════════════════
#  BUILD: Word Document (Text Only)
# ════════════════════════════════════════════════════════════════════════════

def build_docx(page_title, content, url):
    doc = Document()

    # ── Margins ──
    for section in doc.sections:
        section.top_margin    = Pt(72)
        section.bottom_margin = Pt(72)
        section.left_margin   = Pt(80)
        section.right_margin  = Pt(80)

    # ── Title ──
    title_para = doc.add_heading(page_title, level=0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title_para.runs:
        title_para.runs[0].font.color.rgb = RGBColor(0x1A, 0x73, 0xE8)

    # ── Source URL ──
    src_para = doc.add_paragraph()
    src_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    src_run = src_para.add_run(f"Source: {url}")
    src_run.font.size      = Pt(9)
    src_run.font.italic    = True
    src_run.font.color.rgb = RGBColor(0x75, 0x75, 0x75)

    doc.add_paragraph()  # Spacer

    # ── Heading Map ──
    heading_map = {
        "h1": 1, "h2": 2, "h3": 3,
        "h4": 4, "h5": 5, "h6": 6
    }

    # ── Content ──
    for item in content:
        tag  = item["tag"]
        text = item["text"]

        if tag in heading_map:
            doc.add_heading(text, level=heading_map[tag])

        elif tag == "p":
            para = doc.add_paragraph(text)
            para.paragraph_format.space_after = Pt(6)
            for run in para.runs:
                run.font.size = Pt(11)

        elif tag == "li":
            doc.add_paragraph(text, style="List Bullet")

        elif tag == "blockquote":
            doc.add_paragraph(text, style="Intense Quote")

    # ── Footer ──
    footer      = doc.sections[0].footer
    footer_para = footer.paragraphs[0]
    footer_para.text = f"Transcribed by URL Transcriber  |  {url}"
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if footer_para.runs:
        footer_para.runs[0].font.size      = Pt(8)
        footer_para.runs[0].font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)

    # ── Save ──
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════════════════

def is_valid_url(url):
    try:
        r = urlparse(url)
        return all([r.scheme in ("http", "https"), r.netloc])
    except Exception:
        return False


def url_to_filename(url):
    parsed = urlparse(url)
    name   = parsed.path.strip("/").replace("/", "_") or parsed.netloc
    name   = re.sub(r'[^a-zA-Z0-9_\-]', '_', name)
    return f"{name[:60]}.docx" if name else "transcription.docx"


# ════════════════════════════════════════════════════════════════════════════
#  APP UI — Minimal: Input + Download Only
# ════════════════════════════════════════════════════════════════════════════

st.markdown("## 📄 URL Text Transcriber")
st.caption("Extracts only the blog/article text. No images, no UI elements.")
st.divider()

url_input = st.text_input(
    "",
    placeholder="🔗  Paste URL here... e.g. https://example.com/article",
    label_visibility="collapsed"
)

if st.button("Extract & Download"):
    if not url_input.strip():
        st.warning("⚠️ Please enter a URL.")
    elif not is_valid_url(url_input.strip()):
        st.error("❌ Invalid URL. Make sure it starts with http:// or https://")
    else:
        url = url_input.strip()
        with st.spinner("Extracting blog text..."):
            try:
                page_title, content = extract_blog_text(url)

                if not content:
                    st.error("❌ No readable text found. The page may be JavaScript-rendered or blocked.")
                else:
                    docx_buffer = build_docx(page_title, content, url)
                    filename    = url_to_filename(url)

                    st.success(f"✅ Done! **{len(content)} text blocks** extracted.")
                    st.download_button(
                        label     = "📥 Download Word Document",
                        data      = docx_buffer,
                        file_name = filename,
                        mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except requests.exceptions.ConnectionError:
                st.error("❌ Connection failed. Check the URL.")
            except requests.exceptions.Timeout:
                st.error("❌ Request timed out.")
            except requests.exceptions.HTTPError as e:
                st.error(f"❌ HTTP Error: {e}")
            except Exception as e:
                st.error(f"❌ Error: {e}")

st.divider()
st.caption("⚠️ Some sites block automated requests or use JavaScript rendering — results may vary.")
