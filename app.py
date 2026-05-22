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

# ─── Custom CSS ─────────────────────────────────────────────────────────────
st.markdown("""
    <style>
        .main { background-color: #f9f9f9; }
        .stButton>button {
            background-color: #4CAF50;
            color: white;
            font-size: 16px;
            padding: 10px 24px;
            border-radius: 8px;
            border: none;
            width: 100%;
        }
        .stButton>button:hover {
            background-color: #45a049;
        }
        .stTextInput>div>div>input {
            border-radius: 8px;
            font-size: 15px;
        }
        .success-box {
            background-color: #e6f4ea;
            padding: 15px;
            border-radius: 8px;
            border-left: 5px solid #4CAF50;
        }
        .error-box {
            background-color: #fdecea;
            padding: 15px;
            border-radius: 8px;
            border-left: 5px solid #f44336;
        }
    </style>
""", unsafe_allow_html=True)


# ─── Helper: Validate URL ────────────────────────────────────────────────────
def is_valid_url(url):
    try:
        result = urlparse(url)
        return all([result.scheme in ("http", "https"), result.netloc])
    except:
        return False


# ─── Helper: Scrape Text from URL ───────────────────────────────────────────
def scrape_text(url):
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

    # Remove unwanted tags (images, scripts, styles, navs, footers)
    for tag in soup(["script", "style", "img", "nav", "footer",
                     "aside", "iframe", "noscript", "svg", "figure"]):
        tag.decompose()

    # Extract structured content
    content = []

    # Page title
    title_tag = soup.find("title")
    page_title = title_tag.get_text(strip=True) if title_tag else "Untitled Page"

    # Extract headings and paragraphs in document order
    for element in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "blockquote"]):
        text = element.get_text(separator=" ", strip=True)
        text = re.sub(r'\s+', ' ', text)  # Clean extra whitespace
        if text and len(text) > 2:        # Skip empty/tiny fragments
            content.append({
                "tag": element.name,
                "text": text
            })

    return page_title, content


# ─── Helper: Build Word Document ─────────────────────────────────────────────
def build_docx(page_title, content, url):
    doc = Document()

    # ── Document Title ──
    title_para = doc.add_heading(page_title, level=0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.runs[0]
    title_run.font.color.rgb = RGBColor(0x1A, 0x73, 0xE8)  # Google Blue

    # ── Source URL ──
    source_para = doc.add_paragraph()
    source_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    source_run = source_para.add_run(f"Source: {url}")
    source_run.font.size = Pt(9)
    source_run.font.italic = True
    source_run.font.color.rgb = RGBColor(0x75, 0x75, 0x75)

    doc.add_paragraph()  # Spacer

    # ── Content ──
    heading_map = {
        "h1": 1, "h2": 2, "h3": 3,
        "h4": 4, "h5": 5, "h6": 6
    }

    for item in content:
        tag = item["tag"]
        text = item["text"]

        if tag in heading_map:
            doc.add_heading(text, level=heading_map[tag])

        elif tag == "p":
            para = doc.add_paragraph(text)
            para.paragraph_format.space_after = Pt(6)

        elif tag == "li":
            doc.add_paragraph(text, style="List Bullet")

        elif tag == "blockquote":
            para = doc.add_paragraph(text, style="Intense Quote")

    # ── Footer ──
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.text = f"Transcribed by URL Transcriber  |  {url}"
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── Save to buffer ──
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ─── Helper: Generate filename from URL ─────────────────────────────────────
def url_to_filename(url):
    parsed = urlparse(url)
    name = parsed.path.strip("/").replace("/", "_") or parsed.netloc
    name = re.sub(r'[^a-zA-Z0-9_\-]', '_', name)
    return f"{name[:60]}.docx" if name else "transcription.docx"


# ════════════════════════════════════════════════════════════════════════════
#  MAIN APP UI
# ════════════════════════════════════════════════════════════════════════════

st.title("📄 URL Text Transcriber")
st.markdown("Paste any URL below and download its text content as a **Word Document** — no images included.")
st.divider()

# ── Input: Single or Multiple URLs ──
mode = st.radio("Mode", ["Single URL", "Multiple URLs"], horizontal=True)

urls = []

if mode == "Single URL":
    url_input = st.text_input("🔗 Enter URL", placeholder="https://example.com/article")
    if url_input:
        urls = [url_input.strip()]

else:
    url_text = st.text_area(
        "🔗 Enter URLs (one per line)",
        placeholder="https://example.com/article-1\nhttps://example.com/article-2",
        height=150
    )
    if url_text:
        urls = [u.strip() for u in url_text.strip().splitlines() if u.strip()]

st.divider()

# ── Transcribe Button ──
if st.button("✨ Transcribe & Download"):

    if not urls:
        st.warning("⚠️ Please enter at least one URL.")

    else:
        invalid = [u for u in urls if not is_valid_url(u)]
        if invalid:
            st.error(f"❌ Invalid URL(s): {', '.join(invalid)}")
        else:
            # ── Single URL → direct download ──
            if len(urls) == 1:
                url = urls[0]
                with st.spinner(f"Scraping content from {url}..."):
                    try:
                        page_title, content = scrape_text(url)

                        if not content:
                            st.error("❌ No text content found on that page.")
                        else:
                            docx_buffer = build_docx(page_title, content, url)
                            filename = url_to_filename(url)

                            st.success(f"✅ **{len(content)} text blocks** extracted from: _{page_title}_")

                            # Preview
                            with st.expander("👁️ Preview extracted text"):
                                for item in content[:20]:
                                    if item["tag"].startswith("h"):
                                        st.markdown(f"**{item['text']}**")
                                    else:
                                        st.write(item["text"])
                                if len(content) > 20:
                                    st.caption(f"... and {len(content) - 20} more blocks in the document.")

                            st.download_button(
                                label="📥 Download Word Document",
                                data=docx_buffer,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

                    except requests.exceptions.ConnectionError:
                        st.error("❌ Could not connect. Check the URL or your internet connection.")
                    except requests.exceptions.Timeout:
                        st.error("❌ Request timed out. The site may be too slow.")
                    except requests.exceptions.HTTPError as e:
                        st.error(f"❌ HTTP Error: {e}")
                    except Exception as e:
                        st.error(f"❌ Unexpected error: {e}")

            # ── Multiple URLs → combined doc ──
            else:
                combined_doc = Document()
                combined_doc.add_heading("Combined URL Transcription", level=0)
                combined_doc.add_paragraph()

                success_count = 0
                progress = st.progress(0)
                status = st.empty()

                for i, url in enumerate(urls):
                    status.text(f"Processing {i+1}/{len(urls)}: {url}")
                    try:
                        page_title, content = scrape_text(url)
                        if content:
                            combined_doc.add_heading(page_title, level=1)
                            source_para = combined_doc.add_paragraph(f"Source: {url}")
                            source_para.runs[0].font.italic = True
                            source_para.runs[0].font.size = Pt(9)
                            combined_doc.add_paragraph()

                            for item in content:
                                tag = item["tag"]
                                text = item["text"]
                                heading_map = {"h1":1,"h2":2,"h3":3,"h4":4,"h5":5,"h6":6}
                                if tag in heading_map:
                                    combined_doc.add_heading(text, level=heading_map[tag])
                                elif tag == "p":
                                    combined_doc.add_paragraph(text)
                                elif tag == "li":
                                    combined_doc.add_paragraph(text, style="List Bullet")
                                elif tag == "blockquote":
                                    combined_doc.add_paragraph(text, style="Intense Quote")

                            combined_doc.add_page_break()
                            success_count += 1

                    except Exception as e:
                        combined_doc.add_paragraph(f"⚠️ Failed to scrape {url}: {e}")

                    progress.progress((i + 1) / len(urls))

                status.empty()
                progress.empty()

                buffer = io.BytesIO()
                combined_doc.save(buffer)
                buffer.seek(0)

                st.success(f"✅ Successfully transcribed **{success_count}/{len(urls)}** URLs.")
                st.download_button(
                    label="📥 Download Combined Word Document",
                    data=buffer,
                    file_name="combined_transcription.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

st.divider()
st.caption("💡 Note: Some sites may block automated requests. Results may vary depending on the website.")
