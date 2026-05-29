"""
Mortgage Suitability Report Simplifier
Run: uvicorn main:app --reload
"""

import io
import os
import re
import json
import urllib.request
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, Response
from fastapi.middleware.cors import CORSMiddleware
import anthropic
from pypdf import PdfReader
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, CondPageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER

app = FastAPI(title="Mortgage Report Simplifier")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
USAGE_FILE = Path("usage.json")
LOGO_PATH = Path("logo")          # stored without extension; we keep the original bytes + mime type
LOGO_META_FILE = Path("logo_meta.json")
LOGO_COLORS_FILE = Path("logo_colors.json")
ALLOWED_IMAGE_TYPES = {"image/png", "image/jpeg", "image/svg+xml", "image/webp"}

FONTS_DIR = Path("fonts")

# Script families — languages grouped by the font they need
DEVANAGARI_LANGS = {"Hindi", "Nepali"}
ARABIC_LANGS     = {"Arabic", "Urdu"}
BENGALI_LANGS    = {"Bengali"}
GURMUKHI_LANGS   = {"Punjabi"}
CJK_LANGS        = {"Mandarin Chinese", "Korean", "Japanese"}

# Noto font download URLs (Google Fonts GitHub)
_NOTO_BASE = "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf"
NOTO_FONT_URLS = {
    "devanagari": (
        f"{_NOTO_BASE}/NotoSansDevanagari/NotoSansDevanagari-Regular.ttf",
        f"{_NOTO_BASE}/NotoSansDevanagari/NotoSansDevanagari-Bold.ttf",
        "NotoDevanagari", "NotoDevanagari-Bold",
    ),
    "arabic": (
        f"{_NOTO_BASE}/NotoSansArabic/NotoSansArabic-Regular.ttf",
        f"{_NOTO_BASE}/NotoSansArabic/NotoSansArabic-Bold.ttf",
        "NotoArabic", "NotoArabic-Bold",
    ),
    "bengali": (
        f"{_NOTO_BASE}/NotoSansBengali/NotoSansBengali-Regular.ttf",
        f"{_NOTO_BASE}/NotoSansBengali/NotoSansBengali-Bold.ttf",
        "NotoBengali", "NotoBengali-Bold",
    ),
    "gurmukhi": (
        f"{_NOTO_BASE}/NotoSansGurmukhi/NotoSansGurmukhi-Regular.ttf",
        f"{_NOTO_BASE}/NotoSansGurmukhi/NotoSansGurmukhi-Bold.ttf",
        "NotoGurmukhi", "NotoGurmukhi-Bold",
    ),
}

# Characters that represent emoji / unrenderable symbols (keep this narrow)
_EMOJI_RE = re.compile(
    r"[\U0001F300-\U0001F9FF\U0001FA00-\U0001FAFF"
    r"\U00002600-\U000027BF\U0000FE00-\U0000FE0F]+"
)


def _script_family(language: str) -> str | None:
    if language in DEVANAGARI_LANGS: return "devanagari"
    if language in ARABIC_LANGS:     return "arabic"
    if language in BENGALI_LANGS:    return "bengali"
    if language in GURMUKHI_LANGS:   return "gurmukhi"
    return None


def ensure_noto_font(language: str) -> tuple[str, str]:
    """
    Return (body_font, bold_font) names for the given language.
    Downloads and registers a Noto font if needed.
    Falls back to Helvetica on any error.
    """
    family = _script_family(language)
    if family is None or family not in NOTO_FONT_URLS:
        return "Helvetica", "Helvetica-Bold"

    url_reg, url_bold, name_reg, name_bold = NOTO_FONT_URLS[family]

    # Already registered this session?
    if name_reg in pdfmetrics._fonts:
        return name_reg, name_bold

    try:
        FONTS_DIR.mkdir(exist_ok=True)
        path_reg  = FONTS_DIR / f"{name_reg}.ttf"
        path_bold = FONTS_DIR / f"{name_bold}.ttf"
        if not path_reg.exists():
            urllib.request.urlretrieve(url_reg,  str(path_reg))
        if not path_bold.exists():
            urllib.request.urlretrieve(url_bold, str(path_bold))
        pdfmetrics.registerFont(TTFont(name_reg,  str(path_reg)))
        pdfmetrics.registerFont(TTFont(name_bold, str(path_bold)))
        return name_reg, name_bold
    except Exception:
        return "Helvetica", "Helvetica-Bold"


# ── Logo helpers ──────────────────────────────────────────────────────────────

def save_logo(data: bytes, content_type: str):
    LOGO_PATH.write_bytes(data)
    LOGO_META_FILE.write_text(json.dumps({"content_type": content_type}))

def delete_logo():
    if LOGO_PATH.exists():
        LOGO_PATH.unlink()
    if LOGO_META_FILE.exists():
        LOGO_META_FILE.unlink()
    if LOGO_COLORS_FILE.exists():
        LOGO_COLORS_FILE.unlink()

def logo_content_type() -> str | None:
    if LOGO_META_FILE.exists():
        return json.loads(LOGO_META_FILE.read_text()).get("content_type")
    return None


def _hex_to_luminance(hex_color: str) -> float:
    """Return relative luminance (0=black, 1=white) for a hex colour."""
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i+2], 16) / 255.0 for i in (0, 2, 4))
    def lin(c):
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


def extract_logo_colors(data: bytes, content_type: str) -> dict:
    """
    Extract brand colours from the uploaded logo.
    Returns {"primary": "#rrggbb"} — the darkest colour found,
    suitable for use as a heading/title colour.
    Falls back to the default navy if nothing useful is found.
    """
    DEFAULT = "#1e3a5f"
    candidates: list[str] = []

    if content_type == "image/svg+xml":
        svg_text = data.decode("utf-8", errors="ignore")
        # Pick up fill: #rrggbb in <style> blocks and inline fill="..." attributes
        candidates += re.findall(r'fill:\s*(#[0-9a-fA-F]{6})\b', svg_text)
        candidates += re.findall(r'fill="(#[0-9a-fA-F]{6})"', svg_text)
    else:
        try:
            from PIL import Image as PILImage
            import colorsys

            img = PILImage.open(io.BytesIO(data)).convert("RGBA")
            img.thumbnail((200, 200))  # shrink for speed

            # Collect all non-transparent pixels
            pixels = [
                (r, g, b)
                for r, g, b, a in img.getdata()
                if a > 128  # skip transparent
            ]
            if not pixels:
                return {"primary": DEFAULT}

            # Quantise: bucket each channel to nearest 16
            buckets: dict[tuple, int] = {}
            for r, g, b in pixels:
                key = (r >> 4, g >> 4, b >> 4)
                buckets[key] = buckets.get(key, 0) + 1

            # Sort by frequency, convert back to full hex
            for (r4, g4, b4), _ in sorted(buckets.items(), key=lambda x: -x[1])[:20]:
                r, g, b = r4 * 16 + 8, g4 * 16 + 8, b4 * 16 + 8
                # Skip near-white and near-black (unlikely brand colours)
                lum = _hex_to_luminance(f"#{r:02x}{g:02x}{b:02x}")
                if 0.02 < lum < 0.70:
                    candidates.append(f"#{r:02x}{g:02x}{b:02x}")

        except Exception:
            pass

    # Filter out near-white colours and pick the darkest (best for dark headings)
    dark_candidates = [
        c for c in candidates
        if _hex_to_luminance(c) < 0.35
    ]

    if dark_candidates:
        primary = min(dark_candidates, key=_hex_to_luminance)
    elif candidates:
        primary = min(candidates, key=_hex_to_luminance)
    else:
        primary = DEFAULT

    return {"primary": primary}


def load_logo_colors() -> dict:
    if LOGO_COLORS_FILE.exists():
        return json.loads(LOGO_COLORS_FILE.read_text())
    return {}


COST_PER_REPORT = 0.02  # estimated USD mid-point


# ── Usage tracking ────────────────────────────────────────────────────────────

def _month_key() -> str:
    return datetime.utcnow().strftime("%Y-%m")

def load_usage() -> dict:
    if USAGE_FILE.exists():
        return json.loads(USAGE_FILE.read_text())
    return {}

def increment_usage() -> dict:
    """Increment this month's counter and return current stats."""
    data = load_usage()
    key = _month_key()
    data[key] = data.get(key, 0) + 1
    USAGE_FILE.write_text(json.dumps(data, indent=2))
    return data

def get_stats() -> dict:
    data = load_usage()
    key = _month_key()
    this_month = data.get(key, 0)
    all_time = sum(data.values())
    return {
        "month": key,
        "this_month": this_month,
        "all_time": all_time,
        "estimated_cost_usd": round(this_month * COST_PER_REPORT, 2),
        "history": {k: v for k, v in sorted(data.items())},
    }

SIMPLIFY_PROMPT = """You are an expert at making complex mortgage suitability reports easy to understand for everyday people.

You will be given a mortgage suitability report. Your job is to produce a simplified, plain-English summary that:
- Uses simple, everyday language (no jargon)
- Highlights the KEY FACTS about the recommended mortgage (rate, term, monthly payment, lender, type)
- Summarises the CLIENT'S PREFERENCES and circumstances that led to this recommendation
- Explains WHY this mortgage was recommended in 2-3 plain sentences
- Flags any important risks or things the client should be aware of
- Is warm, reassuring, and easy to read

Structure your response with these exact section headings (use markdown ## for headings):
## Your Mortgage at a Glance
## About You & Your Situation
## Why This Mortgage Was Recommended
## Key Things to Know
## Next Steps

Keep each section concise — bullet points are fine. Avoid anything that sounds like a legal document.

{language_instruction}

Here is the mortgage suitability report to simplify:

---
{report_text}
---
"""

LANGUAGE_INSTRUCTION = {
    "English": "",
    "Spanish": "Write your entire response in Spanish (Español).",
    "French": "Write your entire response in French (Français).",
    "German": "Write your entire response in German (Deutsch).",
    "Italian": "Write your entire response in Italian (Italiano).",
    "Portuguese": "Write your entire response in Portuguese (Português).",
    "Polish": "Write your entire response in Polish (Polski).",
    "Romanian": "Write your entire response in Romanian (Română).",
    "Dutch": "Write your entire response in Dutch (Nederlands).",
    "Arabic": "Write your entire response in Arabic (العربية). Use right-to-left formatting where possible.",
    "Mandarin Chinese": "Write your entire response in Simplified Chinese (简体中文).",
    "Hindi": "Write your entire response in Hindi (हिन्दी).",
    "Nepali": "Write your entire response in Nepali (नेपाली). Nepali is written in the Devanagari script.",
    "Bengali": "Write your entire response in Bengali (বাংলা).",
    "Urdu": "Write your entire response in Urdu (اردو).",
    "Punjabi": "Write your entire response in Punjabi (ਪੰਜਾਬੀ).",
    "Turkish": "Write your entire response in Turkish (Türkçe).",
    "Vietnamese": "Write your entire response in Vietnamese (Tiếng Việt).",
    "Tagalog": "Write your entire response in Tagalog (Filipino).",
    "Korean": "Write your entire response in Korean (한국어).",
    "Japanese": "Write your entire response in Japanese (日本語).",
    "Swahili": "Write your entire response in Swahili (Kiswahili).",
    "Somali": "Write your entire response in Somali (Soomaali).",
    "Welsh": "Write your entire response in Welsh (Cymraeg).",
    "Greek": "Write your entire response in Greek (Ελληνικά).",
    "Hungarian": "Write your entire response in Hungarian (Magyar).",
    "Czech": "Write your entire response in Czech (Čeština).",
    "Slovak": "Write your entire response in Slovak (Slovenčina).",
    "Ukrainian": "Write your entire response in Ukrainian (Українська).",
    "Russian": "Write your entire response in Russian (Русский).",
}


def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    """Extract plain text from PDF or docx."""
    fn = filename.lower()
    if fn.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(file_bytes))
        return "\n\n".join(page.extract_text() or "" for page in reader.pages)
    elif fn.endswith(".docx"):
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join(p.text for p in doc.paragraphs)
    else:
        raise ValueError("Unsupported file type. Please upload a PDF or .docx file.")


def simplify_with_claude(report_text: str, language: str) -> str:
    """Call Claude API to produce a simplified summary."""
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not set. Add it to your .env file.")

    lang_instr = LANGUAGE_INSTRUCTION.get(language, f"Write your entire response in {language}.")
    prompt = SIMPLIFY_PROMPT.format(
        language_instruction=lang_instr,
        report_text=report_text[:15000],  # stay within context limits
    )

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def md_to_rl(text: str, latin_only: bool = True) -> str:
    """Convert **bold** markdown to ReportLab <b> tags safely, strip emoji and escape XML."""
    if latin_only:
        # Strip anything outside basic Latin + Latin Extended (Helvetica limitation)
        text = re.sub(r'[^\x00-\x7FÀ-ɏ£€]', '', text)
    else:
        # Non-Latin font in use — only strip emoji/symbols that no font renders cleanly
        text = _EMOJI_RE.sub('', text)
    # Escape ampersands
    text = text.replace("&", "&amp;")
    # Convert **bold** pairs — non-greedy
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    # Strip any leftover lone **
    text = text.replace("**", "")
    return text.strip()


def build_pdf(simplified_text: str, language: str, a11y: dict | None = None) -> bytes:  # noqa: C901
    """Render the simplified text into a clean, branded PDF."""
    a11y = a11y or {}

    # ── Accessibility settings ──────────────────────────────────────────────
    high_contrast = a11y.get("hc", False)
    dyslexic      = a11y.get("dyslexic", False)
    colour_blind  = a11y.get("cb", False)
    text_size     = a11y.get("size", "normal")   # "normal" | "lg" | "xl"

    size_factor = {"normal": 1.0, "lg": 1.15, "xl": 1.3}.get(text_size, 1.0)

    def fs(base: float) -> float:
        """Scale font size by accessibility factor."""
        return round(base * size_factor, 1)

    # Brand colour from uploaded logo (used when no a11y override is active)
    logo_colors_data = load_logo_colors()
    brand_primary = logo_colors_data.get("primary", "#1e3a5f")

    # Colour palette
    if high_contrast:
        col_heading  = colors.black
        col_body     = colors.black
        col_muted    = colors.black
        col_rule     = colors.black
    elif colour_blind:
        col_heading  = colors.HexColor("#0072B2")   # CB-safe blue
        col_body     = colors.HexColor("#111827")
        col_muted    = colors.HexColor("#555555")
        col_rule     = colors.HexColor("#0072B2")
    else:
        col_heading  = colors.HexColor(brand_primary)
        col_body     = colors.HexColor("#111827")
        col_muted    = colors.HexColor("#6b7280")
        col_rule     = colors.HexColor(brand_primary)

    # Font — resolve per language, dyslexia mode adds extra spacing
    body_font, bold_font = ensure_noto_font(language)
    latin_only   = _script_family(language) is None
    leading_mult = 1.8 if dyslexic else 1.5
    word_space   = 2 if dyslexic else 0

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "Title",
        parent=styles["Normal"],
        fontSize=fs(22),
        leading=fs(22) * leading_mult,
        textColor=col_heading,
        spaceAfter=2,
        fontName=bold_font,
        alignment=TA_LEFT,
        wordSpace=word_space,
    )
    subtitle_style = ParagraphStyle(
        "Subtitle",
        parent=styles["Normal"],
        fontSize=fs(11),
        leading=fs(11) * leading_mult,
        textColor=col_muted,
        spaceAfter=4,
        fontName=body_font,
        alignment=TA_LEFT,
        wordSpace=word_space,
    )
    date_style = ParagraphStyle(
        "Date",
        parent=styles["Normal"],
        fontSize=fs(10),
        leading=fs(10) * leading_mult,
        textColor=col_muted,
        spaceAfter=12,
        fontName=body_font,
        alignment=TA_LEFT,
    )
    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Normal"],
        fontSize=fs(13),
        leading=fs(13) * leading_mult,
        textColor=col_heading,
        spaceBefore=14,
        spaceAfter=6,
        fontName=bold_font,
        keepWithNext=True,
        wordSpace=word_space,
    )
    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontSize=fs(10),
        leading=fs(10) * leading_mult,
        textColor=col_body,
        spaceAfter=5,
        fontName=body_font,
        wordSpace=word_space,
    )
    bullet_style = ParagraphStyle(
        "Bullet",
        parent=body_style,
        leftIndent=14,
        bulletIndent=4,
        spaceAfter=1,
    )

    story = []

    # Logo (if uploaded)
    story.append(Spacer(1, 4 * mm))
    if LOGO_PATH.exists():
        ct = logo_content_type() or "image/png"
        # SVG not directly supported by reportlab — skip gracefully
        if ct != "image/svg+xml":
            from reportlab.platypus import Image as RLImage
            logo_img = RLImage(str(LOGO_PATH), width=70 * mm, height=25 * mm, kind="proportional")
            logo_img.hAlign = "RIGHT"
            story.append(logo_img)
            story.append(Spacer(1, 4 * mm))

    # Header
    report_date = datetime.utcnow().strftime("%d %B %Y")
    story.append(Paragraph("Your Mortgage Summary", title_style))
    story.append(Paragraph(f"Simplified for you · {language}", subtitle_style))
    story.append(Paragraph(report_date, date_style))
    story.append(HRFlowable(width="100%", thickness=2, color=col_rule, spaceAfter=14))

    # Parse markdown-ish output from Claude
    lines = simplified_text.split("\n")
    for line in lines:
        stripped = line.strip()
        if not stripped:
            story.append(Spacer(1, 3 * mm))
            continue

        if re.match(r'^-{2,}$', stripped):
            continue  # skip markdown horizontal rules

        if stripped.startswith("## "):
            heading_text = stripped[3:].strip()
            story.append(CondPageBreak(70 * mm))
            story.append(Paragraph(heading_text, heading_style))

        elif stripped.startswith("# "):
            story.append(CondPageBreak(70 * mm))
            story.append(Paragraph(stripped[2:].strip(), heading_style))

        elif stripped.startswith(("- ", "• ", "* ")):
            bullet_text = md_to_rl(stripped[2:].strip(), latin_only=latin_only)
            story.append(Paragraph(f"• {bullet_text}", bullet_style))

        else:
            story.append(Paragraph(md_to_rl(stripped, latin_only=latin_only), body_style))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ── Endpoints ────────────────────────────────────────────────────────────────

@app.post("/simplify")
async def simplify(
    file: UploadFile = File(...),
    language: str = Form(default="English"),
    a11y: str = Form(default="{}"),
):
    """Upload a mortgage suitability report, get a simplified PDF back."""
    content = await file.read()

    try:
        report_text = extract_text_from_file(content, file.filename)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    if len(report_text.strip()) < 100:
        raise HTTPException(status_code=400, detail="Could not extract enough text from the file. Is it a scanned image PDF? Try a text-based PDF or Word document.")

    try:
        a11y_prefs = json.loads(a11y)
    except Exception:
        a11y_prefs = {}

    simplified = simplify_with_claude(report_text, language)
    pdf_bytes = build_pdf(simplified, language, a11y_prefs)
    increment_usage()

    safe_lang = language.replace(" ", "_")
    out_name = f"mortgage_summary_{safe_lang}.pdf"

    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )


@app.post("/logo")
async def upload_logo(file: UploadFile = File(...)):
    """Upload a brand logo (PNG, JPG, SVG, WebP). Replaces any existing logo."""
    ct = file.content_type or ""
    if ct not in ALLOWED_IMAGE_TYPES:
        raise HTTPException(status_code=400, detail="Unsupported image type. Use PNG, JPG, SVG, or WebP.")
    data = await file.read()
    if len(data) > 2 * 1024 * 1024:  # 2 MB limit
        raise HTTPException(status_code=400, detail="Logo must be under 2 MB.")
    save_logo(data, ct)
    colors_data = extract_logo_colors(data, ct)
    LOGO_COLORS_FILE.write_text(json.dumps(colors_data))
    return {"status": "ok", "content_type": ct, "brand_color": colors_data.get("primary")}


@app.get("/logo")
def get_logo():
    """Serve the current logo."""
    if not LOGO_PATH.exists():
        raise HTTPException(status_code=404, detail="No logo uploaded.")
    ct = logo_content_type() or "image/png"
    return Response(content=LOGO_PATH.read_bytes(), media_type=ct)


@app.delete("/logo")
def remove_logo():
    """Remove the current logo."""
    delete_logo()
    return {"status": "removed"}


@app.get("/stats")
def stats():
    """Return usage statistics."""
    return get_stats()


@app.get("/", response_class=HTMLResponse)
def index():
    html_path = Path("index.html")
    if html_path.exists():
        return HTMLResponse(html_path.read_text())
    return HTMLResponse("<h1>Place index.html next to main.py</h1>")
