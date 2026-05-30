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

EXTRA_PAGES_PATH = Path("extra_pages.pdf")
EXTRA_PAGES_META_FILE = Path("extra_pages_meta.json")

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


# ── Extra pages helpers ───────────────────────────────────────────────────────

def save_extra_pages(data: bytes, filename: str, page_count: int):
    EXTRA_PAGES_PATH.write_bytes(data)
    EXTRA_PAGES_META_FILE.write_text(json.dumps({"filename": filename, "pages": page_count}))

def delete_extra_pages():
    if EXTRA_PAGES_PATH.exists():
        EXTRA_PAGES_PATH.unlink()
    if EXTRA_PAGES_META_FILE.exists():
        EXTRA_PAGES_META_FILE.unlink()

def extra_pages_meta() -> dict | None:
    if EXTRA_PAGES_META_FILE.exists():
        return json.loads(EXTRA_PAGES_META_FILE.read_text())
    return None

def merge_pdfs(main_bytes: bytes, extra_bytes: bytes) -> bytes:
    """Append extra_bytes pages onto main_bytes and return merged PDF."""
    from pypdf import PdfWriter
    writer = PdfWriter()
    writer.append(io.BytesIO(main_bytes))
    writer.append(io.BytesIO(extra_bytes))
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


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

# ── Next step options ─────────────────────────────────────────────────────────

NEXT_STEP_OPTIONS = [
    {
        "id": "confirm",
        "template": "Please make sure the contents of this report match your understanding and ask {adviser_name} if you have any questions.",
        "uses_adviser": True,
    },
    {
        "id": "illustration",
        "template": "Read the mortgage illustration that came with this report for more information.",
        "uses_adviser": False,
    },
    {
        "id": "insurance",
        "template": "Arrange buildings and contents insurance.",
        "uses_adviser": False,
    },
    {
        "id": "meeting",
        "template": "Arrange a meeting with {adviser_name} to discuss how this new mortgage impacts your current life assurance and income protection requirements.",
        "uses_adviser": True,
    },
    {
        "id": "will",
        "template": "Update your will as your circumstances are changing.",
        "uses_adviser": False,
    },
    {
        "id": "protection",
        "template": "Please read the attached information about protecting your income.",
        "uses_adviser": False,
    },
]


def build_next_steps(selected_ids: list, adviser_name: str, custom_note: str = "") -> list:
    """Return list of plain-text next step strings from selected option IDs."""
    steps = []
    for opt in NEXT_STEP_OPTIONS:
        if opt["id"] in selected_ids:
            text = opt["template"]
            if opt.get("uses_adviser"):
                text = text.format(adviser_name=adviser_name or "your adviser")
            steps.append(text)
    if custom_note.strip():
        steps.append(custom_note.strip())
    return steps


SIMPLIFY_PROMPT = """You are an expert at making complex mortgage suitability reports easy to understand for everyday people.

You will be given a mortgage suitability report. Your job is to produce a simplified, plain-English summary.

First, on the very first line of your response write exactly this (replace with the actual name found in the report):
ADVISER: [Full name of the mortgage adviser or broker from the report. Write "your adviser" if not found.]

Then on the next line start the sections.

The summary should:
- Use simple, everyday language (no jargon)
- Be warm, reassuring, and easy to read
- Avoid anything that sounds like a legal document
- Write in short, clear sentences and paragraphs — do NOT use bullet points anywhere
- Use **bold** (double asterisks) to highlight key facts and figures, e.g. **Halifax**, **4.89%**, **£1,245 per month**, **25 years**

Structure your response with these EXACT section headings (use markdown ## for headings):

## Your New Mortgage — The Facts
A short, plain-English overview: lender name, mortgage type (e.g. fixed rate, tracker), and what it means for the client in one or two sentences.

## About You
A brief summary of the client's circumstances, needs, and preferences that shaped this recommendation (e.g. first-time buyer, remortgage, income, dependants, priorities).

## Why We Chose This Lender for You
2–3 sentences explaining specifically why this lender and product were selected — criteria met, flexibility, rate competitiveness, any unique features.

## The Numbers
Key financial figures in bullet points: property value, mortgage amount, loan-to-value (LTV), deposit, and total amount repayable over the full term.

## The Initial Scheme
Details of the initial deal period only: the interest rate, type (fixed/tracker/discount), how long it lasts, and the monthly payment during this period.

## The Overall Term
The full mortgage term length, what happens after the initial scheme ends (e.g. reverts to SVR), and the monthly payment on the standard rate if known.

## Key Things to Know
Important conditions, risks, or caveats: early repayment charges, portability, overpayment allowances, anything the client must not overlook.

## What Happens Next
Write only: [NEXT_STEPS_PLACEHOLDER]

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


def simplify_with_claude(report_text: str, language: str) -> tuple[str, str]:
    """Call Claude API to produce a simplified summary.
    Returns (simplified_text, adviser_name).
    """
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
    raw = message.content[0].text

    # Extract ADVISER: line from the first line
    adviser_name = "your adviser"
    lines = raw.splitlines()
    if lines and lines[0].startswith("ADVISER:"):
        adviser_name = lines[0][len("ADVISER:"):].strip() or "your adviser"
        raw = "\n".join(lines[1:]).lstrip("\n")

    return raw, adviser_name


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


def build_pdf(simplified_text: str, language: str, a11y: dict | None = None, custom_next_steps: list | None = None) -> bytes:  # noqa: C901
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

        if "[NEXT_STEPS_PLACEHOLDER]" in stripped:
            continue  # will be replaced by custom next steps below

        if stripped.startswith("## "):
            heading_text = stripped[3:].strip()
            # Skip Claude's final heading when we're supplying our own next steps
            if heading_text == "What Happens Next" and custom_next_steps is not None:
                continue
            story.append(CondPageBreak(70 * mm))
            story.append(Paragraph(heading_text, heading_style))

        elif stripped.startswith("# "):
            story.append(CondPageBreak(70 * mm))
            story.append(Paragraph(stripped[2:].strip(), heading_style))

        elif stripped.startswith(("- ", "• ", "* ")):
            # Render as plain body paragraph (no bullet point)
            story.append(Paragraph(md_to_rl(stripped[2:].strip(), latin_only=latin_only), body_style))

        else:
            story.append(Paragraph(md_to_rl(stripped, latin_only=latin_only), body_style))

    # Append custom next steps if provided
    if custom_next_steps:
        story.append(CondPageBreak(70 * mm))
        story.append(Paragraph("What Happens Next", heading_style))
        for step in custom_next_steps:
            story.append(Paragraph(md_to_rl(step, latin_only=latin_only), body_style))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ── Endpoints ────────────────────────────────────────────────────────────────

@app.post("/simplify")
async def simplify(
    file: UploadFile = File(...),
    language: str = Form(default="English"),
    a11y: str = Form(default="{}"),
    next_steps: str = Form(default="[]"),
    custom_note: str = Form(default=""),
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

    try:
        selected_ids = json.loads(next_steps)
    except Exception:
        selected_ids = []

    simplified, adviser_name = simplify_with_claude(report_text, language)
    custom_steps = build_next_steps(selected_ids, adviser_name, custom_note) if selected_ids or custom_note.strip() else None
    pdf_bytes = build_pdf(simplified, language, a11y_prefs, custom_next_steps=custom_steps)

    # Append extra pages if uploaded
    if EXTRA_PAGES_PATH.exists():
        pdf_bytes = merge_pdfs(pdf_bytes, EXTRA_PAGES_PATH.read_bytes())

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


@app.post("/pages")
async def upload_extra_pages(file: UploadFile = File(...)):
    """Upload a PDF to append to every generated report."""
    if not (file.filename or "").lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are supported.")
    data = await file.read()
    if len(data) > 20 * 1024 * 1024:  # 20 MB limit
        raise HTTPException(status_code=400, detail="File must be under 20 MB.")
    try:
        reader = PdfReader(io.BytesIO(data))
        page_count = len(reader.pages)
    except Exception:
        raise HTTPException(status_code=400, detail="Could not read the PDF. Please check the file is valid.")
    save_extra_pages(data, file.filename, page_count)
    return {"status": "ok", "pages": page_count, "filename": file.filename}


@app.get("/pages/meta")
def get_extra_pages_meta():
    """Return metadata about the uploaded extra pages."""
    meta = extra_pages_meta()
    if not meta:
        raise HTTPException(status_code=404, detail="No extra pages uploaded.")
    return meta


@app.delete("/pages")
def remove_extra_pages():
    """Remove the uploaded extra pages."""
    delete_extra_pages()
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
