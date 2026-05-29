"""
Mortgage Suitability Report Simplifier
Run: uvicorn main:app --reload
"""

import io
import os
import json
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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
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
ALLOWED_IMAGE_TYPES = {"image/png", "image/jpeg", "image/svg+xml", "image/webp"}


# ── Logo helpers ──────────────────────────────────────────────────────────────

def save_logo(data: bytes, content_type: str):
    LOGO_PATH.write_bytes(data)
    LOGO_META_FILE.write_text(json.dumps({"content_type": content_type}))

def delete_logo():
    if LOGO_PATH.exists():
        LOGO_PATH.unlink()
    if LOGO_META_FILE.exists():
        LOGO_META_FILE.unlink()

def logo_content_type() -> str | None:
    if LOGO_META_FILE.exists():
        return json.loads(LOGO_META_FILE.read_text()).get("content_type")
    return None
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


def build_pdf(simplified_text: str, language: str) -> bytes:  # noqa: C901
    """Render the simplified text into a clean, branded PDF."""
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
        fontSize=22,
        leading=28,
        textColor=colors.HexColor("#1e3a5f"),
        spaceAfter=4,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
    )
    subtitle_style = ParagraphStyle(
        "Subtitle",
        parent=styles["Normal"],
        fontSize=11,
        textColor=colors.HexColor("#6b7280"),
        spaceAfter=16,
        fontName="Helvetica",
        alignment=TA_CENTER,
    )
    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Normal"],
        fontSize=13,
        leading=18,
        textColor=colors.HexColor("#1e3a5f"),
        spaceBefore=14,
        spaceAfter=6,
        fontName="Helvetica-Bold",
    )
    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontSize=10,
        leading=15,
        textColor=colors.HexColor("#111827"),
        spaceAfter=5,
        fontName="Helvetica",
    )
    bullet_style = ParagraphStyle(
        "Bullet",
        parent=body_style,
        leftIndent=14,
        bulletIndent=4,
        spaceAfter=4,
    )
    footer_style = ParagraphStyle(
        "Footer",
        parent=styles["Normal"],
        fontSize=8,
        textColor=colors.HexColor("#9ca3af"),
        alignment=TA_CENTER,
        fontName="Helvetica",
    )

    story = []

    # Logo (if uploaded)
    story.append(Spacer(1, 4 * mm))
    if LOGO_PATH.exists():
        ct = logo_content_type() or "image/png"
        # SVG not directly supported by reportlab — skip gracefully
        if ct != "image/svg+xml":
            from reportlab.platypus import Image as RLImage
            logo_img = RLImage(str(LOGO_PATH), width=40 * mm, height=15 * mm, kind="proportional")
            logo_img.hAlign = "CENTER"
            story.append(logo_img)
            story.append(Spacer(1, 3 * mm))

    # Header
    story.append(Paragraph("Your Mortgage Summary", title_style))
    story.append(Paragraph(f"Simplified for you · {language}", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#1e3a5f"), spaceAfter=14))

    # Parse markdown-ish output from Claude
    lines = simplified_text.split("\n")
    for line in lines:
        stripped = line.strip()
        if not stripped:
            story.append(Spacer(1, 3 * mm))
            continue

        if stripped.startswith("## "):
            heading_text = stripped[3:].strip()
            story.append(Paragraph(heading_text, heading_style))
            story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#e5e7eb"), spaceAfter=4))

        elif stripped.startswith("# "):
            story.append(Paragraph(stripped[2:].strip(), heading_style))

        elif stripped.startswith(("- ", "• ", "* ")):
            bullet_text = stripped[2:].strip()
            # Convert any **bold** markers for reportlab
            bullet_text = bullet_text.replace("**", "<b>", 1)
            while "**" in bullet_text:
                bullet_text = bullet_text.replace("**", "</b>", 1)
            story.append(Paragraph(f"• {bullet_text}", bullet_style))

        else:
            # Regular paragraph — handle inline bold
            text = stripped.replace("**", "<b>", 1)
            while "**" in text:
                text = text.replace("**", "</b>", 1)
            story.append(Paragraph(text, body_style))

    # Footer
    story.append(Spacer(1, 8 * mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#e5e7eb"), spaceAfter=6))
    story.append(Paragraph(
        "This simplified summary is for informational purposes only and does not constitute financial advice. "
        "Please refer to your full Mortgage Suitability Report for complete details.",
        footer_style,
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ── Endpoints ────────────────────────────────────────────────────────────────

@app.post("/simplify")
async def simplify(
    file: UploadFile = File(...),
    language: str = Form(default="English"),
):
    """Upload a mortgage suitability report, get a simplified PDF back."""
    content = await file.read()

    try:
        report_text = extract_text_from_file(content, file.filename)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    if len(report_text.strip()) < 100:
        raise HTTPException(status_code=400, detail="Could not extract enough text from the file. Is it a scanned image PDF? Try a text-based PDF or Word document.")

    simplified = simplify_with_claude(report_text, language)
    pdf_bytes = build_pdf(simplified, language)
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
    return {"status": "ok", "content_type": ct}


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
