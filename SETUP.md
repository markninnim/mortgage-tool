# Mortgage Summary Tool — Setup Guide

Upload a mortgage suitability report (PDF or Word), pick a language, and download a plain-English summary as a PDF.

---

## 1. Get your Anthropic API key

1. Go to [console.anthropic.com](https://console.anthropic.com) and sign up / log in
2. Click **API Keys** → **Create Key**
3. Copy the key — you'll use it in Step 3

---

## 2. Install dependencies

You need Python 3.9+. Then run:

```bash
pip install fastapi uvicorn python-multipart pypdf python-docx reportlab anthropic python-dotenv
```

---

## 3. Add your API key

In the `mortgage_app` folder, rename `.env.example` to `.env` and paste your key:

```
ANTHROPIC_API_KEY=sk-ant-...your key here...
```

---

## 4. Run the app

```bash
cd mortgage_app
uvicorn main:app --reload
```

Open **http://localhost:8000** in your browser.

---

## 5. Using the tool

1. Upload a mortgage suitability report (PDF or .docx)
2. Select the language for the output
3. Click **Generate Simplified Summary**
4. The simplified PDF downloads automatically (15–30 seconds)

---

## 6. Share with others

### Same Wi-Fi network
```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```
Others on the same network visit `http://YOUR_LOCAL_IP:8000`.

### On the internet (free hosting)
| Service | Steps |
|---------|-------|
| **Railway** | Push folder to GitHub → connect at railway.app → set `ANTHROPIC_API_KEY` in environment variables |
| **Render** | Same — free tier at render.com |
| **Fly.io** | `fly launch` from the project folder |

On any hosting service, set `ANTHROPIC_API_KEY` as an environment variable (don't commit your `.env` file).

---

## File structure

```
mortgage_app/
├── main.py          ← FastAPI backend + Claude integration
├── index.html       ← Frontend interface
├── .env             ← Your API key (create from .env.example)
├── .env.example     ← Template
└── SETUP.md         ← This file
```

---

## Notes

- Works best with **text-based PDFs** (not scanned images). If a scanned PDF doesn't work, try exporting the original report as a Word document first.
- The tool supports **29 languages** including English, Spanish, French, Arabic, Mandarin, Hindi, Welsh, and more.
- API costs are very low — roughly $0.01–0.03 per report.
