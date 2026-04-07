# Acme & Dorf — Door Tag Generator

Web app: upload an engineering listing PDF → download a printable DOCX of door tags.

## Run locally (fastest)

```bash
# 1. Install system dependency (for rendering PDF pages as images)
#    Mac:
brew install poppler
#    Ubuntu/Debian:
sudo apt install poppler-utils
#    Windows: download from https://github.com/oschwartz10612/poppler-windows/releases

# 2. Install Python packages
pip install -r requirements.txt

# 3. Set your API key
export ANTHROPIC_API_KEY=sk-ant-api03-...     # Mac/Linux
set ANTHROPIC_API_KEY=sk-ant-api03-...         # Windows

# 4. Run
python app.py
# Open http://localhost:5000
```

---

## Deploy to Render.com (free, permanent URL)

1. Push this folder to a GitHub repo
2. Go to https://render.com → New → Web Service → connect your repo
3. Set these:
   - **Build Command:** `pip install -r requirements.txt && apt-get install -y poppler-utils`
   - **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --timeout 120 --workers 2`
4. Add environment variable: `ANTHROPIC_API_KEY = sk-ant-api03-...`
5. Deploy — you'll get a public URL like `https://door-tags.onrender.com`

---

## Deploy to Railway.app

1. Push to GitHub
2. Go to https://railway.app → New Project → Deploy from GitHub
3. Add environment variable: `ANTHROPIC_API_KEY`
4. In Settings → Deploy → Nixpacks: add `poppler-utils` to packages
5. Start command: `gunicorn app:app --bind 0.0.0.0:$PORT --timeout 120`

---

## How it works

| Step | What happens |
|------|-------------|
| Upload | PDF saved to temp file |
| Extract | Pages 1-2: text extracted; Pages 3+: rendered as images |
| AI | Claude reads images (handles handwriting) + order text → JSON |
| Generate | python-docx creates 3-column tag table matching T1.doc format |
| Download | DOCX returned automatically |

## Files

```
doortags_app/
├── app.py              ← Flask server + all logic
├── templates/
│   └── index.html      ← Upload UI
├── requirements.txt    ← Python deps
├── Procfile            ← For Heroku/Render
├── render.yaml         ← One-click Render config
└── README.md
```

## Environment variables

| Variable | Required | Description |
|----------|----------|-------------|
| `ANTHROPIC_API_KEY` | Yes | Your Claude API key. Can also be entered per-request in the UI. |
| `PORT` | No | Port to listen on (default 5000). Render/Railway set this automatically. |
