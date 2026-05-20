# DOCX Formatter

SaaS for intelligent DOCX document formatting.

## What it does

Takes a **template document** (with styles, branding, layout) and a **raw content document**, then produces a professionally formatted DOCX output.

## Examples
- Company offer template + raw notes → Branded offer
- CV template + plain text CV → Professional CV
- Report template + raw data → Styled report

## API Usage

### Health Check
```bash
curl https://docx-formatter-api.onrender.com/health
```

### Format Document
```bash
curl -X POST https://docx-formatter-api.onrender.com/api/v1/format/template-upload \
  -F "template=@template.docx" \
  -F "content=@content.docx" \
  -F "output_filename=formatted.docx" \
  -o formatted.docx
```

Response: `200 OK` with formatted DOCX file as attachment.

Interactive docs: `https://docx-formatter-api.onrender.com/docs`

## Web App

A drag & drop web interface for formatting documents:

**Live demo:** (deploy via Vercel below)

### Features
- Drag & drop or click to upload template + content DOCX files
- Real-time progress bar during processing
- Configurable API URL (point to your Render backend)
- Local history of formatted documents
- Fully responsive design

### Deploy Web App to Vercel (One-Click)

1. Go to [vercel.com](https://vercel.com) and log in with GitHub account
2. Click **"Add New..."** → **"Project"**
3. Import `p-trom/docx-formatter` from GitHub
4. Set **Root Directory** to `webapp`
5. Framework preset: **Other** (static site)
6. Click **Deploy** — będzie gotowe w ~30s

Your webapp will be at: `https://docx-formatter-web.vercel.app`

### Use Web App Locally

```bash
cd webapp
python3 -m http.server 8080
# Open http://localhost:8080
```

## Architecture

See `docs/architecture.md` for full technical specification.

## Quick Start

### Deploy to Render (One-Click)

The repository includes a `render.yaml` blueprint for instant deployment:

1. Go to [dashboard.render.com](https://dashboard.render.com) and log in with **ptrominski@gmail.com**
2. Click **"New +"** → **"Blueprint"**
3. Connect your GitHub account and select the `p-trom/docx-formatter` repository
4. Render will read `render.yaml` and create the web service automatically
5. Wait for deployment (2–3 min), then your API will be live at `https://docx-formatter-api.onrender.com`

### Local Development

#### Prerequisites
- Python 3.11+
- Docker & Docker Compose (optional)

#### Development Setup
```bash
# Clone
git clone https://github.com/p-trom/docx-formatter.git
cd docx-formatter

# Create virtual environment
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Run API locally
PYTHONPATH=src uvicorn docx_formatter.api.main:app --reload

# API docs: http://localhost:8000/docs
```

#### Run tests
```bash
pytest
```

#### Docker (production-like)
```bash
docker build -t docx-formatter .
docker run -p 8000:8000 docx-formatter
```

## Project Structure

```
docx-formatter/
├── src/
│   └── docx_formatter/
│       ├── core/          # DOCX parsing, analysis, matching, assembly
│       ├── api/           # FastAPI endpoints
│       ├── workers/       # Celery background tasks
│       └── utils/         # Shared utilities
├── webapp/            # Frontend web application
│   ├── index.html
│   ├── styles.css
│   └── app.js
├── tests/             # Test suite
├── docs/              # Architecture & documentation
├── docker/            # Docker configurations
└── scripts/           # Dev/build scripts
```

## License
MIT
