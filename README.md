# DOCX Formatter

**100% offline** intelligent DOCX document formatting. Pure Python document engineering with optional AI-enhanced style matching.

**Live API:** `https://docx-formatter-axh8.onrender.com`  
**Live Web App:** `https://docx-formatter.vercel.app`

---

## What it does

Takes a **template document** (with styles, branding, layout, headers/footers) and a **raw content document**, then produces a professionally formatted DOCX output where the content inherits the template's visual identity.

### Use cases
- Company offer template + raw notes → Branded offer
- CV template + plain text CV → Professional CV
- Report template + raw data → Styled report
- Any DOCX with defined styles + any DOCX with content → Formatted output

---

## Quick Start

### Web App (easiest)

1. Open [docx-formatter.vercel.app](https://docx-formatter.vercel.app)
2. Drag & drop your **template.docx** and **content.docx**
3. Click **Format Document**
4. Download the result

### API (cURL)

**Health check:**
```bash
curl https://docx-formatter-axh8.onrender.com/health
# → {"status":"ok","version":"0.1.0"}
```

**Format documents:**
```bash
curl -X POST https://docx-formatter-axh8.onrender.com/api/v1/format/template-upload \
  -F "template=@template.docx" \
  -F "content=@content.docx" \
  -F "output_filename=formatted.docx" \
  -o formatted.docx
```

**Interactive docs:** [docx-formatter-axh8.onrender.com/docs](https://docx-formatter-axh8.onrender.com/docs)

---

## Architecture

### Design principle: Template as base document

Unlike tools that create a fresh DOCX and copy styles, this formatter **opens the template DOCX as the base** and replaces content in-place. This preserves:

- Headers, footers, page numbers
- Background images and watermarks
- Section breaks and column layouts
- Custom XML parts (macros, document properties)
- Embedded fonts and themes

### Processing pipeline

```
Template DOCX ──┐
                ├──→ FormatPipeline ──→ Output DOCX
Content DOCX ───┘
```

**Stage 1 — Extract** (`DOCXExtractor`)
- Parse template: all paragraph styles, document defaults, theme, outline levels
- Parse content: paragraphs with runs (bold/italic/color), tables, structure tree
- Read `styles.xml` directly via lxml for complete style definitions

**Stage 2 — Match** (`StyleMatchingEngine`)
- Exact style ID match (e.g. `Heading 1` → `Heading 1`)
- Fuzzy name matching (Levenshtein similarity ≥ 0.6)
- Semantic role matching (title, heading, body, list, quote, caption)
- **LLM-based matching** (optional — via OpenRouter, for ambiguous styles)
- Content heuristic matching (font size, bold, indentation)

**Stage 3 — Assemble** (`DocumentAssembler`)
- Classify template sections: cover, TOC, frontmatter, content
- Split content into heading queue and body queue (preserves document order)
- Fill template slots from appropriate queues
- Preserve structural elements (chapter numbers, appendix labels)
- Append overflow content before the last section break

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Backend API | FastAPI + uvicorn |
| DOCX Engine | python-docx + lxml (OOXML) |
| AI Matching | OpenRouter API (optional, via httpx) |
| Frontend | Vanilla HTML/CSS/JS (static) |
| Backend Hosting | Render (free tier) |
| Frontend Hosting | Vercel (static) |
| Tests | pytest |
| Linting | ruff |
| CI/CD | GitHub Actions |

**Optional AI:** Set `OPENROUTER_API_KEY` to enable LLM-enhanced style matching for ambiguous documents. Zero impact on performance when not set — app works fully offline.

---

## Project Structure

```
docx-formatter/
├── src/docx_formatter/
│   ├── core/
│   │   ├── pipeline.py      # Main entry point: FormatPipeline
│   │   ├── extractor.py     # DOCX → TemplateProfile / ContentProfile
│   │   ├── matcher.py       # StyleMatchingEngine (5-pass cascade, optional LLM)
│   │   ├── llm_matcher.py   # OpenRouter-powered style matching (optional)
│   │   ├── assembler.py     # DocumentAssembler (section-aware replacement)
│   │   └── types.py         # Dataclasses: 18 semantic roles, style models
│   ├── api/
│   │   ├── main.py          # FastAPI app, CORS, lifespan
│   │   ├── models.py        # Pydantic request/response models
│   │   ├── service.py       # Business logic: process_document()
│   │   └── routers/
│   │       └── format.py    # POST /api/v1/format/template-upload
│   └── config.py            # Pydantic-settings (env vars)
├── webapp/                  # Static frontend (deployed to Vercel)
│   ├── index.html
│   ├── styles.css
│   └── app.js               # Drag & drop, fetch API, localStorage history
├── tests/
│   ├── test_api.py          # API integration tests (4 tests)
│   ├── test_core.py         # Core engine tests (17 tests)
│   └── fixtures/            # template_offer.docx, content_offer.docx, etc.
├── docs/
│   └── architecture.md      # Full SaaS architecture spec (v1-v3 roadmap)
├── requirements.txt
├── Dockerfile               # Multi-stage build (Python 3.12 slim)
├── render.yaml              # Render.com blueprint
├── vercel.json              # Vercel static deployment config
└── start.sh                 # Production startup script
```

---

## Local Development

### Prerequisites
- Python 3.11+
- uv (recommended) or pip

### Setup

```bash
# Clone
git clone https://github.com/p-trom/docx-formatter.git
cd docx-formatter

# Install dependencies
uv pip install -e ".[dev]"
# or: pip install -e ".[dev]"

# Run API
uvicorn docx_formatter.api.main:app --reload

# API docs: http://localhost:8000/docs
```

### Run tests

```bash
pytest -q
```

Current status: **30 tests passing** (4 API integration + 17 core engine + 9 LLM matcher).

### Docker

```bash
docker build -t docx-formatter .
docker run -p 8000:8000 docx-formatter
```

---

## Deployment

### Backend → Render

1. Push to `main` branch on GitHub
2. Render auto-deploys via `render.yaml` blueprint
3. Health check at `/health`

### Frontend → Vercel

1. Import `p-trom/docx-formatter` on Vercel
2. Set root directory to `webapp`
3. Framework: **Other** (static)
4. Deploy

The frontend calls the backend via configurable `apiUrl` (default: Render URL).

---

## Configuration

Environment variables (via `.env` or platform dashboard):

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `8000` | Server port |
| `MAX_FILE_SIZE_MB` | `16` | Max upload size |
| `CORS_ORIGINS` | `["*"]` | Allowed frontend origins |
| `OPENROUTER_API_KEY` | `None` | Optional: enable LLM-enhanced style matching |

---

## Roadmap

### Current (v0.1.0) ✅
- [x] Upload 2 DOCX files via API and web UI
- [x] Extract styles from template (paragraph + document defaults)
- [x] 5-pass style matching (exact → fuzzy → semantic → **LLM** → heuristic)
- [x] Optional AI-enhanced matching via OpenRouter (zero config = offline)
- [x] Section-aware assembly (preserve headers/footers/background)
- [x] Download formatted DOCX
- [x] Health endpoint
- [x] 30 tests

### v0.2 (planned)
- [ ] Table style matching and formatting
- [ ] Image preservation from content document
- [ ] Batch processing (multiple content files)
- [x] ~~pyproject.toml for proper packaging~~ ✅ Done

### v0.3 (planned)
- [ ] Template marketplace (pre-built templates)
- [ ] Placeholder resolution (`{{company_name}}`)
- [ ] PDF output option
- [ ] Rate limiting & usage tracking

---

## License

MIT
