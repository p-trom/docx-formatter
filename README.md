# DOCX Formatter

SaaS for intelligent DOCX document formatting.

## What it does

Takes a **template document** (with styles, branding, layout) and a **raw content document**, then produces a professionally formatted DOCX output.

## Examples
- Company offer template + raw notes → Branded offer
- CV template + plain text CV → Professional CV
- Report template + raw data → Styled report

## Architecture

See `docs/architecture.md` for full technical specification.

## Quick Start

### Prerequisites
- Python 3.11+
- Docker & Docker Compose

### Development Setup
```bash
# Clone
git clone https://github.com/p-trom/docx-formatter.git
cd docx-formatter

# Copy and fill environment
cp .env.example .env

# Start services
docker-compose -f docker/docker-compose.dev.yml up -d

# Install Python dependencies in venv
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Run tests
pytest
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
├── tests/             # Test suite
├── docs/              # Architecture & documentation
├── docker/            # Docker configurations
└── scripts/           # Dev/build scripts
```

## License
MIT
