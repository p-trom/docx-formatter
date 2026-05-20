#! /usr/bin/env bash
# Start script for Render (and local dev)

set -e

PORT="${PORT:-8000}"
WORKERS="${WORKERS:-2}"

echo "Starting DOCX Formatter API on port $PORT with $WORKERS workers..."

exec uvicorn docx_formatter.api.main:app \
    --host 0.0.0.0 \
    --port "$PORT" \
    --workers "$WORKERS" \
    --access-log \
    --proxy-headers
