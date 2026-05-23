"""API application entrypoint."""

from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from docx_formatter.api.routers import format_router
from docx_formatter.config import settings


def cleanup_temp_files():
    """Clean up leftover temp files from previous runs."""
    if settings.temp_dir.exists():
        for path in settings.temp_dir.iterdir():
            try:
                if path.is_file():
                    path.unlink()
                elif path.is_dir():
                    import shutil

                    shutil.rmtree(path)
            except OSError:
                pass


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan handler."""
    settings.temp_dir.mkdir(parents=True, exist_ok=True)
    cleanup_temp_files()
    yield
    cleanup_temp_files()


app = FastAPI(
    title="DOCX Formatter API",
    description="API for intelligent DOCX formatting based on templates.",
    version="0.1.0",
    docs_url="/docs",
    redoc_url="/redoc",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(format_router, prefix="/api/v1")


@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {"status": "ok", "version": "0.1.0"}


@app.get("/")
async def root():
    """Root redirect to API documentation."""
    return {
        "message": "DOCX Formatter API",
        "version": "0.1.0",
        "docs": "/docs",
        "health": "/health",
        "format": "POST /api/v1/format/template-upload",
    }


@app.exception_handler(Exception)
async def general_exception_handler(request, exc):
    """Catch-all exception handler."""
    return JSONResponse(
        status_code=500,
        content={"detail": "Internal server error", "message": str(exc)},
    )
