"""Business logic layer for the API."""

import shutil
import time
import uuid
from pathlib import Path
from typing import Optional

from docx_formatter.config import settings
from docx_formatter.core.pipeline import FormatPipeline
from docx_formatter.core.types import ProcessingResult


async def process_document(
    template_path: Path,
    content_path: Path,
    output_filename: Optional[str] = None,
) -> tuple[Path, ProcessingResult, int]:
    """
    Process a document formatting request.

    Args:
        template_path: Path to the template DOCX file
        content_path: Path to the content DOCX file
        output_filename: Optional custom output filename

    Returns:
        Tuple of (output_path, result, processing_time_ms)
    """
    start = time.perf_counter()
    job_id = uuid.uuid4().hex[:12]
    output_path = settings.temp_dir / (output_filename or f"output_{job_id}.docx")

    pipeline = FormatPipeline()
    result = pipeline.process(
        template_path=str(template_path),
        content_path=str(content_path),
        output_path=str(output_path),
    )

    elapsed_ms = int((time.perf_counter() - start) * 1000)

    if not result.success and output_path.exists():
        output_path.unlink(missing_ok=True)

    return output_path, result, elapsed_ms


def validate_file_extension(filename: str) -> bool:
    """Check if filename has allowed extension."""
    return any(filename.lower().endswith(ext) for ext in settings.allowed_extensions)


def save_upload_file(source: Path, destination: Path) -> Path:
    """Move or copy uploaded file to destination."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source == destination:
        return destination
    shutil.copy2(str(source), str(destination))
    return destination
