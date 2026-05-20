"""Formatting endpoints."""

import tempfile
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile, status
from fastapi.responses import FileResponse

from docx_formatter.api.models import FormatResponse
from docx_formatter.api.service import process_document, validate_file_extension
from docx_formatter.config import settings

router = APIRouter(tags=["format"])


@router.post("/format/template-upload", response_model=FormatResponse)
async def format_template_upload(
    template: UploadFile = File(..., description="Template DOCX file containing styles"),
    content: UploadFile = File(..., description="Content DOCX file with raw text"),
    output_filename: str | None = Form(None, description="Optional output filename (must end with .docx)"),
):
    """
    Format a content document using a template document.

    Upload two DOCX files:
    - **template**: The document containing desired styles/branding
    - **content**: The document with raw text to be reformatted

    Returns the formatted DOCX file as a download.
    """
    # Validate input files
    if not template.filename or not validate_file_extension(template.filename):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Template must be a .docx file",
        )
    if not content.filename or not validate_file_extension(content.filename):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Content must be a .docx file",
        )

    if output_filename and not output_filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Output filename must end with .docx",
        )

    # Save uploaded files to temp directory
    temp_dir = settings.temp_dir
    temp_dir.mkdir(parents=True, exist_ok=True)

    template_path = temp_dir / f"template_{template.filename}"
    content_path = temp_dir / f"content_{content.filename}"

    try:
        template_path.write_bytes(await template.read())
        content_path.write_bytes(await content.read())
    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to save uploaded files: {exc}",
        ) from exc

    # Process
    output_path, result, elapsed_ms = await process_document(
        template_path=template_path,
        content_path=content_path,
        output_filename=output_filename,
    )

    if not result.success:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail={"message": "Formatting failed", "warnings": result.warnings},
        )

    # Return file as download
    return FileResponse(
        path=str(output_path),
        filename=output_filename or output_path.name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "X-Processing-Time-Ms": str(elapsed_ms),
            "X-Warnings": "; ".join(result.warnings) if result.warnings else "none",
        },
    )
