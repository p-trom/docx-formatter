"""Formatting endpoints."""

from fastapi import APIRouter, File, Form, HTTPException, UploadFile, status
from fastapi.responses import FileResponse

from docx_formatter.api.models import DebugResponse, FormatResponse
from docx_formatter.api.service import process_document, validate_file_extension
from docx_formatter.config import settings

router = APIRouter(tags=["format"])


@router.post("/format/template-upload", response_model=FormatResponse)
async def format_template_upload(
    template: UploadFile = File(..., description="Template DOCX file containing styles"),
    content: UploadFile = File(..., description="Content DOCX file with raw text"),
    output_filename: str | None = Form(
        None, description="Optional output filename (must end with .docx)"
    ),
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


@router.post("/format/template-upload/debug", response_model=DebugResponse)
async def format_template_upload_debug(
    template: UploadFile = File(..., description="Template DOCX file containing styles"),
    content: UploadFile = File(..., description="Content DOCX file with raw text"),
):
    """
    Debug endpoint: format documents and return detailed matching logs.

    Instead of returning the formatted file, this endpoint returns a JSON
    response showing:
    - How many styles were found in the template
    - How many paragraphs were in the content
    - Which matching passes were used (exact, fuzzy, semantic, LLM, heuristic)
    - Per-match details with confidence scores and reasons
    - Whether LLM was available and used
    - Any styles that remained unmatched
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
    )

    if not result.success:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail={"message": "Formatting failed", "warnings": result.warnings},
        )

    # Build debug response from processing_log
    log = result.processing_log
    if log is None:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Processing log not available",
        )

    # Clean up output file (we don't return it in debug mode)
    if output_path.exists():
        output_path.unlink(missing_ok=True)

    return DebugResponse(
        success=result.success,
        processing_time_ms=elapsed_ms,
        template_styles_found=log.template_styles_found,
        content_paragraphs=log.content_paragraphs,
        llm_available=log.llm_available,
        llm_used=log.llm_used,
        total_matches=len(log.match_logs),
        matches=[
            {
                "pass": m.pass_name,
                "source_style": m.source_style,
                "target_style": m.target_style,
                "confidence": round(m.confidence, 2),
                "reason": m.reason,
                "paragraph_preview": m.paragraph_preview,
            }
            for m in log.match_logs
        ],
        unmatched=log.unmatched_after_all_passes,
        warnings=result.warnings,
    )
