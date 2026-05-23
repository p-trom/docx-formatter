"""
Format Pipeline - main entry point for DOCX formatting.
Orchestrates extraction, matching, and assembly.
"""

import logging
import time

from docx_formatter.config import settings

from .assembler import DocumentAssembler
from .extractor import DOCXExtractor
from .matcher import StyleMatchingEngine
from .types import ProcessingResult

logger = logging.getLogger(__name__)


class FormatPipeline:
    """
    Main pipeline for formatting documents.

    Usage:
        pipeline = FormatPipeline()
        result = pipeline.process("template.docx", "content.docx", "output.docx")
    """

    def __init__(self, min_match_confidence: float = 0.5, use_llm: bool = True):
        self.extractor = DOCXExtractor()
        self.matcher = StyleMatchingEngine(
            min_confidence=min_match_confidence,
            use_llm=use_llm and bool(settings.openrouter_api_key),
        )
        self.assembler = DocumentAssembler()

    def process(
        self,
        template_path: str,
        content_path: str,
        output_path: str,
    ) -> ProcessingResult:
        """
        Process a template and content document into a formatted output.

        Args:
            template_path: Path to template DOCX (Document A)
            content_path: Path to content DOCX (Document B)
            output_path: Path to save output DOCX (Document C)

        Returns:
            ProcessingResult with status and metadata
        """
        start_time = time.time()

        logger.info(f"Starting format pipeline: {template_path} + {content_path} -> {output_path}")

        try:
            # Step 1: Extract template profile
            logger.info("Extracting template profile...")
            template_profile = self.extractor.extract_template_profile(template_path)
            logger.info(f"Template type: {template_profile.template_type.name}")

            # Step 2: Extract content profile
            logger.info("Extracting content profile...")
            content_profile = self.extractor.extract_content_profile(content_path)

            # Step 3: Match styles
            logger.info("Matching styles...")
            style_matches = self.matcher.match_all(template_profile, content_profile)

            # Step 4: Assemble output
            logger.info("Assembling output document...")
            result = self.assembler.assemble(
                template=template_profile,
                content=content_profile,
                style_matches=style_matches,
                output_path=output_path,
                template_docx_path=template_path,
            )

            result.processing_time_ms = int((time.time() - start_time) * 1000)
            result.processing_log = self.matcher.last_log

            logger.info(
                f"Processing complete in {result.processing_time_ms}ms. "
                f"Success: {result.success}, "
                f"Matched styles: {len(result.matched_styles)}"
            )

            return result

        except Exception as e:
            logger.error(f"Pipeline failed: {e}", exc_info=True)
            return ProcessingResult(
                success=False,
                warnings=[str(e)],
                processing_time_ms=int((time.time() - start_time) * 1000),
            )

    def process_files(
        self,
        template_bytes: bytes,
        content_bytes: bytes,
    ) -> ProcessingResult:
        """
        Process documents from bytes (for web/API usage).

        Saves to temp files, processes, returns result with output_bytes.
        """
        import os
        import tempfile

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            tf.write(template_bytes)
            template_path = tf.name

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as cf:
            cf.write(content_bytes)
            content_path = cf.name

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as of:
            output_path = of.name

        try:
            result = self.process(template_path, content_path, output_path)

            if result.success and result.output_path:
                with open(result.output_path, "rb") as f:
                    result.output_bytes = f.read()

            return result

        finally:
            # Cleanup temp files
            for path in [template_path, content_path, output_path]:
                try:
                    os.unlink(path)
                except OSError:
                    pass


__all__ = ["FormatPipeline"]
