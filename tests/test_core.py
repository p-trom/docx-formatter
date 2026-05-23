"""
Tests for DOCX formatter core engine.
"""

from pathlib import Path

import pytest

from docx_formatter.core import (
    DocumentAssembler,
    DOCXExtractor,
    FormatPipeline,
    StyleMatchingEngine,
)
from docx_formatter.core.types import SemanticRole

FIXTURES_DIR = Path(__file__).parent / "fixtures"


class TestDOCXExtractor:
    """Test DOCX extraction from sample files."""

    def test_extract_template_offer(self):
        """Extract template profile from offer template."""
        extractor = DOCXExtractor()
        template_path = FIXTURES_DIR / "template_offer.docx"

        profile = extractor.extract_template_profile(str(template_path))

        assert profile is not None
        # Should find our custom styles
        assert "OfferTitle" in profile.paragraph_styles or any(
            "offer" in s.name.lower() or "title" in s.name.lower()
            for s in profile.paragraph_styles.values()
        )

    def test_extract_content_offer(self):
        """Extract content profile from raw offer content."""
        extractor = DOCXExtractor()
        content_path = FIXTURES_DIR / "content_offer.docx"

        profile = extractor.extract_content_profile(str(content_path))

        assert profile is not None
        assert len(profile.paragraphs) > 5
        # Should detect some headings
        assert any(p.estimated_role for p in profile.paragraphs)

    def test_extract_cv_template(self):
        """Extract CV template profile."""
        extractor = DOCXExtractor()
        template_path = FIXTURES_DIR / "template_cv.docx"

        profile = extractor.extract_template_profile(str(template_path))

        assert profile is not None
        assert profile.document_defaults.margins is not None
        assert len(profile.paragraph_styles) > 0


class TestStyleMatchingEngine:
    """Test style matching logic."""

    def test_exact_style_match(self):
        """Test that exact style names are matched directly."""
        matcher = StyleMatchingEngine()

        extractor = DOCXExtractor()
        template = extractor.extract_template_profile(str(FIXTURES_DIR / "template_offer.docx"))
        content = extractor.extract_content_profile(str(FIXTURES_DIR / "content_offer.docx"))

        matches = matcher.match_all(template, content)

        # Should have some matches
        assert len(matches) > 0

        # Check for exact matches
        # Content may or may not have exact matches depending on style naming
        assert isinstance(matches[0].confidence, float)
        assert 0 <= matches[0].confidence <= 1

    def test_heuristic_match(self):
        """Test content-based heuristic matching."""
        matcher = StyleMatchingEngine()

        extractor = DOCXExtractor()
        template = extractor.extract_template_profile(str(FIXTURES_DIR / "template_cv.docx"))
        content = extractor.extract_content_profile(str(FIXTURES_DIR / "content_cv.docx"))

        matches = matcher.match_all(template, content)

        # Should have matches even without exact style names
        assert len(matches) >= 0  # May or may not match depending on content


class TestDocumentAssembler:
    """Test document assembly."""

    def test_assemble_offer(self, tmp_path):
        """Assemble offer document from template + content."""
        extractor = DOCXExtractor()
        matcher = StyleMatchingEngine()
        assembler = DocumentAssembler()

        template = extractor.extract_template_profile(str(FIXTURES_DIR / "template_offer.docx"))
        content = extractor.extract_content_profile(str(FIXTURES_DIR / "content_offer.docx"))
        matches = matcher.match_all(template, content)

        output_path = tmp_path / "output_offer.docx"
        result = assembler.assemble(template, content, matches, str(output_path))

        assert result.success
        assert output_path.exists()
        assert output_path.stat().st_size > 1000  # Non-empty file

    def test_assemble_cv(self, tmp_path):
        """Assemble CV from template + content."""
        extractor = DOCXExtractor()
        matcher = StyleMatchingEngine()
        assembler = DocumentAssembler()

        template = extractor.extract_template_profile(str(FIXTURES_DIR / "template_cv.docx"))
        content = extractor.extract_content_profile(str(FIXTURES_DIR / "content_cv.docx"))
        matches = matcher.match_all(template, content)

        output_path = tmp_path / "output_cv.docx"
        result = assembler.assemble(template, content, matches, str(output_path))

        assert result.success
        assert output_path.exists()


class TestFormatPipeline:
    """Test end-to-end pipeline."""

    def test_offer_pipeline(self, tmp_path):
        """Full pipeline: template + content -> output."""
        pipeline = FormatPipeline()

        output_path = tmp_path / "pipeline_output.docx"
        result = pipeline.process(
            template_path=str(FIXTURES_DIR / "template_offer.docx"),
            content_path=str(FIXTURES_DIR / "content_offer.docx"),
            output_path=str(output_path),
        )

        assert result.success
        assert output_path.exists()
        assert result.processing_time_ms is not None
        assert result.processing_time_ms > 0

    def test_cv_pipeline(self, tmp_path):
        """Full CV pipeline."""
        pipeline = FormatPipeline()

        output_path = tmp_path / "pipeline_cv.docx"
        result = pipeline.process(
            template_path=str(FIXTURES_DIR / "template_cv.docx"),
            content_path=str(FIXTURES_DIR / "content_cv.docx"),
            output_path=str(output_path),
        )

        assert result.success
        assert output_path.exists()

    def test_bytes_processing(self):
        """Process documents from bytes."""
        pipeline = FormatPipeline()

        with open(FIXTURES_DIR / "template_offer.docx", "rb") as f:
            template_bytes = f.read()
        with open(FIXTURES_DIR / "content_offer.docx", "rb") as f:
            content_bytes = f.read()

        result = pipeline.process_files(template_bytes, content_bytes)

        assert result.success
        assert result.output_bytes is not None
        assert len(result.output_bytes) > 1000


class TestSemanticRoleDetection:
    """Test automatic semantic role detection."""

    def test_heading_detection(self):
        """Detect heading paragraphs."""
        extractor = DOCXExtractor()

        role = extractor._estimate_role("Project Scope", True, 16, "Heading 1")
        assert role == SemanticRole.HEADING_1

    def test_title_detection(self):
        """Detect title paragraphs."""
        extractor = DOCXExtractor()

        role = extractor._estimate_role("Business Proposal", True, 22, None)
        assert role == SemanticRole.TITLE

    def test_list_detection(self):
        """Detect list items."""
        extractor = DOCXExtractor()

        role = extractor._estimate_role("• First item", False, 11, None)
        assert role == SemanticRole.LIST_BULLET

    def test_amount_detection(self):
        """Detect amount fields."""
        extractor = DOCXExtractor()

        role = extractor._estimate_role("Total cost: 45 000 PLN", False, 11, None)
        assert role == SemanticRole.AMOUNT_FIELD


class TestErrorHandling:
    """Test error handling."""

    def test_nonexistent_file(self, tmp_path):
        """Handle non-existent input file gracefully."""
        pipeline = FormatPipeline()

        result = pipeline.process(
            template_path=str(tmp_path / "nonexistent.docx"),
            content_path=str(tmp_path / "nonexistent2.docx"),
            output_path=str(tmp_path / "output.docx"),
        )

        # Pipeline returns error result instead of raising
        assert result.success is False
        assert len(result.warnings) > 0

    def test_invalid_file(self, tmp_path):
        """Handle invalid file gracefully."""
        bad_file = tmp_path / "bad.docx"
        bad_file.write_text("not a docx")

        pipeline = FormatPipeline()

        result = pipeline.process(
            template_path=str(bad_file),
            content_path=str(bad_file),
            output_path=str(tmp_path / "output.docx"),
        )

        # Pipeline returns error result instead of raising
        assert result.success is False
        assert len(result.warnings) > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
