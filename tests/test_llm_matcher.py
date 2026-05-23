"""Tests for LLM style matcher."""

from unittest.mock import Mock, patch

from docx_formatter.core.llm_matcher import LLMStyleMatcher, _build_prompt
from docx_formatter.core.types import (
    FontStyle,
    ParagraphContent,
    ParagraphStyle,
    SemanticRole,
    TemplateProfile,
)


class TestLLMStyleMatcher:
    """Test LLM style matcher functionality."""

    def test_is_available_with_key(self):
        """Matcher is available when API key is provided."""
        matcher = LLMStyleMatcher(api_key="test-key")
        assert matcher.is_available() is True

    def test_is_available_without_key(self):
        """Matcher is not available without API key."""
        matcher = LLMStyleMatcher(api_key=None)
        assert matcher.is_available() is False

    def test_match_returns_empty_when_unavailable(self):
        """Returns empty list when API key is missing."""
        matcher = LLMStyleMatcher(api_key=None)
        template = TemplateProfile()
        paras = [ParagraphContent(text="Hello", style_name="Normal")]
        result = matcher.match(template, paras)
        assert result == []

    def test_match_returns_empty_for_empty_paras(self):
        """Returns empty list when no paragraphs provided."""
        matcher = LLMStyleMatcher(api_key="test-key")
        template = TemplateProfile()
        result = matcher.match(template, [])
        assert result == []

    @patch("docx_formatter.core.llm_matcher.httpx.Client")
    def test_match_successful_api_call(self, mock_client_class):
        """Successful API call returns parsed matches."""
        # Setup mock response
        mock_response = Mock()
        mock_response.json.return_value = {
            "choices": [
                {
                    "message": {
                        "content": '[{"content_style_name": "CustomStyle", "template_style_id": "Heading1", "confidence": 0.92, "reason": "Bold heading text"}]'
                    }
                }
            ]
        }
        mock_response.raise_for_status = Mock()

        mock_client = Mock()
        mock_client.post.return_value = mock_response
        mock_client.__enter__ = Mock(return_value=mock_client)
        mock_client.__exit__ = Mock(return_value=False)
        mock_client_class.return_value = mock_client

        matcher = LLMStyleMatcher(api_key="test-key")

        template = TemplateProfile()
        template.paragraph_styles["Heading1"] = ParagraphStyle(
            style_id="Heading1",
            name="Heading 1",
            font=FontStyle(size_pt=16, bold=True),
            semantic_role=SemanticRole.HEADING_1,
        )

        paras = [ParagraphContent(text="Introduction", style_name="CustomStyle", has_bold=True)]
        result = matcher.match(template, paras)

        assert len(result) == 1
        assert result[0].source_style_id == "CustomStyle"
        assert result[0].target_style_id == "Heading1"
        assert result[0].confidence == 0.92
        assert result[0].matcher_type == "llm"

    @patch("docx_formatter.core.llm_matcher.httpx.Client")
    def test_match_api_error_graceful(self, mock_client_class):
        """API errors are handled gracefully."""
        mock_client = Mock()
        mock_client.post.side_effect = Exception("Connection timeout")
        mock_client.__enter__ = Mock(return_value=mock_client)
        mock_client.__exit__ = Mock(return_value=False)
        mock_client_class.return_value = mock_client

        matcher = LLMStyleMatcher(api_key="test-key")
        template = TemplateProfile()
        paras = [ParagraphContent(text="Test", style_name="Normal")]

        result = matcher.match(template, paras)
        assert result == []

    @patch("docx_formatter.core.llm_matcher.httpx.Client")
    def test_match_fallback_to_normal(self, mock_client_class):
        """Invalid style IDs fall back to Normal."""
        mock_response = Mock()
        mock_response.json.return_value = {
            "choices": [
                {
                    "message": {
                        "content": '[{"content_style_name": "Unknown", "template_style_id": "NonExistent", "confidence": 0.8, "reason": "test"}]'
                    }
                }
            ]
        }
        mock_response.raise_for_status = Mock()

        mock_client = Mock()
        mock_client.post.return_value = mock_response
        mock_client.__enter__ = Mock(return_value=mock_client)
        mock_client.__exit__ = Mock(return_value=False)
        mock_client_class.return_value = mock_client

        matcher = LLMStyleMatcher(api_key="test-key")
        template = TemplateProfile()
        template.paragraph_styles["Normal"] = ParagraphStyle(
            style_id="Normal", name="Normal", font=FontStyle()
        )

        paras = [ParagraphContent(text="Test", style_name="Unknown")]
        result = matcher.match(template, paras)

        assert len(result) == 1
        assert result[0].target_style_id == "Normal"
        assert result[0].confidence < 0.8  # Reduced due to fallback


class TestBuildPrompt:
    """Test prompt construction."""

    def test_prompt_includes_styles(self):
        """Prompt includes template styles."""
        template = TemplateProfile()
        template.paragraph_styles["Heading1"] = ParagraphStyle(
            style_id="Heading1",
            name="Heading 1",
            font=FontStyle(size_pt=16, bold=True),
            semantic_role=SemanticRole.HEADING_1,
        )
        paras = [ParagraphContent(text="Hello World", style_name="Custom")]

        prompt = _build_prompt(template, paras)
        assert "Heading1" in prompt
        assert "Hello World" in prompt
        assert "TEMPLATE STYLES" in prompt
        assert "CONTENT PARAGRAPHS" in prompt

    def test_prompt_truncates_long_text(self):
        """Long text is truncated in prompt."""
        template = TemplateProfile()
        paras = [ParagraphContent(text="A" * 200, style_name="Normal")]

        prompt = _build_prompt(template, paras)
        assert len(prompt) < 1000  # Should be truncated
