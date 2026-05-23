"""
LLM Style Matcher - intelligent style matching via OpenRouter API.

Uses a language model to map content styles to template styles when
heuristic methods produce low-confidence results.
"""

import json
import logging
import os
from typing import List, Optional

import httpx

from .types import ParagraphContent, StyleMatch, TemplateProfile

logger = logging.getLogger(__name__)

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "moonshotai/moonshot-v1-8k"

SYSTEM_PROMPT = """You are a document formatting expert. Your task is to match content paragraph styles to template styles.

Given:
- A list of template styles with their properties (name, semantic role, font size, bold/italic)
- A list of content paragraphs with their text and properties

Return a JSON array of matches. Each match must have:
- content_style_name: the style name from the content document (or "__unknown__" if unstyled)
- template_style_id: the best matching template style ID
- confidence: float 0.0-1.0
- reason: brief explanation

Rules:
1. Match by semantic role first (Heading 1 → Heading 1, body → body)
2. Consider font size, bold, italic for fine-tuning
3. If no good match exists, use the template's default "Normal" style with lower confidence
4. Be conservative — assign confidence < 0.6 only when truly uncertain
"""


def _build_prompt(template: TemplateProfile, unmatched_paras: List[ParagraphContent]) -> str:
    """Build the user prompt for the LLM."""
    # Template styles summary
    styles_summary = []
    for sid, style in template.paragraph_styles.items():
        role = style.semantic_role.value if style.semantic_role else "none"
        props = []
        if style.font.size_pt:
            props.append(f"size={style.font.size_pt}pt")
        if style.font.bold:
            props.append("bold")
        if style.font.italic:
            props.append("italic")
        props_str = ", ".join(props) if props else "no special props"
        styles_summary.append(
            f"  - {sid} (name='{style.name}', role={role}, {props_str})"
        )

    # Content paragraphs summary (limit to avoid token overflow)
    paras_summary = []
    for i, para in enumerate(unmatched_paras[:30]):
        role = para.estimated_role.value if para.estimated_role else "none"
        text_preview = para.text[:60].replace('"', "'")
        props = []
        if para.max_font_size:
            props.append(f"size={para.max_font_size}pt")
        if para.has_bold:
            props.append("bold")
        if para.has_italic:
            props.append("italic")
        props_str = ", ".join(props) if props else "no special props"
        style_name = para.style_name or "(no style)"
        paras_summary.append(
            f"  [{i}] style='{style_name}', role={role}, text=\"{text_preview}\", {props_str}"
        )

    return f"""TEMPLATE STYLES:
{chr(10).join(styles_summary)}

CONTENT PARAGRAPHS TO MATCH:
{chr(10).join(paras_summary)}

Return JSON array of matches for ALL content paragraphs above."""


class LLMStyleMatcher:
    """Matches styles using OpenRouter LLM API."""

    def __init__(self, api_key: Optional[str] = None, model: str = MODEL):
        self.api_key = api_key or os.environ.get("OPENROUTER_API_KEY")
        self.model = model
        self.client = httpx.Client(timeout=30.0)

    def is_available(self) -> bool:
        """Check if API key is configured."""
        return bool(self.api_key)

    def match(
        self,
        template: TemplateProfile,
        unmatched_paras: List[ParagraphContent],
    ) -> List[StyleMatch]:
        """
        Get LLM-based style matches for unmatched paragraphs.
        Returns empty list if API is not available or call fails.
        """
        if not self.is_available():
            logger.warning("OpenRouter API key not configured, skipping LLM matching")
            return []

        if not unmatched_paras:
            return []

        prompt = _build_prompt(template, unmatched_paras)

        try:
            response = self.client.post(
                OPENROUTER_URL,
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                    "HTTP-Referer": "https://docx-formatter.vercel.app",
                    "X-Title": "DOCX Formatter",
                },
                json={
                    "model": self.model,
                    "messages": [
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": prompt},
                    ],
                    "temperature": 0.1,
                    "max_tokens": 2000,
                    "response_format": {"type": "json_object"},
                },
            )
            response.raise_for_status()
            data = response.json()

            content = data["choices"][0]["message"]["content"]
            result = json.loads(content)
            if isinstance(result, list):
                matches = result
            elif isinstance(result, dict):
                matches = result.get("matches", [result])
            else:
                matches = []
            if not isinstance(matches, list):
                matches = [matches]

            return self._parse_matches(matches, template, unmatched_paras)

        except Exception as e:
            logger.error(f"LLM matching failed: {e}")
            return []

    def _parse_matches(
        self,
        raw_matches: List[dict],
        template: TemplateProfile,
        unmatched_paras: List[ParagraphContent],
    ) -> List[StyleMatch]:
        """Parse LLM response into StyleMatch objects."""
        matches = []
        seen = set()

        for raw in raw_matches:
            content_style = raw.get("content_style_name", "__unknown__")
            template_id = raw.get("template_style_id", "Normal")
            confidence = float(raw.get("confidence", 0.5))
            reason = raw.get("reason", "LLM match")

            # Validate template style exists
            if template_id not in template.paragraph_styles:
                template_id = next(iter(template.paragraph_styles), "Normal")
                confidence *= 0.8
                reason += " (fallback to available style)"

            key = content_style
            if key not in seen:
                matches.append(StyleMatch(
                    source_style_id=content_style,
                    target_style_id=template_id,
                    confidence=min(confidence, 0.95),
                    reason=reason,
                    matcher_type="llm",
                ))
                seen.add(key)

        logger.info(f"LLM matcher produced {len(matches)} matches")
        return matches
