"""
Style Matching Engine - maps content styles to template styles.

Uses a cascade of matchers:
1. Exact style ID match
2. Fuzzy name matching
3. Outline level matching
4. Semantic role matching
5. LLM-based matching (optional, if OpenRouter key configured)
6. Content heuristic matching
"""

import logging
from difflib import SequenceMatcher
from typing import Dict, List, Optional

from .llm_matcher import LLMStyleMatcher
from .types import (
    ContentProfile,
    MatchLog,
    ParagraphContent,
    ParagraphStyle,
    ProcessingLog,
    SemanticRole,
    StyleMatch,
    TemplateProfile,
)

logger = logging.getLogger(__name__)


class StyleMatchingEngine:
    """
    Matches styles from content document to template styles.

    Uses a cascade of matchers:
    1. Exact style ID match
    2. Fuzzy name matching
    3. Semantic role matching
    4. LLM-based matching (if API key available)
    5. Content heuristic matching
    """

    def __init__(self, min_confidence: float = 0.5, use_llm: bool = True):
        self.min_confidence = min_confidence
        self.llm_matcher = LLMStyleMatcher() if use_llm else None
        self.last_log: ProcessingLog = ProcessingLog()

    def match_all(self, template: TemplateProfile, content: ContentProfile) -> List[StyleMatch]:
        """
        Match all content paragraphs to template styles.
        Returns list of unique style mappings.
        Populates self.last_log with detailed matching logs.
        """
        self.last_log = ProcessingLog(
            template_styles_found=len(template.paragraph_styles),
            content_paragraphs=len(content.paragraphs),
            llm_available=self.llm_matcher is not None and self.llm_matcher.is_available(),
        )
        matches = []
        matched_source = set()

        # Pass 1: Exact style ID matching (only for non-generic styles)
        for para in content.paragraphs:
            if not para.style_name:
                continue
            if para.style_name in ("Normal", "BodyText", "BodyText2", "NoSpacing"):
                continue
            if para.style_name in template.paragraph_styles:
                match = StyleMatch(
                    source_style_id=para.style_name,
                    target_style_id=para.style_name,
                    confidence=1.0,
                    reason="Exact style ID match",
                    matcher_type="exact",
                )
                if para.style_name not in matched_source:
                    matches.append(match)
                    matched_source.add(para.style_name)
                    self.last_log.match_logs.append(MatchLog(
                        pass_name="exact",
                        source_style=para.style_name,
                        target_style=para.style_name,
                        confidence=1.0,
                        reason="Exact style ID match",
                        paragraph_preview=para.text[:60],
                    ))

        # Pass 2: Fuzzy name matching for unmatched styles
        for para in content.paragraphs:
            if para.style_name in matched_source:
                continue
            if not para.style_name:
                continue
            if para.style_name in ("Normal", "BodyText", "BodyText2", "NoSpacing"):
                continue

            best_match = self._fuzzy_match_style(para.style_name, template.paragraph_styles)
            if best_match and best_match.confidence >= self.min_confidence:
                matches.append(best_match)
                matched_source.add(para.style_name)
                self.last_log.match_logs.append(MatchLog(
                    pass_name="fuzzy",
                    source_style=best_match.source_style_id,
                    target_style=best_match.target_style_id,
                    confidence=best_match.confidence,
                    reason=best_match.reason,
                    paragraph_preview=para.text[:60],
                ))

        # Pass 3: Semantic role matching for ALL paragraphs (including Normal-styled)
        for para in content.paragraphs:
            if para.style_name and para.style_name in matched_source:
                continue

            best_match = self._semantic_match(para, template)
            if best_match and best_match.confidence >= self.min_confidence:
                source_key = para.style_name or f"__role_{para.estimated_role.value}__"
                match = StyleMatch(
                    source_style_id=source_key,
                    target_style_id=best_match.target_style_id,
                    confidence=best_match.confidence,
                    reason=best_match.reason,
                    matcher_type=best_match.matcher_type,
                )
                if source_key not in matched_source:
                    matches.append(match)
                    matched_source.add(source_key)
                    self.last_log.match_logs.append(MatchLog(
                        pass_name="semantic",
                        source_style=source_key,
                        target_style=best_match.target_style_id,
                        confidence=best_match.confidence,
                        reason=best_match.reason,
                        paragraph_preview=para.text[:60],
                    ))

        # Pass 4: LLM-based matching for anything still unmatched (if available)
        if self.llm_matcher and self.llm_matcher.is_available():
            unmatched_for_llm = [
                p for p in content.paragraphs
                if (p.style_name and p.style_name not in matched_source)
                or (not p.style_name and p.estimated_role == SemanticRole.UNKNOWN)
            ]
            if unmatched_for_llm:
                self.last_log.llm_used = True
                llm_matches = self.llm_matcher.match(template, unmatched_for_llm)
                for llm_match in llm_matches:
                    if llm_match.source_style_id not in matched_source:
                        matches.append(llm_match)
                        matched_source.add(llm_match.source_style_id)
                        self.last_log.match_logs.append(MatchLog(
                            pass_name="llm",
                            source_style=llm_match.source_style_id,
                            target_style=llm_match.target_style_id,
                            confidence=llm_match.confidence,
                            reason=llm_match.reason,
                            paragraph_preview=unmatched_for_llm[0].text[:60] if unmatched_for_llm else "",
                        ))

        # Pass 5: Content heuristic for anything still unmatched
        for para in content.paragraphs:
            if para.style_name and para.style_name in matched_source:
                continue

            best_match = self._content_heuristic_match(para, template)
            if best_match and best_match.confidence >= self.min_confidence:
                source_key = para.style_name or f"__heuristic_{para.text[:20]}__"
                if source_key not in matched_source:
                    matches.append(StyleMatch(
                        source_style_id=source_key,
                        target_style_id=best_match.target_style_id,
                        confidence=best_match.confidence,
                        reason=best_match.reason,
                        matcher_type=best_match.matcher_type,
                    ))
                    matched_source.add(source_key)
                    self.last_log.match_logs.append(MatchLog(
                        pass_name="heuristic",
                        source_style=source_key,
                        target_style=best_match.target_style_id,
                        confidence=best_match.confidence,
                        reason=best_match.reason,
                        paragraph_preview=para.text[:60],
                    ))

        # Build final style mapping with deduplication
        final_map = {}
        for match in matches:
            if match.source_style_id not in final_map:
                final_map[match.source_style_id] = match
            elif match.confidence > final_map[match.source_style_id].confidence:
                final_map[match.source_style_id] = match

        # Collect unmatched styles
        for para in content.paragraphs:
            sid = para.style_name or f"__unstyled_{para.text[:20]}__"
            if sid not in matched_source and sid not in self.last_log.unmatched_after_all_passes:
                self.last_log.unmatched_after_all_passes.append(sid)

        result = list(final_map.values())
        logger.info(f"Matched {len(result)} unique style mappings")
        return result

    def _fuzzy_match_style(
        self, source_style_name: str,
        template_styles: Dict[str, ParagraphStyle]
    ) -> Optional[StyleMatch]:
        """Find best matching template style by fuzzy string comparison."""
        best_ratio = 0.0
        best_target = None

        source_clean = source_style_name.lower().replace(" ", "").replace("-", "").replace("_", "")

        for style_id, tstyle in template_styles.items():
            target_clean = style_id.lower().replace(" ", "").replace("-", "").replace("_", "")
            ratio_id = SequenceMatcher(None, source_clean, target_clean).ratio()

            target_name_clean = tstyle.name.lower().replace(" ", "").replace("-", "").replace("_", "")
            ratio_name = SequenceMatcher(None, source_clean, target_name_clean).ratio()

            ratio = max(ratio_id, ratio_name)

            if ratio > best_ratio:
                best_ratio = ratio
                best_target = style_id

        if best_target and best_ratio >= 0.6:
            return StyleMatch(
                source_style_id=source_style_name,
                target_style_id=best_target,
                confidence=best_ratio,
                reason=f"Fuzzy name match ({best_ratio:.2f})",
                matcher_type="fuzzy",
            )
        return None

    def _semantic_match(self, para: ParagraphContent, template: TemplateProfile) -> Optional[StyleMatch]:
        """Match based on semantic role and content heuristics."""
        content_role = para.estimated_role

        if not content_role or content_role == SemanticRole.UNKNOWN:
            content_role = self._infer_role_from_content(para)

        best_target = None
        best_confidence = 0.0

        for style_id, tstyle in template.paragraph_styles.items():
            trole = tstyle.semantic_role
            if not trole:
                continue

            if trole == content_role:
                confidence = max(0.7, tstyle.role_confidence)
                if confidence > best_confidence:
                    best_confidence = confidence
                    best_target = style_id
                continue

            if content_role.value.startswith("heading") and trole.value.startswith("heading"):
                content_level = self._extract_level(content_role.value)
                target_level = self._extract_level(trole.value)
                if content_level == target_level:
                    confidence = 0.85
                    if confidence > best_confidence:
                        best_confidence = confidence
                        best_target = style_id

        if best_target:
            source_id = para.style_name or f"__estimated_{content_role.value}__"
            return StyleMatch(
                source_style_id=source_id,
                target_style_id=best_target,
                confidence=best_confidence,
                reason=f"Semantic role match: {content_role.value} -> {template.paragraph_styles[best_target].name}",
                matcher_type="semantic",
            )

        return self._content_heuristic_match(para, template)

    def _infer_role_from_content(self, para: ParagraphContent) -> Optional[SemanticRole]:
        """Infer semantic role from content when no style hint is available."""
        text = para.text.strip()

        if not text:
            return SemanticRole.BODY_TEXT

        if len(text) < 100 and para.has_bold:
            if para.max_font_size and para.max_font_size >= 20:
                return SemanticRole.TITLE
            if para.max_font_size and para.max_font_size >= 14:
                return SemanticRole.HEADING_1
            return SemanticRole.HEADING_2

        if para.is_list_item or text.startswith(("•", "●", "○", "▪", "-")):
            return SemanticRole.LIST_BULLET

        if text[:3].strip().endswith(".") and text[:2].strip()[0].isdigit():
            return SemanticRole.LIST_NUMBER

        return SemanticRole.BODY_TEXT

    def _content_heuristic_match(self, para: ParagraphContent,
                                  template: TemplateProfile) -> Optional[StyleMatch]:
        """
        Heuristic matching based on content properties vs template style properties.
        """
        text = para.text.strip()
        estimated_size = para.max_font_size or 11

        best_target = None
        best_score = 0.0

        for style_id, tstyle in template.paragraph_styles.items():
            score = 0.0
            checks = 0

            if tstyle.font.size_pt:
                size_diff = abs(tstyle.font.size_pt - estimated_size)
                checks += 1
                if size_diff < 1:
                    score += 1.0
                elif size_diff < 3:
                    score += 0.7
                elif size_diff < 5:
                    score += 0.4

            if para.has_bold is not None and tstyle.font.bold is not None:
                checks += 1
                if para.has_bold == tstyle.font.bold:
                    score += 1.0

            if checks > 0:
                avg_score = score / checks
                if avg_score > best_score:
                    best_score = avg_score
                    best_target = style_id

        if best_target and best_score >= self.min_confidence:
            source_id = para.style_name or f"__heuristic_{text[:20]}__"
            return StyleMatch(
                source_style_id=source_id,
                target_style_id=best_target,
                confidence=best_score,
                reason=f"Content heuristic match (score: {best_score:.2f})",
                matcher_type="heuristic",
            )

        return None

    def _extract_level(self, role_value: str) -> Optional[int]:
        """Extract numeric level from role string like 'heading_1'."""
        parts = role_value.split("_")
        if len(parts) == 2 and parts[1].isdigit():
            return int(parts[1]) - 1
        return None


__all__ = ["StyleMatchingEngine"]
