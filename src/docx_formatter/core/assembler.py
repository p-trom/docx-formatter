"""
Document Assembler - builds final DOCX output from template + content + matches.

Approach: Use the template DOCX as the base document and replace paragraph text
in each section based on semantic role matching. This preserves:
- Section structure (headers, footers, page setup per section)
- Background images, watermarks
- Paragraph styles and style definitions
- Content layout within each section
"""

import copy
import logging
from typing import Dict, List, Optional, Any
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .types import (
    TemplateProfile, ContentProfile, ParagraphContent,
    ParagraphStyle, StyleMatch, ProcessingResult,
    SemanticRole, FontStyle
)

logger = logging.getLogger(__name__)


class DocumentAssembler:
    """
    Assembles the final DOCX document from template + content + style matches.
    
    Strategy:
    1. Open template DOCX as base (preserves ZIP, media, relationships)
    2. For each template paragraph, identify its semantic role (title, heading1, etc.)
    3. Find matching content paragraphs by same semantic role
    4. Replace template paragraph text with content text (keep pPr/style intact)
    5. Handle images by keeping them as-is (they exist in the template)
    6. Add non-template content content (body text) in sections near its matching heading
    """
    
    def __init__(self, preserve_template_content: bool = False):
        self.preserve_template_content = preserve_template_content
    
    def assemble(
        self,
        template: TemplateProfile,
        content: ContentProfile,
        style_matches: List[StyleMatch],
        output_path: Optional[str] = None,
        template_docx_path: Optional[str] = None,
    ) -> ProcessingResult:
        """
        Build output document from template profile + content + style mappings.

        If template_docx_path is provided, uses the template DOCX as the base
        document. This preserves headers, footers, section properties, background
        images, watermarks, and page setup. Body content is replaced with the
        new formatted content.

        Args:
            template: Profile of the template document (styles, settings)
            content: Profile of the content document (paragraphs, tables)
            style_matches: List of style mappings
            output_path: Where to save output DOCX
            template_docx_path: Optional path to the original template DOCX file

        Returns:
            ProcessingResult with output path or bytes
        """
        result = ProcessingResult()

        try:
            if template_docx_path:
                doc = Document(template_docx_path)
                self._apply_section_aware_replacement(doc, template, content, style_matches)
            else:
                doc = Document()
                self._apply_document_defaults(doc, template)
                self._copy_template_styles(doc, template)
                self._append_all_content(doc, template, content, style_matches)
                self._copy_headers_footers(doc, template)

            if output_path:
                doc.save(output_path)
                result.output_path = output_path
                logger.info(f"Document saved to {output_path}")

            result.matched_styles = style_matches
            result.success = True
            result.warnings = []

        except Exception as e:
            logger.error(f"Assembly failed: {e}", exc_info=True)
            result.warnings.append(str(e))
            result.success = False

        return result

    # ──────────────────────────────────────────────
    # SECTION-AWARE REPLACEMENT (template as base)
    # ──────────────────────────────────────────────

    def _apply_section_aware_replacement(
        self,
        doc: Document,
        template: TemplateProfile,
        content: ContentProfile,
        style_matches: List[StyleMatch],
    ) -> None:
        """Replace template content section by section using semantic role matching."""

        # Build style mapping lookup: source_style_id -> target_style_id
        style_map = {m.source_style_id: m.target_style_id for m in style_matches}

        # 1. Split content paragraphs by semantic role for quick lookup
        content_by_role = self._group_content_by_role(content.paragraphs)

        # 2. Build section boundaries from template document
        sections_boundaries = self._get_section_boundaries(doc)
        logger.info(f"Template has {len(sections_boundaries)} sections")

        # 3. Track which content paragraphs we've used
        used_content_indices = set()

        # 4. For each section, replace template text with matching content
        for sec_idx, (start, end) in enumerate(sections_boundaries):
            for t_para_idx in range(start, end + 1):
                t_para_element = doc.paragraphs[t_para_idx]._element
                # Determine semantic role of this template paragraph
                role = self._get_template_para_role(doc.paragraphs[t_para_idx], template)
                
                # Find best matching content paragraph with same role
                match = self._find_matching_content_para(role, content_by_role, used_content_indices)
                if match:
                    c_para, c_idx = match
                    self._replace_para_text_keep_structure(t_para_element, c_para, template, style_map)
                    used_content_indices.add(c_idx)

        # 5. Handle remaining unmatched content paragraphs
        remaining = [p for i, p in enumerate(content.paragraphs) if i not in used_content_indices]
        if remaining:
            logger.info(f"Adding {len(remaining)} remaining content paragraphs")
            # Insert remaining body text after the last filled section or at the end
            self._add_remaining_content(doc, remaining, template, style_map)

    def _get_section_boundaries(self, doc: Document) -> List[tuple]:
        """Get (start_para_idx, end_para_idx) for each section."""
        sect_breaks = []
        for i, p in enumerate(doc.paragraphs):
            pPr = p._element.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                sect_breaks.append(i)

        boundaries = []
        start = 0
        for break_idx in sect_breaks:
            boundaries.append((start, break_idx))
            start = break_idx + 1
        # Last section goes from last break to end
        boundaries.append((start, len(doc.paragraphs) - 1))
        return boundaries

    def _group_content_by_role(self, paragraphs: List[ParagraphContent]) -> Dict[SemanticRole, List[tuple]]:
        """Group content paragraphs by their semantic role.
        Returns dict of role -> [(para, idx), ...]."""
        result = {}
        for idx, para in enumerate(paragraphs):
            role = self._resolve_content_role(para)
            if role not in result:
                result[role] = []
            result[role].append((para, idx))
        return result

    def _resolve_content_role(self, para: ParagraphContent) -> SemanticRole:
        """Determine semantic role of a content paragraph."""
        if para.estimated_role and para.estimated_role != SemanticRole.UNKNOWN:
            return para.estimated_role
        return self._infer_role_from_content(para)

    def _get_template_para_role(self, para, template: TemplateProfile) -> SemanticRole:
        """Determine semantic role of a template paragraph."""
        # Try by style
        style_name = para.style.name if para.style else None
        if style_name:
            # Check template profile styles
            for sid, tstyle in template.paragraph_styles.items():
                if tstyle.name == style_name or sid == style_name:
                    if tstyle.semantic_role and tstyle.semantic_role != SemanticRole.UNKNOWN:
                        return tstyle.semantic_role
        # Infer from content
        text = para.text.strip()
        if para.style and para.style.name:
            sname = para.style.name.lower()
            if 'title' in sname or 'cover' in sname:
                return SemanticRole.TITLE
            if 'heading 1' == sname or 'heading1' in sname:
                return SemanticRole.HEADING_1
            if 'heading 2' == sname or 'heading2' in sname:
                return SemanticRole.HEADING_2
            if 'heading 3' == sname or 'heading3' in sname:
                return SemanticRole.HEADING_3
            if 'heading 4' == sname or 'heading4' in sname:
                return SemanticRole.HEADING_4
            if 'heading 5' == sname or 'heading5' in sname:
                return SemanticRole.HEADING_4
        if text and len(text) < 100:
            if para.style and hasattr(para.style, 'font'):
                size = para.style.font.size
                if size:
                    size_pt = size.pt if hasattr(size, 'pt') else 11
                    if size_pt >= 20:
                        return SemanticRole.TITLE
                    if size_pt >= 16:
                        return SemanticRole.HEADING_1
                    if size_pt >= 14:
                        return SemanticRole.HEADING_2
        return SemanticRole.BODY_TEXT

    def _find_matching_content_para(
        self,
        role: SemanticRole,
        content_by_role: Dict[SemanticRole, List[tuple]],
        used_indices: set,
    ) -> Optional[tuple]:
        """Find the next unused content paragraph matching the given role."""
        # Try exact role match first
        for candidate_role in [role, self._fallback_role(role)]:
            if candidate_role in content_by_role:
                for para, idx in content_by_role[candidate_role]:
                    if idx not in used_indices:
                        return (para, idx)
        return None

    def _fallback_role(self, role: SemanticRole) -> SemanticRole:
        """Map specific roles to less-specific fallbacks."""
        fallback_map = {
            SemanticRole.HEADING_4: SemanticRole.HEADING_4,
            SemanticRole.HEADING_4: SemanticRole.HEADING_3,
            SemanticRole.HEADING_3: SemanticRole.HEADING_2,
            SemanticRole.HEADING_2: SemanticRole.HEADING_1,
            SemanticRole.TITLE: SemanticRole.HEADING_1,
            SemanticRole.SUBTITLE: SemanticRole.HEADING_2,
        }
        return fallback_map.get(role, SemanticRole.BODY_TEXT)

    def _replace_para_text_keep_structure(
        self,
        t_para_element,
        c_para: ParagraphContent,
        template: TemplateProfile,
        style_map: Dict[str, str],
    ) -> None:
        """Replace text in a template paragraph while keeping pPr (style, sectPr, etc.)."""
        # Clear existing runs
        for run in list(t_para_element.findall(qn('w:r'))):
            t_para_element.remove(run)

        # Add new runs with content text (and any inline formatting from content)
        if c_para.runs:
            # Use content runs with inline formatting
            for run_info in c_para.runs:
                new_run = OxmlElement('w:r')
                rPr = OxmlElement('w:rPr')
                if run_info.get('bold'):
                    r_bold = OxmlElement('w:b')
                    rPr.append(r_bold)
                if run_info.get('italic'):
                    r_italic = OxmlElement('w:i')
                    rPr.append(r_italic)
                new_run.append(rPr)
                t_elem = OxmlElement('w:t')
                t_elem.text = run_info.get('text', '')
                new_run.append(t_elem)
                t_para_element.append(new_run)
        else:
            # Simple text
            new_run = OxmlElement('w:r')
            t_elem = OxmlElement('w:t')
            t_elem.text = c_para.text or ''
            new_run.append(t_elem)
            t_para_element.append(new_run)

    def _add_remaining_content(
        self,
        doc: Document,
        remaining: List[ParagraphContent],
        template: TemplateProfile,
        style_map: Dict[str, str],
    ) -> None:
        """Add remaining unmatched content paragraphs to the document.
        Find the body section and insert there (after the last heading)."""
        body = doc.element.body
        
        # Find last section break para - add after it
        last_sect_para = None
        for child in reversed(list(body)):
            if child.tag == qn('w:p'):
                pPr = child.find(qn('w:pPr'))
                if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                    last_sect_para = child
                    break

        # Insert remaining content after the last section break paragraph
        insert_after = last_sect_para if last_sect_para is not None else None

        for para in remaining:
            new_p = OxmlElement('w:p')
            target_style_id = self._resolve_target_style(para, style_map, template)
            
            # Apply style via pPr
            pPr = OxmlElement('w:pPr')
            if target_style_id:
                pStyle = OxmlElement('w:pStyle')
                pStyle.set(qn('w:val'), target_style_id)
                pPr.append(pStyle)
            new_p.append(pPr)
            
            # Add runs
            if para.runs:
                for run_info in para.runs:
                    new_r = OxmlElement('w:r')
                    if run_info.get('bold') or run_info.get('italic') or run_info.get('underline'):
                        rPr = OxmlElement('w:rPr')
                        if run_info.get('bold'):
                            rPr.append(OxmlElement('w:b'))
                        if run_info.get('italic'):
                            rPr.append(OxmlElement('w:i'))
                        if run_info.get('underline'):
                            rPr.append(OxmlElement('w:u'))
                        new_r.append(rPr)
                    t_elem = OxmlElement('w:t')
                    t_elem.text = run_info.get('text', '')
                    new_r.append(t_elem)
                    new_p.append(new_r)
            else:
                new_r = OxmlElement('w:r')
                t_elem = OxmlElement('w:t')
                t_elem.text = para.text or ''
                new_r.append(t_elem)
                new_p.append(new_r)
            
            if insert_after is not None:
                # Insert after the section-break paragraph using lxml insert
                idx = list(body).index(insert_after)
                body.insert(idx + 1, new_p)
                insert_after = new_p
            else:
                body.append(new_p)

    # ──────────────────────────────────────────────
    # FALLBACK LEGACY (new blank document)
    # ──────────────────────────────────────────────
    
    def _append_all_content(
        self,
        doc: Document,
        template: TemplateProfile,
        content: ContentProfile,
        style_matches: List[StyleMatch],
    ) -> None:
        """Fallback: append all content to a fresh document."""
        style_map = {m.source_style_id: m.target_style_id for m in style_matches}
        for para in content.paragraphs:
            self._add_paragraph(doc, para, style_map, template)
        for table in content.tables:
            self._add_table(doc, table, style_map, template)
    
    def _ensure_template_styles(self, doc: Document, template: TemplateProfile) -> None:
        """Ensure all custom template styles exist in the base document."""
        for style_id, tstyle in template.paragraph_styles.items():
            if tstyle.is_built_in and not tstyle.is_default:
                continue
            try:
                try:
                    existing = doc.styles[style_id]
                    self._apply_paragraph_style(existing, tstyle)
                except KeyError:
                    new_style = doc.styles.add_style(style_id, 1)
                    self._apply_paragraph_style(new_style, tstyle)
            except Exception as e:
                logger.debug(f"Could not ensure style {style_id}: {e}")
    
    def _apply_document_defaults(self, doc: Document, template: TemplateProfile) -> None:
        """Apply page settings and margins from template."""
        defaults = template.document_defaults
        
        if not defaults.margins:
            return
        
        try:
            section = doc.sections[0]
            if defaults.margins:
                if 'top' in defaults.margins:
                    section.top_margin = Inches(defaults.margins['top'])
                if 'bottom' in defaults.margins:
                    section.bottom_margin = Inches(defaults.margins['bottom'])
                if 'left' in defaults.margins:
                    section.left_margin = Inches(defaults.margins['left'])
                if 'right' in defaults.margins:
                    section.right_margin = Inches(defaults.margins['right'])
            
            if defaults.page_width_inches:
                section.page_width = Inches(defaults.page_width_inches)
            if defaults.page_height_inches:
                section.page_height = Inches(defaults.page_height_inches)
            
            if defaults.orientation == 'landscape':
                section.orientation = 1  # WD_ORIENT.LANDSCAPE
        except Exception as e:
            logger.warning(f"Could not apply document defaults: {e}")
    
    def _copy_template_styles(self, doc: Document, template: TemplateProfile) -> None:
        """Copy style definitions from template profile to new document."""
        for style_id, tstyle in template.paragraph_styles.items():
            if tstyle.is_built_in and not tstyle.is_default:
                continue
            try:
                try:
                    existing = doc.styles[style_id]
                    self._apply_paragraph_style(existing, tstyle)
                except KeyError:
                    new_style = doc.styles.add_style(style_id, 1)
                    self._apply_paragraph_style(new_style, tstyle)
            except Exception as e:
                logger.debug(f"Could not copy style {style_id}: {e}")
    
    def _apply_paragraph_style(self, docx_style, tstyle: ParagraphStyle) -> None:
        """Apply ParagraphStyle properties to a python-docx style object."""
        try:
            font = docx_style.font
            if tstyle.font:
                if tstyle.font.name:
                    font.name = tstyle.font.name
                if tstyle.font.size_pt:
                    font.size = Pt(tstyle.font.size_pt)
                if tstyle.font.bold is not None:
                    font.bold = tstyle.font.bold
                if tstyle.font.italic is not None:
                    font.italic = tstyle.font.italic
                if tstyle.font.color and not tstyle.font.color.startswith('theme:'):
                    try:
                        font.color.rgb = RGBColor.from_string(tstyle.font.color)
                    except:
                        pass
            
            pfmt = docx_style.paragraph_format
            if tstyle.alignment and tstyle.alignment.alignment:
                align_map = {
                    '0': WD_ALIGN_PARAGRAPH.LEFT,
                    '1': WD_ALIGN_PARAGRAPH.CENTER,
                    '2': WD_ALIGN_PARAGRAPH.RIGHT,
                    '3': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'LEFT': WD_ALIGN_PARAGRAPH.LEFT,
                    'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
                    'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT,
                    'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'BOTH': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'None': WD_ALIGN_PARAGRAPH.LEFT,
                }
                if tstyle.alignment.alignment in align_map:
                    pfmt.alignment = align_map[tstyle.alignment.alignment]
            
            if tstyle.spacing:
                if tstyle.spacing.before_pt:
                    pfmt.space_before = Pt(tstyle.spacing.before_pt)
                if tstyle.spacing.after_pt:
                    pfmt.space_after = Pt(tstyle.spacing.after_pt)
                if tstyle.spacing.line_spacing:
                    pfmt.line_spacing = tstyle.spacing.line_spacing
            
            if tstyle.indentation:
                if tstyle.indentation.left_inches:
                    pfmt.left_indent = Inches(tstyle.indentation.left_inches)
                if tstyle.indentation.right_inches:
                    pfmt.right_indent = Inches(tstyle.indentation.right_inches)
                if tstyle.indentation.first_line_inches:
                    pfmt.first_line_indent = Inches(tstyle.indentation.first_line_inches)
        except Exception as e:
            logger.debug(f"Error applying style {tstyle.style_id}: {e}")
    
    def _add_paragraph(
        self,
        doc: Document,
        para: ParagraphContent,
        style_map: Dict[str, str],
        template: TemplateProfile,
    ) -> None:
        """Add a paragraph to a NEW document with proper style mapping (legacy mode)."""
        target_style_id = self._resolve_target_style(para, style_map, template)
        new_para = doc.add_paragraph()
        
        if para.runs:
            for run_info in para.runs:
                run = new_para.add_run(run_info.get('text', ''))
                self._apply_run_formatting(run, run_info)
        else:
            new_para.add_run(para.text)
        
        if target_style_id:
            try:
                new_para.style = target_style_id
            except KeyError:
                self._apply_style_by_definition(new_para, target_style_id, template)
        
        self._apply_direct_paragraph_formatting(new_para, para, template, target_style_id)
    
    def _resolve_target_style(
        self,
        para: ParagraphContent,
        style_map: Dict[str, str],
        template: TemplateProfile,
    ) -> Optional[str]:
        """Resolve which template style to apply to this paragraph."""
        if para.style_name and para.style_name not in ('Normal', 'BodyText', 'BodyText2', 'NoSpacing'):
            if para.style_name in style_map:
                return style_map[para.style_name]
            if para.style_name in template.paragraph_styles:
                return para.style_name

        role = para.estimated_role
        if not role or role == SemanticRole.UNKNOWN:
            role = self._infer_role_from_content(para)

        if role and role != SemanticRole.UNKNOWN:
            for style_id, tstyle in template.paragraph_styles.items():
                if tstyle.semantic_role == role:
                    return style_id
            if role.value.startswith('heading') or role.value.startswith('title'):
                for style_id, tstyle in template.paragraph_styles.items():
                    trole = tstyle.semantic_role
                    if trole and (trole.value.startswith('heading') or trole.value.startswith('title')):
                        return style_id

        return self._heuristic_style_match(para, template)

    def _infer_role_from_content(self, para: ParagraphContent) -> SemanticRole:
        """Infer semantic role from content when no style hint is available."""
        text = para.text.strip()
        if not text:
            return SemanticRole.BODY_TEXT
        if len(text) < 100:
            if para.max_font_size and para.max_font_size >= 20:
                return SemanticRole.TITLE
            if para.max_font_size and para.max_font_size >= 14:
                return SemanticRole.HEADING_1
            if para.has_bold:
                return SemanticRole.HEADING_2
        if para.is_list_item or text.startswith(('•', '●', '○', '▪', '-')):
            return SemanticRole.LIST_BULLET
        if text[:3].strip().endswith('.') and text[:2].strip()[0].isdigit():
            return SemanticRole.LIST_NUMBER
        return SemanticRole.BODY_TEXT

    def _heuristic_style_match(self, para: ParagraphContent, template: TemplateProfile) -> Optional[str]:
        """Match paragraph to template style by font size and properties."""
        estimated_size = para.max_font_size or 11
        best_target = None
        best_score = 0.0
        for style_id, tstyle in template.paragraph_styles.items():
            score = 0.0
            checks = 0
            if tstyle.font and tstyle.font.size_pt:
                size_diff = abs(tstyle.font.size_pt - estimated_size)
                checks += 1
                if size_diff < 1:
                    score += 1.0
                elif size_diff < 3:
                    score += 0.7
                elif size_diff < 5:
                    score += 0.4
            if para.has_bold is not None and tstyle.font and tstyle.font.bold is not None:
                checks += 1
                if para.has_bold == tstyle.font.bold:
                    score += 1.0
            if checks > 0 and score / checks > best_score:
                best_score = score / checks
                best_target = style_id
        return best_target if best_score >= 0.3 else None

    def _apply_style_by_definition(self, new_para, style_id: str, template: TemplateProfile) -> None:
        """Apply style properties directly to paragraph when style ID doesn't exist in doc."""
        if style_id not in template.paragraph_styles:
            return
        tstyle = template.paragraph_styles[style_id]
        try:
            if tstyle.font:
                for run in new_para.runs:
                    if tstyle.font.name:
                        run.font.name = tstyle.font.name
                    if tstyle.font.size_pt:
                        run.font.size = Pt(tstyle.font.size_pt)
                    if tstyle.font.bold is not None:
                        run.font.bold = tstyle.font.bold
                    if tstyle.font.italic is not None:
                        run.font.italic = tstyle.font.italic
            pfmt = new_para.paragraph_format
            if tstyle.alignment and tstyle.alignment.alignment:
                align_map = {
                    '0': WD_ALIGN_PARAGRAPH.LEFT,
                    '1': WD_ALIGN_PARAGRAPH.CENTER,
                    '2': WD_ALIGN_PARAGRAPH.RIGHT,
                    '3': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'LEFT': WD_ALIGN_PARAGRAPH.LEFT,
                    'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
                    'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT,
                    'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'BOTH': WD_ALIGN_PARAGRAPH.JUSTIFY,
                    'None': WD_ALIGN_PARAGRAPH.LEFT,
                }
                if tstyle.alignment.alignment in align_map:
                    pfmt.alignment = align_map[tstyle.alignment.alignment]
            if tstyle.spacing:
                if tstyle.spacing.before_pt:
                    pfmt.space_before = Pt(tstyle.spacing.before_pt)
                if tstyle.spacing.after_pt:
                    pfmt.space_after = Pt(tstyle.spacing.after_pt)
                if tstyle.spacing.line_spacing:
                    pfmt.line_spacing = tstyle.spacing.line_spacing
            if tstyle.indentation:
                if tstyle.indentation.left_inches:
                    pfmt.left_indent = Inches(tstyle.indentation.left_inches)
                if tstyle.indentation.right_inches:
                    pfmt.right_indent = Inches(tstyle.indentation.right_inches)
                if tstyle.indentation.first_line_inches:
                    pfmt.first_line_indent = Inches(tstyle.indentation.first_line_inches)
        except Exception as e:
            logger.debug(f"Could not apply style definition {style_id}: {e}")

    def _apply_run_formatting(self, run, run_info: Dict[str, Any]) -> None:
        """Apply inline formatting to a run."""
        try:
            if run_info.get('bold'):
                run.bold = True
            if run_info.get('italic'):
                run.italic = True
            if run_info.get('underline'):
                run.underline = True
        except Exception:
            pass
    
    def _apply_direct_paragraph_formatting(
        self,
        new_para,
        para: ParagraphContent,
        template: TemplateProfile,
        target_style_id: Optional[str],
    ) -> None:
        """Apply direct formatting not covered by the style definition."""
        try:
            pfmt = new_para.paragraph_format
            if para.is_list_item:
                indent = para.indentation_left_inches or 0.5 * para.list_level
                if indent > 0 and not pfmt.left_indent:
                    pfmt.left_indent = Inches(indent)
            if para.alignment and not target_style_id:
                align_map = {
                    'LEFT': WD_ALIGN_PARAGRAPH.LEFT,
                    'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
                    'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT,
                    'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY,
                }
                if para.alignment in align_map:
                    pfmt.alignment = align_map[para.alignment]
        except Exception:
            pass
    
    def _add_table(
        self,
        doc: Document,
        table_info: Any,
        style_map: Dict[str, str],
        template: TemplateProfile,
    ) -> None:
        """Add a table to the document."""
        row_count = getattr(table_info, 'row_count', 0) or 2
        col_count = getattr(table_info, 'column_count', 0) or 2
        
        try:
            table = doc.add_table(rows=row_count, cols=col_count)
            table.style = 'Table Grid'
            
            rows = getattr(table_info, 'rows', None)
            if rows:
                for i, row_data in enumerate(rows):
                    if i >= len(table.rows):
                        break
                    row = table.rows[i]
                    if isinstance(row_data, list):
                        for j, cell_text in enumerate(row_data):
                            if j >= len(row.cells):
                                break
                            row.cells[j].text = str(cell_text)
        except Exception as e:
            logger.warning(f"Could not add table: {e}")
    
    def _copy_headers_footers(self, doc: Document, template: TemplateProfile) -> None:
        """Copy headers and footers to a NEW document (legacy mode)."""
        try:
            if template.headers:
                section = doc.sections[0]
                if section.header is not None and template.headers.get('default'):
                    header_para = section.header.paragraphs[0]
                    header_para.text = template.headers.get('default', '')
            
            if template.footers:
                section = doc.sections[0]
                if section.footer is not None and template.footers.get('default'):
                    footer_para = section.footer.paragraphs[0]
                    footer_para.text = template.footers.get('default', '')
        except Exception as e:
            logger.warning(f"Could not copy headers/footers: {e}")


# ──────────────────────────────────────────────
# Remaining: _clear_body_content (removed)
# ──────────────────────────────────────────────
