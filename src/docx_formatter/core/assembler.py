"""
Document Assembler - builds final DOCX output from template + content + matches.
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
    Assembles the final DOCX document.
    
    Creates a new document based on template styles and content.
    Handles style definitions, paragraph mapping, headers/footers,
    page settings, and placeholder replacement.
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
            # Build style mapping lookup: source_style_id -> target_style_id
            style_map = {m.source_style_id: m.target_style_id for m in style_matches}

            if template_docx_path:
                # Use template as base — preserves headers, footers, sections, backgrounds
                doc = Document(template_docx_path)
                self._clear_body_content(doc)
                # Ensure template styles are available in the doc
                self._ensure_template_styles(doc, template)
            else:
                # Fallback: create new blank document (legacy behaviour)
                doc = Document()
                self._apply_document_defaults(doc, template)
                self._copy_template_styles(doc, template)

            # Add paragraphs from content with mapped styles
            for para in content.paragraphs:
                self._add_paragraph(doc, para, style_map, template)

            # Add tables from content
            for table in content.tables:
                self._add_table(doc, table, style_map, template)

            # If using legacy mode (no template path), copy headers/footers manually
            if not template_docx_path:
                self._copy_headers_footers(doc, template)

            # Save document
            if output_path:
                doc.save(output_path)
                result.output_path = output_path
                logger.info(f"Document saved to {output_path}")

            result.matched_styles = style_matches
            result.success = True
            result.warnings = []

        except Exception as e:
            logger.error(f"Assembly failed: {e}")
            result.warnings.append(str(e))
            result.success = False

        return result
    
    def _clear_body_content(self, doc: Document) -> None:
        """Remove all paragraphs and tables from the document body while preserving
        section properties (headers, footers, page setup, background images)."""
        body = doc.element.body
        # Remove all <w:p> and <w:tbl> elements; keep <w:sectPr> (section properties)
        for child in list(body):
            if child.tag in (qn('w:p'), qn('w:tbl')):
                body.remove(child)
        logger.info("Cleared body content from template; headers/footers/sections preserved")

    def _ensure_template_styles(self, doc: Document, template: TemplateProfile) -> None:
        """Ensure all custom template styles exist in the base document.
        When opening an existing DOCX, built-in styles are already present,
        but custom styles may need updating."""
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
        """
        Copy style definitions from template profile to new document.
        python-docx copies built-in styles automatically, but custom styles
        need to be recreated.
        """
        for style_id, tstyle in template.paragraph_styles.items():
            if tstyle.is_built_in and not tstyle.is_default:
                # Built-in styles don't need recreation
                continue
            
            try:
                # Check if style already exists
                existing = None
                try:
                    existing = doc.styles[style_id]
                except KeyError:
                    pass
                
                if existing is None:
                    # Create custom style
                    new_style = doc.styles.add_style(style_id, 1)  # 1 = PARAGRAPH
                    self._apply_paragraph_style(new_style, tstyle)
                else:
                    # Update existing style
                    self._apply_paragraph_style(existing, tstyle)
            
            except Exception as e:
                logger.debug(f"Could not copy style {style_id}: {e}")
    
    def _apply_paragraph_style(self, docx_style, tstyle: ParagraphStyle) -> None:
        """Apply ParagraphStyle properties to a python-docx style object."""
        try:
            # Font
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
            
            # Paragraph format
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
        """Add a paragraph to the document with proper style mapping."""
        # Determine target style using full resolution cascade
        target_style_id = self._resolve_target_style(para, style_map, template)

        # Create paragraph - try to use style, fall back to Normal + direct formatting
        new_para = doc.add_paragraph()

        # Add runs with inline formatting first
        if para.runs:
            for run_info in para.runs:
                run = new_para.add_run(run_info.get('text', ''))
                self._apply_run_formatting(run, run_info)
        else:
            new_para.add_run(para.text)

        # Now apply template style definition directly (runs exist now)
        if target_style_id:
            try:
                new_para.style = target_style_id
            except KeyError:
                # Style doesn't exist in new doc — apply definition directly to runs + paragraph
                self._apply_style_by_definition(new_para, target_style_id, template)

        # Apply direct paragraph formatting for things not covered by style
        self._apply_direct_paragraph_formatting(new_para, para, template, target_style_id)
    
    def _apply_style_by_definition(self, new_para, style_id: str, template: TemplateProfile) -> None:
        """Apply style properties directly to paragraph when style ID doesn't exist in doc."""
        if style_id not in template.paragraph_styles:
            return
        tstyle = template.paragraph_styles[style_id]
        
        try:
            # Font
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
            
            # Paragraph format
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
    
    def _resolve_target_style(
        self,
        para: ParagraphContent,
        style_map: Dict[str, str],
        template: TemplateProfile,
    ) -> Optional[str]:
        """Resolve which template style to apply to this paragraph."""
        # 1. If has explicit non-generic style and it's mapped in style_map
        if para.style_name and para.style_name not in ('Normal', 'BodyText', 'BodyText2', 'NoSpacing'):
            if para.style_name in style_map:
                return style_map[para.style_name]
            if para.style_name in template.paragraph_styles:
                return para.style_name

        # 2. Semantic role matching on the paragraph's estimated role
        role = para.estimated_role
        if not role or role == SemanticRole.UNKNOWN:
            role = self._infer_role_from_content(para)

        if role and role != SemanticRole.UNKNOWN:
            # Direct semantic role match against template styles
            for style_id, tstyle in template.paragraph_styles.items():
                if tstyle.semantic_role == role:
                    return style_id

            # Heading level fallback: any heading style if role is heading
            if role.value.startswith('heading') or role.value.startswith('title'):
                for style_id, tstyle in template.paragraph_styles.items():
                    trole = tstyle.semantic_role
                    if trole and (trole.value.startswith('heading') or trole.value.startswith('title')):
                        return style_id

        # 3. Fallback: heuristic matching
        return self._heuristic_style_match(para, template)

    def _infer_role_from_content(self, para: ParagraphContent) -> SemanticRole:
        """Infer semantic role from content when no style hint is available."""
        text = para.text.strip()
        if not text:
            return SemanticRole.BODY_TEXT
        # Title: short, large font
        if len(text) < 100:
            if para.max_font_size and para.max_font_size >= 20:
                return SemanticRole.TITLE
            if para.max_font_size and para.max_font_size >= 14:
                return SemanticRole.HEADING_1
            if para.has_bold:
                return SemanticRole.HEADING_2
        # List item
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

    def _get_semantic_style(self, role: SemanticRole, template: TemplateProfile) -> Optional[str]:
        """Get template style for semantic role."""
        if role in template.semantic_styles:
            return template.semantic_styles[role]
        for style_id, tstyle in template.paragraph_styles.items():
            if tstyle.semantic_role == role:
                return style_id
        return None
    
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
            
            # List item formatting
            if para.is_list_item:
                # Apply list indentation
                indent = para.indentation_left_inches or 0.5 * para.list_level
                if indent > 0 and not pfmt.left_indent:
                    pfmt.left_indent = Inches(indent)
            
            # Alignment from content if not in style
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
        table_content,
        style_map: Dict[str, str],
        template: TemplateProfile,
    ) -> None:
        """Add a table to the document."""
        if not table_content.rows:
            return
        
        num_rows = len(table_content.rows)
        num_cols = max(len(row) for row in table_content.rows) if table_content.rows else 0
        
        if num_rows == 0 or num_cols == 0:
            return
        
        try:
            table = doc.add_table(rows=num_rows, cols=num_cols)
            
            # Apply table style from template if available
            if table_content.style_name:
                mapped_style = style_map.get(table_content.style_name)
                if mapped_style:
                    try:
                        table.style = mapped_style
                    except:
                        pass
            
            # Fill cell content
            for i, row_data in enumerate(table_content.rows):
                if i >= num_rows:
                    break
                row = table.rows[i]
                for j, cell_text in enumerate(row_data):
                    if j >= num_cols:
                        break
                    cell = row.cells[j]
                    cell.text = str(cell_text) if cell_text else ''
        
        except Exception as e:
            logger.warning(f"Could not add table: {e}")
    
    def _copy_headers_footers(self, doc: Document, template: TemplateProfile) -> None:
        """Copy headers and footers from template profile."""
        if not template.headers_footers:
            return
        
        try:
            section = doc.sections[0]
            for hf in template.headers_footers:
                if hf.type == 'header':
                    if hf.section_type == 'default':
                        section.header.is_linked_to_previous = False
                        # Note: Full header/footer content copying is complex
                        # and requires low-level XML manipulation
                        if hf.content:
                            section.header.paragraphs[0].text = hf.content
                elif hf.type == 'footer':
                    if hf.section_type == 'default':
                        section.footer.is_linked_to_previous = False
                        if hf.content:
                            section.footer.paragraphs[0].text = hf.content
        except Exception as e:
            logger.warning(f"Could not copy headers/footers: {e}")


__all__ = ['DocumentAssembler']
