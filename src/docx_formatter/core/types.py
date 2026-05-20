"""
Core data models for DOCX processing.
"""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Optional, List, Dict, Any, Tuple
from decimal import Decimal


class TemplateType(Enum):
    """Typ template'u na podstawie zawartości."""
    STYLE_DEFINITIONS = auto()    # Tylko definicje stylów, bez treści
    EXAMPLE_DOCUMENT = auto()     # Dokument przykładowy z "lorem ipsum"
    TEXT_INSTRUCTIONS = auto()    # Dokument z tekstowymi instrukcjami


class SemanticRole(Enum):
    """Semantyczna rola elementu dokumentu."""
    TITLE = "title"
    SUBTITLE = "subtitle"
    HEADING_1 = "heading_1"
    HEADING_2 = "heading_2"
    HEADING_3 = "heading_3"
    HEADING_4 = "heading_4"
    BODY_TEXT = "body_text"
    QUOTE = "quote"
    CAPTION = "caption"
    LIST_BULLET = "list_bullet"
    LIST_NUMBER = "list_number"
    TABLE_HEADER = "table_header"
    TABLE_BODY = "table_body"
    FOOTER = "footer"
    DATE_FIELD = "date_field"
    AMOUNT_FIELD = "amount_field"
    CLIENT_NAME_FIELD = "client_name_field"
    UNKNOWN = "unknown"


@dataclass
class FontStyle:
    """Opis czcionki - char style + run properties."""
    name: Optional[str] = None
    size_pt: Optional[float] = None           # w punktach
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[str] = None           # 'single', 'double', 'none'
    color: Optional[str] = None               # hex RGB, np. "FF0000"
    highlight: Optional[str] = None           # kolor podświetlenia
    small_caps: Optional[bool] = None
    all_caps: Optional[bool] = None
    strike: Optional[bool] = None
    subscript: Optional[bool] = None
    superscript: Optional[bool] = None

    def is_similar_to(self, other: "FontStyle", threshold: float = 0.8) -> bool:
        """Sprawdza czy dwa style czcionki są podobne."""
        if not isinstance(other, FontStyle):
            return False
        score = 0
        checks = 0
        if self.name and other.name:
            score += 1 if self.name.lower() == other.name.lower() else 0
            checks += 1
        if self.size_pt is not None and other.size_pt is not None:
            score += 1 if abs(self.size_pt - other.size_pt) < 0.5 else 0
            checks += 1
        if self.bold is not None and other.bold is not None:
            score += 1 if self.bold == other.bold else 0
            checks += 1
        if self.italic is not None and other.italic is not None:
            score += 1 if self.italic == other.italic else 0
            checks += 1
        if self.color and other.color:
            score += 1 if self.color.upper() == other.color.upper() else 0
            checks += 1
        if checks == 0:
            return False
        return (score / checks) >= threshold


@dataclass
class ParagraphSpacing:
    """Odstępy paragrafu."""
    before_pt: Optional[float] = None
    after_pt: Optional[float] = None
    line_spacing: Optional[float] = None      # 1.0 = single, 1.5 = one-half, 2.0 = double
    line_spacing_rule: Optional[str] = None   # 'auto', 'exact', 'atLeast'


@dataclass
class ParagraphBorder:
    """Ramka paragrafu."""
    top: Optional[Dict[str, Any]] = None      # {style, color, space, size}
    bottom: Optional[Dict[str, Any]] = None
    left: Optional[Dict[str, Any]] = None
    right: Optional[Dict[str, Any]] = None


@dataclass
class ParagraphAlignment:
    """Wyrównanie paragrafu."""
    alignment: Optional[str] = None           # 'left', 'right', 'center', 'justify'


@dataclass
class ParagraphIndentation:
    """Wcięcia paragrafu."""
    left_inches: Optional[float] = None
    right_inches: Optional[float] = None
    first_line_inches: Optional[float] = None
    hanging_inches: Optional[float] = None


@dataclass
class ParagraphStyle:
    """Pełny opis stylu paragrafowego."""
    style_id: str
    name: str
    based_on: Optional[str] = None            # styl bazowy
    next_style: Optional[str] = None          # następny styl
    font: FontStyle = field(default_factory=FontStyle)
    alignment: ParagraphAlignment = field(default_factory=ParagraphAlignment)
    spacing: ParagraphSpacing = field(default_factory=ParagraphSpacing)
    indentation: ParagraphIndentation = field(default_factory=ParagraphIndentation)
    border: Optional[ParagraphBorder] = None
    outline_level: Optional[int] = None       # 0 - Heading 1, 1 - Heading 2, ...
    is_default: bool = False
    is_built_in: bool = False
    hidden: bool = False
    quick_format: bool = False
    # Dodatkowe role semantyczne przypisane przez analyzer
    semantic_role: Optional[SemanticRole] = None
    role_confidence: float = 0.0


@dataclass
class CharacterStyle:
    """Styl znakowy (inline style)."""
    style_id: str
    name: str
    font: FontStyle = field(default_factory=FontStyle)
    based_on: Optional[str] = None


@dataclass
class TableStyle:
    """Styl tabeli."""
    style_id: str
    name: str
    # Simplifikacja - pełne wsparcie w v2
    border: Optional[Dict[str, Any]] = None
    shading: Optional[Dict[str, Any]] = None


@dataclass
class DocumentDefaults:
    """Domyślne ustawienia dokumentu."""
    margins: Dict[str, float] = field(default_factory=dict)   # top, bottom, left, right w inches
    page_width_inches: Optional[float] = None
    page_height_inches: Optional[float] = None
    orientation: str = "portrait"                               # 'portrait' | 'landscape'
    default_font: Optional[FontStyle] = None
    default_spacing: Optional[ParagraphSpacing] = None


@dataclass
class Theme:
    """Motyw kolorystyczny i czcionek."""
    color_scheme: Dict[str, str] = field(default_factory=dict)    # accent1, accent2, ... -> hex
    fonts: Dict[str, str] = field(default_factory=dict)           # majorFont, minorFont -> name


@dataclass 
class HeaderFooter:
    """Nagłówek lub stopka."""
    type: str                                                   # 'header', 'footer'
    section_type: str                                           # 'default', 'even', 'first'
    content: str = ""
    style_refs: List[str] = field(default_factory=list)
    contains_page_number: bool = False


@dataclass
class SectionBreak:
    """Podział sekcji."""
    break_type: str = "continuous"                              # 'continuous', 'nextPage', 'evenPage', 'oddPage'
    columns: int = 1
    margins: Optional[Dict[str, float]] = None


@dataclass
class TemplateProfile:
    """Profil template'u - wynik analizy Document A."""
    template_type: TemplateType = TemplateType.EXAMPLE_DOCUMENT
    paragraph_styles: Dict[str, ParagraphStyle] = field(default_factory=dict)
    character_styles: Dict[str, CharacterStyle] = field(default_factory=dict)
    table_styles: Dict[str, TableStyle] = field(default_factory=dict)
    document_defaults: DocumentDefaults = field(default_factory=DocumentDefaults)
    theme: Optional[Theme] = None
    headers_footers: List[HeaderFooter] = field(default_factory=list)
    sections: List[SectionBreak] = field(default_factory=list)
    # Placeholders wykryte w template (np. {{nazwa_firmy}})
    placeholders: Dict[str, str] = field(default_factory=dict)
    # Domyślna mapa semantyczna stylów
    semantic_styles: Dict[SemanticRole, str] = field(default_factory=dict)  # role -> style_id


@dataclass
class ParagraphContent:
    """Pojedynczy paragraf z dokumentu treści."""
    text: str
    style_name: Optional[str] = None
    runs: List[Dict[str, Any]] = field(default_factory=list)    # [{text, bold, italic, color}]
    is_list_item: bool = False
    list_level: int = 0
    list_num_id: Optional[int] = None
    indentation_left_inches: Optional[float] = None
    alignment: Optional[str] = None
    # Szybkie właściwości (run-level)
    has_bold: bool = False
    has_italic: bool = False
    has_underline: bool = False
    max_font_size: Optional[float] = None
    estimated_role: Optional[SemanticRole] = None


@dataclass
class TableContent:
    """Tabela z dokumentu treści."""
    rows: List[List[str]] = field(default_factory=list)
    style_name: Optional[str] = None
    num_rows: int = 0
    num_cols: int = 0


@dataclass
class ContentProfile:
    """Profil dokumentu treści - wynik analizy Document B."""
    paragraphs: List[ParagraphContent] = field(default_factory=list)
    tables: List[TableContent] = field(default_factory=list)
    images: List[Dict[str, Any]] = field(default_factory=list)
    # Struktura hierarchiczna
    structure_tree: Dict[str, Any] = field(default_factory=dict)
    # Metadane
    title: Optional[str] = None
    author: Optional[str] = None
    word_count: int = 0


@dataclass
class StyleMatch:
    """Wynik dopasowania stylu."""
    source_style_id: str
    target_style_id: str
    confidence: float                                              # 0.0 - 1.0
    reason: str
    matcher_type: str = "unknown"                                  # 'exact', 'fuzzy', 'semantic', 'llm', 'heuristic'


@dataclass
class ProcessingResult:
    """Wynik przetwarzania."""
    output_path: Optional[str] = None
    output_bytes: Optional[bytes] = None
    matched_styles: List[StyleMatch] = field(default_factory=list)
    unmatched_styles: List[str] = field(default_factory=list)
    applied_placeholders: Dict[str, str] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    success: bool = False
    processing_time_ms: Optional[int] = None
    preview_html: Optional[str] = None


__all__ = [
    "TemplateType",
    "SemanticRole",
    "FontStyle",
    "ParagraphSpacing",
    "ParagraphBorder",
    "ParagraphAlignment",
    "ParagraphIndentation",
    "ParagraphStyle",
    "CharacterStyle",
    "TableStyle",
    "DocumentDefaults",
    "Theme",
    "HeaderFooter",
    "SectionBreak",
    "TemplateProfile",
    "ParagraphContent",
    "TableContent",
    "ContentProfile",
    "StyleMatch",
    "ProcessingResult",
]
