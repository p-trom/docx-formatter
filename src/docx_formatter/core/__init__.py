"""
DOCX Formatter - Core engine for intelligent document formatting.
"""

from .assembler import DocumentAssembler
from .extractor import DOCXExtractor
from .matcher import StyleMatchingEngine
from .pipeline import FormatPipeline
from .types import *  # noqa: F403

__version__ = "0.1.0"

__all__ = [
    "FormatPipeline",
    "DOCXExtractor",
    "StyleMatchingEngine",
    "DocumentAssembler",
]
