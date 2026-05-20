"""
DOCX Formatter - Core engine for intelligent document formatting.
"""

from .types import *
from .extractor import DOCXExtractor
from .matcher import StyleMatchingEngine
from .assembler import DocumentAssembler
from .pipeline import FormatPipeline

__version__ = "0.1.0"

__all__ = [
    "FormatPipeline",
    "DOCXExtractor", 
    "StyleMatchingEngine",
    "DocumentAssembler",
]
