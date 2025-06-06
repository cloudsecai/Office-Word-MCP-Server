"""
Utility functions for the Word Document Server.

This package contains utility modules for file operations and document handling.
"""

from word_document_server.utils.file_utils import (
    check_file_writeable,
    create_document_copy,
    ensure_docx_extension,
)
from word_document_server.utils.document_utils import (
    get_document_properties,
    extract_document_text,
    get_document_structure,
    find_paragraph_by_text,
    find_and_replace_text,
)

# New class-based utilities
from word_document_server.utils.table_manager import TableManager, CellLocation
from word_document_server.utils.document_analyzer import DocumentAnalyzer
from word_document_server.utils.formatted_editor import FormattedEditor, ScopeLocation

__all__ = [
    # File utilities
    "check_file_writeable",
    "create_document_copy",
    "ensure_docx_extension",
    # Document utilities
    "get_document_properties",
    "extract_document_text",
    "get_document_structure",
    "find_paragraph_by_text",
    "find_and_replace_text",
    # New class-based utilities
    "TableManager",
    "CellLocation",
    "DocumentAnalyzer",
    "FormattedEditor",
    "ScopeLocation",
]
