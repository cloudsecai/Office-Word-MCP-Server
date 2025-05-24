"""
MCP tool implementations for the Word Document Server.

This package contains the MCP tool implementations that expose functionality
to clients through the Model Context Protocol.
"""

# Document tools
from word_document_server.tools.document_tools import (
    create_document, get_document_info, get_document_text, 
    get_document_outline, list_available_documents, 
    copy_document, merge_documents
)

# Content tools
from word_document_server.tools.content_tools import (
    add_heading, add_paragraph, add_table, add_picture,
    add_page_break, add_table_of_contents, delete_paragraph,
    search_and_replace
)

# Format tools
from word_document_server.tools.format_tools import (
    format_text, create_custom_style, format_table
)

# Protection tools
from word_document_server.tools.protection_tools import (
    protect_document, add_restricted_editing,
    add_digital_signature, verify_document
)

# Footnote tools
from word_document_server.tools.footnote_tools import (
    add_footnote_to_document, add_endnote_to_document,
    convert_footnotes_to_endnotes_in_document, customize_footnote_style
)

# Extended document tools
from word_document_server.tools.extended_document_tools import (
    get_paragraph_text_from_document, find_text_in_document,
    convert_to_pdf, get_document_structure_details_from_document,
    get_table_cell_content_from_document, set_table_cell_text,
    set_paragraph_text, insert_paragraph_after_index,
    clear_table_cell_content, add_paragraph_to_table_cell,
    search_and_replace_in_scope, is_element_empty
)

__all__ = [
    # Document tools
    'create_document', 'get_document_info', 'get_document_text',
    'get_document_outline', 'list_available_documents', 
    'copy_document', 'merge_documents',
    
    # Content tools
    'add_heading', 'add_paragraph', 'add_table', 'add_picture',
    'add_page_break', 'add_table_of_contents', 'delete_paragraph',
    'search_and_replace',
    
    # Format tools
    'format_text', 'create_custom_style', 'format_table',
    
    # Protection tools
    'protect_document', 'add_restricted_editing',
    'add_digital_signature', 'verify_document',
    
    # Footnote tools
    'add_footnote_to_document', 'add_endnote_to_document',
    'convert_footnotes_to_endnotes_in_document', 'customize_footnote_style',
    
    # Extended document tools
    'get_paragraph_text_from_document', 'find_text_in_document',
    'convert_to_pdf', 'get_document_structure_details_from_document',
    'get_table_cell_content_from_document', 'set_table_cell_text',
    'set_paragraph_text', 'insert_paragraph_after_index',
    'clear_table_cell_content', 'add_paragraph_to_table_cell',
    'search_and_replace_in_scope', 'is_element_empty'
]
