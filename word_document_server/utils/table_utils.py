"""
Simplified table utilities using TableManager class.

Note: For new code, prefer using TableManager class directly from table_manager.py
These functions are provided for backward compatibility.
"""

from typing import Dict, Any, Optional
from .table_manager import TableManager, CellLocation


def get_table_cell_content(
    doc_path: str, table_index: int, row_index: int, col_index: int
) -> Dict[str, Any]:
    """
    Get detailed content from a specific table cell.

    Note: For new code, prefer using TableManager class directly:
        table_manager = TableManager(doc_path)
        location = CellLocation(table_index, row_index, col_index)
        result = table_manager.get_cell_content(location)
    """
    table_manager = TableManager(doc_path)
    location = CellLocation(table_index, row_index, col_index)
    return table_manager.get_cell_content(location)


def set_table_cell_text_util(
    doc_path: str,
    table_index: int,
    row_index: int,
    col_index: int,
    text_to_set: str,
    clear_existing_content: bool = True,
    paragraph_style: Optional[str] = None,
) -> Dict[str, Any]:
    """Set text in a specific table cell in a Word document."""
    table_manager = TableManager(doc_path)
    location = CellLocation(table_index, row_index, col_index)
    return table_manager.set_cell_text(
        location, text_to_set, clear_existing_content, paragraph_style
    )


def clear_table_cell_content_util(
    doc_path: str, table_index: int, row_index: int, col_index: int
) -> Dict[str, Any]:
    """Clear all content from a specific table cell in a Word document."""
    table_manager = TableManager(doc_path)
    location = CellLocation(table_index, row_index, col_index)
    return table_manager.clear_cell_content(location)


def add_paragraph_to_table_cell_util(
    doc_path: str,
    table_index: int,
    row_index: int,
    col_index: int,
    paragraph_text: str,
    paragraph_style: Optional[str] = None,
) -> Dict[str, Any]:
    """Add a new paragraph to a specific table cell in a Word document."""
    table_manager = TableManager(doc_path)
    location = CellLocation(table_index, row_index, col_index)
    return table_manager.add_paragraph_to_cell(
        location, paragraph_text, paragraph_style
    )


"""
# End of table_utils.py
"""
