"""
Simplified editing utilities using FormattedEditor class.
"""

from typing import Dict, Any
from .formatted_editor import FormattedEditor, ScopeLocation


def search_and_replace_in_scope_util(
    doc_path: str,
    find_text_str: str,
    replace_text_str: str,
    scope_type: str,
    scope_identifier: Dict[str, int],
) -> Dict[str, Any]:
    """
    Search and replace text within a specific scope (paragraph or table cell)
    in a Word document, preserving run-level formatting.
    """
    editor = FormattedEditor(doc_path)

    # Create scope location object
    if scope_type == "paragraph":
        scope = ScopeLocation(
            scope_type=scope_type,
            paragraph_index=scope_identifier.get("paragraph_index"),
        )
    elif scope_type == "table_cell":
        scope = ScopeLocation(
            scope_type=scope_type,
            table_index=scope_identifier.get("table_index"),
            row_index=scope_identifier.get("row_index"),
            col_index=scope_identifier.get("col_index"),
        )
    else:
        return {
            "error": f"Invalid scope_type: {scope_type}. Must be 'paragraph' or 'table_cell'"
        }

    return editor.search_and_replace_in_scope(find_text_str, replace_text_str, scope)


"""
End of editing_utils.py
"""
