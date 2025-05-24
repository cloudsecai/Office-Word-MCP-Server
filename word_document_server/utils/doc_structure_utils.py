"""
Simplified document structure utilities using DocumentAnalyzer class.
"""

from typing import Dict, Any
from .document_analyzer import DocumentAnalyzer


def get_document_structure_details(doc_path: str) -> Dict[str, Any]:
    """
    Get detailed structure information about a Word document including
    paragraphs, tables, styles, and run-level formatting.

    Args:
        doc_path: Path to the Word document

    Returns:
        Dictionary with comprehensive document structure details
    """
    analyzer = DocumentAnalyzer(doc_path)
    return analyzer.get_complete_structure()


def find_text(
    doc_path: str, text_to_find: str, match_case: bool = True, whole_word: bool = False
) -> Dict[str, Any]:
    """
    Find all occurrences of specific text in a Word document.

    Args:
        doc_path: Path to the Word document
        text_to_find: Text to search for
        match_case: Whether to perform case-sensitive search
        whole_word: Whether to match whole words only

    Returns:
        Dictionary with search results
    """
    analyzer = DocumentAnalyzer(doc_path)
    return analyzer.find_text(text_to_find, match_case, whole_word)


def is_element_empty_util(
    doc_path: str, element_type: str, element_identifier: Dict[str, int]
) -> Dict[str, Any]:
    """
    Check if a specific element (paragraph or table cell) is empty.

    Args:
        doc_path: Path to the Word document
        element_type: Type of element ("paragraph" or "table_cell")
        element_identifier: Dictionary with element location info

    Returns:
        Dictionary indicating whether element is empty
    """
    analyzer = DocumentAnalyzer(doc_path)

    try:
        if element_type == "paragraph":
            if "paragraph_index" not in element_identifier:
                return {
                    "error": "Missing 'paragraph_index' in element_identifier for paragraph type"
                }

            paragraphs = analyzer.get_paragraphs_analysis()
            if not paragraphs:
                return {"error": "Failed to analyze document paragraphs"}

            para_index = element_identifier["paragraph_index"]
            if not (0 <= para_index < len(paragraphs)):
                return {
                    "error": f"Invalid paragraph index: {para_index}. Document has {len(paragraphs)} paragraphs."
                }

            paragraph = paragraphs[para_index]
            is_empty = not paragraph["text"].strip()

            return {
                "element_type": element_type,
                "element_identifier": element_identifier,
                "is_empty": is_empty,
                "text_length": len(paragraph["text"]),
                "has_content": bool(paragraph["text"].strip()),
            }

        elif element_type == "table_cell":
            required_keys = ["table_index", "row_index", "col_index"]
            for key in required_keys:
                if key not in element_identifier:
                    return {
                        "error": f"Missing '{key}' in element_identifier for table_cell type"
                    }

            tables = analyzer.get_tables_analysis()
            if not tables:
                return {"error": "Failed to analyze document tables"}

            table_index = element_identifier["table_index"]
            if not (0 <= table_index < len(tables)):
                return {
                    "error": f"Invalid table index: {table_index}. Document has {len(tables)} tables."
                }

            table = tables[table_index]
            row_index = element_identifier["row_index"]
            col_index = element_identifier["col_index"]

            if not (0 <= row_index < len(table["cells"])):
                return {
                    "error": f"Invalid row index: {row_index}. Table has {len(table['cells'])} rows."
                }

            if not (0 <= col_index < len(table["cells"][row_index])):
                return {
                    "error": f"Invalid column index: {col_index}. Row has {len(table['cells'][row_index])} columns."
                }

            cell = table["cells"][row_index][col_index]
            is_empty = not cell["text"].strip()

            return {
                "element_type": element_type,
                "element_identifier": element_identifier,
                "is_empty": is_empty,
                "text_length": len(cell["text"]),
                "has_content": bool(cell["text"].strip()),
                "paragraph_count": len(cell["paragraphs"]),
            }

        else:
            return {
                "error": f"Invalid element_type: {element_type}. Must be 'paragraph' or 'table_cell'"
            }

    except Exception as e:
        return {"error": f"Failed to check if element is empty: {str(e)}"}


"""
# End of doc_structure_utils.py
"""
