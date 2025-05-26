"""
Document utility functions for Word Document Server.
"""

from typing import Dict, Any
from docx import Document


def get_document_properties(doc_path: str) -> Dict[str, Any]:
    """Get properties of a Word document."""
    import os

    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}

    try:
        doc = Document(doc_path)
        core_props = doc.core_properties

        return {
            "title": core_props.title or "",
            "author": core_props.author or "",
            "subject": core_props.subject or "",
            "keywords": core_props.keywords or "",
            "created": str(core_props.created) if core_props.created else "",
            "modified": str(core_props.modified) if core_props.modified else "",
            "last_modified_by": core_props.last_modified_by or "",
            "revision": core_props.revision or 0,
            "page_count": len(doc.sections),
            "word_count": sum(
                len(paragraph.text.split()) for paragraph in doc.paragraphs
            ),
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables),
        }
    except Exception as e:
        return {"error": f"Failed to get document properties: {str(e)}"}


def extract_document_text(doc_path: str) -> str:
    """Extract all text from a Word document with structured table formatting for LLM parsing.

    Extracts text while preserving table structure with clear row/cell boundaries
    to maintain relationships between cells (essential for Q&A pairs, forms, etc.).

    Args:
        doc_path: Path to the Word document

    Returns:
        Structured text with clear table formatting for optimal LLM parsing
    """
    import os

    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"

    try:
        from word_document_server.utils.document_analyzer import DocumentAnalyzer

        doc = Document(doc_path)
        analyzer = DocumentAnalyzer(doc_path)
        structure = analyzer.get_complete_structure()

        if "error" in structure:
            return structure["error"]

        text_parts = []
        table_index = 0

        # Get all body elements and process them in order
        for element in doc.element.body:
            tag_name = element.tag.split("}")[-1] if "}" in element.tag else element.tag

            if tag_name == "p":  # paragraph
                # Find the corresponding paragraph in the parsed document
                for para in doc.paragraphs:
                    if para._element == element:
                        para_text = para.text.strip()
                        if para_text:  # Only add non-empty paragraphs
                            text_parts.append(para_text)
                        break

            elif tag_name == "tbl":  # table
                if table_index < len(structure["tables"]):
                    table_text = _format_table_for_llm(
                        structure["tables"][table_index], table_index
                    )
                    text_parts.append(table_text)
                    table_index += 1

        return "\n\n".join(text_parts)

    except Exception:
        # Fallback to simple extraction if structured fails
        doc = Document(doc_path)
        text = []

        for paragraph in doc.paragraphs:
            text.append(paragraph.text)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        text.append(paragraph.text)

        return "\n".join(text)


def _format_table_for_llm(table_data: Dict[str, Any], table_index: int) -> str:
    """Format table data with clear row/cell boundaries for LLM parsing."""
    if not table_data.get("cells"):
        return f"=== TABLE {table_index + 1} (empty) ==="

    formatted_lines = [f"=== TABLE {table_index + 1} ==="]

    for row_idx, row_cells in enumerate(table_data["cells"]):
        row_parts = []
        for col_idx, cell in enumerate(row_cells):
            cell_text = cell.get("text", "").strip()
            # Handle merged cells
            if cell.get("v_merge") == "continue":
                cell_text = "(merged with above)"
            elif not cell_text:
                cell_text = "(empty)"

            row_parts.append(f"Col{col_idx}: {cell_text}")

        row_line = f"Row{row_idx}: | " + " | ".join(row_parts) + " |"
        formatted_lines.append(row_line)

    formatted_lines.append("=== END TABLE ===")
    return "\n".join(formatted_lines)


def get_document_structure(doc_path: str) -> Dict[str, Any]:
    """Get the structure of a Word document."""
    import os

    if not os.path.exists(doc_path):
        return {"error": f"Document {doc_path} does not exist"}

    try:
        doc = Document(doc_path)
        structure = {"paragraphs": [], "tables": []}

        # Get paragraphs
        for i, para in enumerate(doc.paragraphs):
            structure["paragraphs"].append(
                {
                    "index": i,
                    "text": para.text[:100] + ("..." if len(para.text) > 100 else ""),
                    "style": para.style.name if para.style else "Normal",
                }
            )

        # Get tables
        for i, table in enumerate(doc.tables):
            table_data = {
                "index": i,
                "rows": len(table.rows),
                "columns": len(table.columns),
                "preview": [],
            }

            # Get sample of table data
            max_rows = min(3, len(table.rows))
            for row_idx in range(max_rows):
                row_data = []
                max_cols = min(3, len(table.columns))
                for col_idx in range(max_cols):
                    try:
                        cell_text = table.cell(row_idx, col_idx).text
                        row_data.append(
                            cell_text[:20] + ("..." if len(cell_text) > 20 else "")
                        )
                    except IndexError:
                        row_data.append("N/A")
                table_data["preview"].append(row_data)

            structure["tables"].append(table_data)

        return structure
    except Exception as e:
        return {"error": f"Failed to get document structure: {str(e)}"}


def find_paragraph_by_text(doc, text, partial_match=False):
    """
    Find paragraphs containing specific text.

    Args:
        doc: Document object
        text: Text to search for
        partial_match: If True, matches paragraphs containing the text; if False, matches exact text

    Returns:
        List of paragraph indices that match the criteria
    """
    matching_paragraphs = []

    for i, para in enumerate(doc.paragraphs):
        if partial_match and text in para.text:
            matching_paragraphs.append(i)
        elif not partial_match and para.text == text:
            matching_paragraphs.append(i)

    return matching_paragraphs


def find_and_replace_text(doc, old_text, new_text):
    """
    Find and replace text throughout the document.

    Args:
        doc: Document object
        old_text: Text to find
        new_text: Text to replace with

    Returns:
        Number of replacements made
    """
    count = 0

    # Search in paragraphs
    for para in doc.paragraphs:
        if old_text in para.text:
            for run in para.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
                    count += 1

    # Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if old_text in para.text:
                        for run in para.runs:
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_text)
                                count += 1

    return count
