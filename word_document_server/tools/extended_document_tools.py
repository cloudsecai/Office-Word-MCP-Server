"""
Extended document manipulation tools for Word Document Server.

This module provides advanced document manipulation capabilities including
detailed content inspection, targeted editing, and scope-specific operations.
"""

import os
import json
import subprocess
import platform
import shutil
from typing import Optional

from word_document_server.utils.file_utils import (
    check_file_writeable,
    ensure_docx_extension,
)
from word_document_server.utils.paragraph_utils import (
    get_paragraph_text,
    set_paragraph_text_util,
    insert_paragraph_after_index_util,
)

from word_document_server.utils.document_analyzer import DocumentAnalyzer
from word_document_server.utils.table_manager import TableManager, CellLocation
from word_document_server.utils.formatted_editor import FormattedEditor, ScopeLocation


def _validate_file_exists(filename: str) -> Optional[str]:
    """Validate that a file exists and return user-friendly error if not."""
    if not os.path.exists(filename):
        return f"The document '{filename}' could not be found. Please check the file path and try again."
    return None


def _validate_non_negative_index(value: int, param_name: str) -> Optional[str]:
    """Validate that an index is non-negative and return user-friendly error if not."""
    if value < 0:
        return f"The {param_name} must be 0 or greater (you provided {value}). Indexes start from 0."
    return None


def _validate_table_coordinates(
    table_index: int, row_index: int, col_index: int
) -> Optional[str]:
    """Validate table coordinates and return user-friendly error if invalid."""
    if table_index < 0:
        return f"The table_index must be 0 or greater (you provided {table_index}). Table indexes start from 0."
    if row_index < 0:
        return f"The row_index must be 0 or greater (you provided {row_index}). Row indexes start from 0."
    if col_index < 0:
        return f"The col_index must be 0 or greater (you provided {col_index}). Column indexes start from 0."
    return None


def _check_file_writable(filename: str) -> Optional[str]:
    """Check if file is writable and return user-friendly error if not."""
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify the document: {error_message}. Try creating a copy first or check file permissions."
    return None


def _validate_scope_identifier(
    scope_type: str, scope_identifier: dict
) -> Optional[str]:
    """Validate scope identifier structure and return user-friendly error if invalid."""
    if not isinstance(scope_identifier, dict):
        return "The scope_identifier must be a dictionary. See the function documentation for examples."

    if scope_type == "paragraph":
        if "paragraph_index" not in scope_identifier:
            return 'For paragraph scope, use: {"paragraph_index": 0}'
    elif scope_type == "table_cell":
        required_keys = ["table_index", "row_index", "col_index"]
        missing_keys = [key for key in required_keys if key not in scope_identifier]
        if missing_keys:
            return f'For table_cell scope, use: {{"table_index": 0, "row_index": 1, "col_index": 2}}. Missing: {missing_keys}'

    return None


async def get_paragraph_text_from_document(filename: str, paragraph_index: int) -> str:
    """Get text content from a specific paragraph in a Word document (.docx only).

    Retrieves the complete text content and metadata from a single paragraph by index.
    Useful for reading specific parts of a document without loading all content.

    Use this tool when:
    - Need to read just one specific paragraph from a document
    - Want paragraph metadata (style, formatting details)
    - Checking content of a particular paragraph before editing
    - Reading document content paragraph by paragraph

    Args:
        filename: Path to the Word document (.docx format only)
        paragraph_index: Index of the paragraph to read (0-based, first paragraph = 0)

    Returns:
        JSON string containing:
        - paragraph text content
        - style information
        - formatting details
        - paragraph index
        - error message if paragraph doesn't exist or can't be read

    Limitations:
        - Only works with .docx format (Microsoft Word 2007+)
        - Cannot read password-protected documents
        - Paragraph index must exist in document
        - Does not include tables (tables are separate from paragraphs)

    Example:
        get_paragraph_text_from_document("report.docx", 0)
        # Returns first paragraph text and metadata
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if error := _validate_non_negative_index(paragraph_index, "paragraph_index"):
        return error

    try:
        result = get_paragraph_text(filename, paragraph_index)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Unable to read paragraph text: {str(e)}"


async def find_text_in_document(
    filename: str, text_to_find: str, match_case: bool = True, whole_word: bool = False
) -> str:
    """Find all occurrences of text in a Word document (.docx only) with precise location details.

    Searches through all paragraphs and table cells to find text matches. Returns detailed
    location information for each match including paragraph/table/cell coordinates.

    Use this tool when:
    - Need to locate specific text before editing or replacing
    - Want to know exact positions of text in document structure
    - Searching for text across both paragraphs and tables
    - Need context around found text for verification

    Args:
        filename: Path to the Word document (.docx format only)
        text_to_find: Text string to search for (cannot be empty)
        match_case: Whether search should be case-sensitive (default: True)
        whole_word: Whether to match whole words only (default: False)

    Returns:
        JSON string containing:
        - query details (search text, options used)
        - array of all matches with locations
        - total count of matches
        - context text around each match
        - table/paragraph location details for each match

    Search locations:
        - All paragraph text content
        - All table cell text content
        - Provides coordinates for both paragraph and table matches

    Limitations:
        - Only works with .docx format (Microsoft Word 2007+)
        - Cannot search password-protected documents
        - Does not search headers, footers, or footnotes
        - Does not search embedded objects or images
        - Whole word matching uses simple word boundary detection

    Example:
        find_text_in_document("report.docx", "summary", match_case=False)
        # Finds all case-insensitive matches of "summary"
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if not text_to_find.strip():
        return "Please provide text to search for. Empty search text is not allowed."

    try:
        analyzer = DocumentAnalyzer(filename)
        result = analyzer.find_text(text_to_find, match_case, whole_word)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Search failed: {str(e)}"


async def convert_to_pdf(filename: str, output_filename: Optional[str] = None) -> str:
    """Convert a Word document (.docx only) to PDF format using available system tools.

    Attempts to convert using platform-appropriate tools: Microsoft Word on Windows,
    LibreOffice on Linux/macOS. Requires appropriate software to be installed.

    Use this tool when:
    - Need to create PDF version of Word document
    - Want to preserve document formatting in PDF
    - Converting for sharing or archival purposes
    - Need non-editable version of document

    Args:
        filename: Path to the Word document (.docx format only)
        output_filename: Optional path for output PDF (auto-generates if not provided)

    Returns:
        Success message with PDF path, or error message with details

    Conversion process:
        - Windows: Uses docx2pdf library (requires Microsoft Word)
        - Linux/macOS: Uses LibreOffice command line (requires LibreOffice)
        - Fallback: Attempts pandoc if available

    Requirements by platform:
        - Windows: Microsoft Word must be installed
        - Linux/macOS: LibreOffice must be installed
        - Alternative: pandoc (cross-platform)

    Limitations:
        - Only works with .docx format (Microsoft Word 2007+)
        - Requires external software for conversion
        - Complex formatting may not convert perfectly
        - Cannot convert password-protected documents
        - Output quality depends on conversion tool used

    Example:
        convert_to_pdf("report.docx", "report.pdf")
        # Creates PDF version of the Word document
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    # Generate output filename if not provided
    if not output_filename:
        base_name, _ = os.path.splitext(filename)
        output_filename = f"{base_name}.pdf"
    elif not output_filename.lower().endswith(".pdf"):
        output_filename = f"{output_filename}.pdf"

    # Convert to absolute path if not already
    if not os.path.isabs(output_filename):
        output_filename = os.path.abspath(output_filename)

    # Ensure the output directory exists
    output_dir = os.path.dirname(output_filename)
    if not output_dir:
        output_dir = os.path.abspath(".")

    # Create the directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Check if output file can be written
    is_writeable, error_message = check_file_writeable(output_filename)
    if not is_writeable:
        return f"Cannot create PDF: {error_message} (Path: {output_filename}, Dir: {output_dir})"

    try:
        # Determine platform for appropriate conversion method
        system = platform.system()

        if system == "Windows":
            # On Windows, try docx2pdf which uses Microsoft Word
            try:
                from docx2pdf import convert

                convert(filename, output_filename)
                return f"Document successfully converted to PDF: {output_filename}"
            except (ImportError, Exception) as e:
                return f"Failed to convert document to PDF: {str(e)}\nNote: docx2pdf requires Microsoft Word to be installed."

        elif system in ["Linux", "Darwin"]:  # Linux or macOS
            # Try using LibreOffice if available (common on Linux/macOS)
            try:
                # Choose the appropriate command based on OS
                if system == "Darwin":  # macOS
                    lo_commands = [
                        "soffice",
                        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                    ]
                else:  # Linux
                    lo_commands = ["libreoffice", "soffice"]

                # Try each possible command
                conversion_successful = False
                errors = []

                for cmd_name in lo_commands:
                    try:
                        # Construct LibreOffice conversion command
                        output_dir = os.path.dirname(output_filename)
                        # If output_dir is empty, use current directory
                        if not output_dir:
                            output_dir = "."
                        # Ensure the directory exists
                        os.makedirs(output_dir, exist_ok=True)

                        cmd = [
                            cmd_name,
                            "--headless",
                            "--convert-to",
                            "pdf",
                            "--outdir",
                            output_dir,
                            filename,
                        ]

                        result = subprocess.run(
                            cmd, capture_output=True, text=True, timeout=60
                        )

                        if result.returncode == 0:
                            # LibreOffice creates the PDF with the same basename
                            base_name = os.path.basename(filename)
                            pdf_base_name = os.path.splitext(base_name)[0] + ".pdf"
                            created_pdf = os.path.join(
                                os.path.dirname(output_filename) or ".", pdf_base_name
                            )

                            # If the created PDF is not at the desired location, move it
                            if created_pdf != output_filename and os.path.exists(
                                created_pdf
                            ):
                                shutil.move(created_pdf, output_filename)

                            conversion_successful = True
                            break  # Exit the loop if successful
                        else:
                            errors.append(f"{cmd_name} error: {result.stderr}")
                    except (subprocess.SubprocessError, FileNotFoundError) as e:
                        errors.append(f"{cmd_name} error: {str(e)}")

                if conversion_successful:
                    return f"Document successfully converted to PDF: {output_filename}"
                else:
                    # If all LibreOffice attempts failed, try docx2pdf as fallback
                    try:
                        from docx2pdf import convert

                        convert(filename, output_filename)
                        return (
                            f"Document successfully converted to PDF: {output_filename}"
                        )
                    except (ImportError, Exception) as e:
                        error_msg = "Failed to convert document to PDF using LibreOffice or docx2pdf.\n"
                        error_msg += "LibreOffice errors: " + "; ".join(errors) + "\n"
                        error_msg += f"docx2pdf error: {str(e)}\n"
                        error_msg += (
                            "To convert documents to PDF, please install either:\n"
                        )
                        error_msg += "1. LibreOffice (recommended for Linux/macOS)\n"
                        error_msg += (
                            "2. Microsoft Word (required for docx2pdf on Windows/macOS)"
                        )
                        return error_msg

            except Exception as e:
                return f"Failed to convert document to PDF: {str(e)}"
        else:
            return f"PDF conversion not supported on {system} platform"

    except Exception as e:
        return f"Failed to convert document to PDF: {str(e)}"


async def get_document_structure_details_from_document(filename: str) -> str:
    """Get comprehensive structure details of a Word document including paragraphs, tables, styles, and run-level formatting.

    This function provides deep analysis of a document's structure, useful for understanding
    content layout, formatting patterns, and preparing for targeted modifications.

    Args:
        filename: Path to the Word document

    Returns:
        JSON string containing detailed document structure information

    Example:
        get_document_structure_details_from_document("report.docx")
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    try:
        analyzer = DocumentAnalyzer(filename)
        result = analyzer.get_complete_structure()
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Unable to get document structure: {str(e)}"


async def get_table_cell_content_from_document(
    filename: str, table_index: int, row_index: int, col_index: int
) -> str:
    """Get detailed content and formatting from a specific table cell in a Word document (.docx only).

    Retrieves comprehensive information about a single table cell including text content,
    paragraph details, formatting information, and cell properties like merged cell status.

    Use this tool when:
    - Need to read content from a specific table cell
    - Want detailed formatting information for a cell
    - Checking cell properties (merged cells, spanning)
    - Reading table data before making modifications

    Args:
        filename: Path to the Word document (.docx format only)
        table_index: Index of the table (0-based, first table = 0)
        row_index: Index of the row (0-based, first row = 0)
        col_index: Index of the column (0-based, first column = 0)

    Returns:
        JSON string containing:
        - cell text content (combined from all paragraphs)
        - detailed paragraph information with formatting
        - cell location coordinates
        - grid_span information (if cell spans multiple columns)
        - v_merge information (if cell is part of vertical merge)
        - run-level formatting details (bold, italic, font, etc.)
        - error message if coordinates are invalid

    Cell properties explained:
        - grid_span: number of columns this cell spans (1 = normal, >1 = spans multiple columns)
        - v_merge: vertical merge status (None = normal, "continue" = merged with cell above)

    Limitations:
        - Only works with .docx format (Microsoft Word 2007+)
        - Cannot read password-protected documents
        - Table, row, and column indexes must exist in document
        - Does not include images or embedded objects in cells

    Example:
        get_table_cell_content_from_document("report.docx", 0, 1, 2)
        # Gets content from first table, second row, third column
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if error := _validate_table_coordinates(table_index, row_index, col_index):
        return error

    try:
        table_manager = TableManager(filename)
        location = CellLocation(table_index, row_index, col_index)
        result = table_manager.get_cell_content(location)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Unable to read table cell content: {str(e)}"


async def set_table_cell_text(
    filename: str,
    table_index: int,
    row_index: int,
    col_index: int,
    text_to_set: str,
    clear_existing_content: bool = True,
    paragraph_style: Optional[str] = None,
) -> str:
    """Set text in a specific table cell in a Word document.

    Args:
        filename: Path to the Word document
        table_index: Index of the table (0-based)
        row_index: Index of the row (0-based)
        col_index: Index of the column (0-based)
        text_to_set: Text to set in the cell
        clear_existing_content: Whether to clear existing content first (default: True)
        paragraph_style: Optional paragraph style to apply (e.g., "Normal", "Heading 1")

    Returns:
        Success message or error description

    Example:
        set_table_cell_text("report.docx", 0, 1, 2, "New content", True, "Normal")
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if error := _validate_table_coordinates(table_index, row_index, col_index):
        return error

    if error := _check_file_writable(filename):
        return error

    try:
        table_manager = TableManager(filename)
        location = CellLocation(table_index, row_index, col_index)
        result = table_manager.set_cell_text(
            location, text_to_set, clear_existing_content, paragraph_style
        )
        if "error" in result:
            return f"Failed to update table cell: {result['error']}"
        return result["message"]
    except Exception as e:
        return f"Unable to set table cell text: {str(e)}"


async def set_paragraph_text(
    filename: str,
    paragraph_index: int,
    new_text: str,
    style_to_apply: Optional[str] = None,
) -> str:
    """Set text in a specific paragraph in a Word document.

    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph (0-based)
        new_text: New text to set
        style_to_apply: Optional style to apply to the paragraph (e.g., "Normal", "Heading 1")

    Returns:
        Success message or error description

    Example:
        set_paragraph_text("report.docx", 0, "New paragraph content", "Heading 1")
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if error := _validate_non_negative_index(paragraph_index, "paragraph_index"):
        return error

    if error := _check_file_writable(filename):
        return error

    try:
        result = set_paragraph_text_util(
            filename, paragraph_index, new_text, style_to_apply
        )
        if "error" in result:
            return f"Failed to update paragraph: {result['error']}"
        return result["message"]
    except Exception as e:
        return f"Unable to set paragraph text: {str(e)}"


async def insert_paragraph_after_index(
    filename: str,
    target_paragraph_index: int,
    text_to_insert: str,
    style_to_apply: Optional[str] = None,
) -> str:
    """Insert a new paragraph after a specific paragraph index in a Word document.

    Args:
        filename: Path to the Word document
        target_paragraph_index: Index of the paragraph after which to insert (0-based)
        text_to_insert: Text for the new paragraph
        style_to_apply: Optional style to apply to the new paragraph (e.g., "Normal", "Heading 1")

    Returns:
        Success message or error description

    Example:
        insert_paragraph_after_index("report.docx", 0, "New paragraph", "Normal")
        # Inserts new paragraph after the first paragraph
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if error := _validate_non_negative_index(
        target_paragraph_index, "target_paragraph_index"
    ):
        return error

    if error := _check_file_writable(filename):
        return error

    try:
        result = insert_paragraph_after_index_util(
            filename, target_paragraph_index, text_to_insert, style_to_apply
        )
        if "error" in result:
            return f"Failed to insert paragraph: {result['error']}"
        return result["message"]
    except Exception as e:
        return f"Unable to insert paragraph: {str(e)}"


async def clear_table_cell_content(
    filename: str, table_index: int, row_index: int, col_index: int
) -> str:
    """Clear all content from a specific table cell in a Word document.

    Args:
        filename: Path to the Word document
        table_index: Index of the table (0-based)
        row_index: Index of the row (0-based)
        col_index: Index of the column (0-based)

    Returns:
        Success message or error description

    Example:
        clear_table_cell_content("report.docx", 0, 1, 2)
        # Clears content from first table, second row, third column
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if error := _validate_table_coordinates(table_index, row_index, col_index):
        return error

    if error := _check_file_writable(filename):
        return error

    try:
        table_manager = TableManager(filename)
        location = CellLocation(table_index, row_index, col_index)
        result = table_manager.clear_cell_content(location)
        if "error" in result:
            return f"Failed to clear cell content: {result['error']}"
        return result["message"]
    except Exception as e:
        return f"Unable to clear table cell content: {str(e)}"


async def add_paragraph_to_table_cell(
    filename: str,
    table_index: int,
    row_index: int,
    col_index: int,
    paragraph_text: str,
    paragraph_style: Optional[str] = None,
) -> str:
    """Add a new paragraph to a specific table cell in a Word document.

    Args:
        filename: Path to the Word document
        table_index: Index of the table (0-based)
        row_index: Index of the row (0-based)
        col_index: Index of the column (0-based)
        paragraph_text: Text for the new paragraph
        paragraph_style: Optional style to apply to the new paragraph (e.g., "Normal", "List Paragraph")

    Returns:
        Success message or error description

    Example:
        add_paragraph_to_table_cell("report.docx", 0, 1, 2, "Additional content", "Normal")
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if error := _validate_table_coordinates(table_index, row_index, col_index):
        return error

    if error := _check_file_writable(filename):
        return error

    try:
        table_manager = TableManager(filename)
        location = CellLocation(table_index, row_index, col_index)
        result = table_manager.add_paragraph_to_cell(
            location, paragraph_text, paragraph_style
        )
        if "error" in result:
            return f"Failed to add paragraph to cell: {result['error']}"
        return result["message"]
    except Exception as e:
        return f"Unable to add paragraph to table cell: {str(e)}"


async def search_and_replace_in_scope(
    filename: str,
    find_text: str,
    replace_text: str,
    scope_type: str,
    scope_identifier: dict,
) -> str:
    """Search and replace text within a specific scope (paragraph or table cell) in a Word document.

    This function performs targeted text replacement while preserving formatting within
    the specified scope. Useful for making precise edits without affecting the entire document.

    Args:
        filename: Path to the Word document
        find_text: Text to find and replace
        replace_text: New text to replace with
        scope_type: Type of scope ("paragraph" or "table_cell")
        scope_identifier: Dictionary identifying the specific scope location
            For paragraph: {"paragraph_index": 0}
            For table_cell: {"table_index": 0, "row_index": 1, "col_index": 2}

    Returns:
        Success message with replacement count or error description

    Example:
        search_and_replace_in_scope("report.docx", "old text", "new text",
                                   "paragraph", {"paragraph_index": 0})
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if not find_text:
        return "Find text cannot be empty."

    if scope_type not in ["paragraph", "table_cell"]:
        return f"Invalid scope_type '{scope_type}'. Must be either 'paragraph' or 'table_cell'."

    if error := _validate_scope_identifier(scope_type, scope_identifier):
        return error

    if error := _check_file_writable(filename):
        return error

    try:
        editor = FormattedEditor(filename)

        # Create scope location object
        if scope_type == "paragraph":
            scope = ScopeLocation(
                scope_type=scope_type,
                paragraph_index=scope_identifier.get("paragraph_index"),
            )
        else:  # table_cell
            scope = ScopeLocation(
                scope_type=scope_type,
                table_index=scope_identifier.get("table_index"),
                row_index=scope_identifier.get("row_index"),
                col_index=scope_identifier.get("col_index"),
            )

        result = editor.search_and_replace_in_scope(find_text, replace_text, scope)
        if "error" in result:
            return f"Failed to perform replacement: {result['error']}"
        return result["message"]
    except Exception as e:
        return f"Unable to perform search and replace: {str(e)}"


async def is_element_empty(
    filename: str, element_type: str, element_identifier: dict
) -> str:
    """Check if a specific element (paragraph or table cell) is empty in a Word document.

    Args:
        filename: Path to the Word document
        element_type: Type of element ("paragraph" or "table_cell")
        element_identifier: Dictionary identifying the specific element location
            For paragraph: {"paragraph_index": 0}
            For table_cell: {"table_index": 0, "row_index": 1, "col_index": 2}

    Returns:
        JSON string containing emptiness status and element details

    Example:
        is_element_empty("report.docx", "paragraph", {"paragraph_index": 0})
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return error

    if element_type not in ["paragraph", "table_cell"]:
        return f"Invalid element_type '{element_type}'. Must be either 'paragraph' or 'table_cell'."

    if error := _validate_scope_identifier(element_type, element_identifier):
        return error

    try:
        # Use DocumentAnalyzer to check if element is empty
        analyzer = DocumentAnalyzer(filename)

        if element_type == "paragraph":
            paragraphs = analyzer.get_paragraphs_analysis()
            para_index = element_identifier["paragraph_index"]

            if not (0 <= para_index < len(paragraphs)):
                return f"Invalid paragraph index: {para_index}. Document has {len(paragraphs)} paragraphs."

            paragraph = paragraphs[para_index]
            is_empty = not paragraph["text"].strip()

            result = {
                "element_type": element_type,
                "element_identifier": element_identifier,
                "is_empty": is_empty,
                "text_length": len(paragraph["text"]),
                "has_content": bool(paragraph["text"].strip()),
            }

        else:  # table_cell
            tables = analyzer.get_tables_analysis()
            table_index = element_identifier["table_index"]
            row_index = element_identifier["row_index"]
            col_index = element_identifier["col_index"]

            if not (0 <= table_index < len(tables)):
                return f"Invalid table index: {table_index}. Document has {len(tables)} tables."

            table = tables[table_index]

            if not (0 <= row_index < len(table["cells"])):
                return f"Invalid row index: {row_index}. Table has {len(table['cells'])} rows."

            if not (0 <= col_index < len(table["cells"][row_index])):
                return f"Invalid column index: {col_index}. Row has {len(table['cells'][row_index])} columns."

            cell = table["cells"][row_index][col_index]
            is_empty = not cell["text"].strip()

            result = {
                "element_type": element_type,
                "element_identifier": element_identifier,
                "is_empty": is_empty,
                "text_length": len(cell["text"]),
                "has_content": bool(cell["text"].strip()),
                "paragraph_count": len(cell["paragraphs"]),
            }

        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Unable to check element status: {str(e)}"
