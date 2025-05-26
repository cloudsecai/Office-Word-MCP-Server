"""
Document imaging tools for Word Document Server.

This module provides tools for generating visual representations of document pages,
particularly useful when textual analysis is insufficient for complex layouts.
"""

import os
import json
import tempfile
from typing import List, Dict, Any, Optional

from word_document_server.utils.file_utils import ensure_docx_extension
from word_document_server.utils.conversion_utils import convert_docx_to_pdf_temp, cleanup_temp_file


def _validate_file_exists(filename: str) -> Optional[str]:
    """Validate that a file exists and return user-friendly error if not."""
    if not os.path.exists(filename):
        return f"The document '{filename}' could not be found. Please check the file path and try again."
    return None


def _validate_page_numbers(page_numbers: List[int]) -> Optional[str]:
    """Validate page numbers list and return user-friendly error if invalid."""
    if not page_numbers:
        return "Page numbers list cannot be empty. Please provide at least one page number."
    
    if not isinstance(page_numbers, list):
        return "Page numbers must be provided as a list of integers."
    
    if len(page_numbers) > 10:
        return f"Too many pages requested. Maximum 10 pages allowed per request (you requested {len(page_numbers)} pages)."
    
    for i, page_num in enumerate(page_numbers):
        if not isinstance(page_num, int):
            return f"All page numbers must be integers. Item at index {i} is not an integer: {page_num}"
        if page_num < 1:
            return f"Page numbers must be 1 or greater (you provided {page_num} at index {i}). Page numbering starts from 1."
    
    return None


def _validate_image_format(image_format: str) -> Optional[str]:
    """Validate image format and return user-friendly error if invalid."""
    supported_formats = ["png", "jpeg", "jpg", "tiff", "bmp"]
    if image_format.lower() not in supported_formats:
        return f"Unsupported image format '{image_format}'. Supported formats: {', '.join(supported_formats)}"
    return None


def _validate_dpi(dpi: int) -> Optional[str]:
    """Validate DPI value and return user-friendly error if invalid."""
    if not isinstance(dpi, int):
        return f"DPI must be an integer (you provided {type(dpi).__name__}: {dpi})"
    if dpi < 50 or dpi > 600:
        return f"DPI must be between 50 and 600 for reasonable performance and quality (you provided {dpi})"
    return None


async def get_document_page_images(
    filename: str,
    page_numbers: List[int],
    output_directory: str = "mcp_server_temp_images",
    image_format: str = "png",
    dpi: int = 200
) -> str:
    """Generate images of specific pages from a Word document (.docx only).

    Converts specified pages from a DOCX document into image files for visual analysis.
    This tool is particularly useful when textual representation is insufficient for
    understanding complex layouts like tables, charts, or formatted content.

    Use this tool when:
    - Textual analysis of tables or layouts is ambiguous
    - Need visual context for accurate information extraction
    - Complex formatting requires visual verification
    - Multi-column layouts or embedded objects need visual inspection

    Args:
        filename: Path to the Word document (.docx format only)
        page_numbers: List of page numbers to convert (1-indexed, e.g., [1, 3, 5], max 10 pages per request)
        output_directory: Server-side directory for generated images (default: "mcp_server_temp_images")
        image_format: Image format for output files (default: "png", supports: png, jpeg, jpg, tiff, bmp)
        dpi: Dots per inch for image quality (default: 200, range: 50-600)

    Returns:
        JSON string containing:
        - success flag and image paths on success
        - error message on failure
        - mapping of page numbers to their server-side file paths

    Process:
        1. Validates input parameters and file existence
        2. Converts DOCX to temporary PDF using platform-appropriate tools
        3. Extracts specified pages as images using pdf2image library
        4. Saves images to the specified output directory
        5. Cleans up temporary PDF file
        6. Returns paths to generated image files

    Requirements:
        - pdf2image library (automatically installed)
        - poppler-utils system package for PDF processing
        - Platform-specific conversion tools (LibreOffice/Microsoft Word)

    Output format:
        Success: {"success": true, "image_paths": {"page_1": "/path/to/image1.png", ...}, "message": "..."}
        Failure: {"error": "Description of the error"}

    Limitations:
        - Only works with .docx format (Microsoft Word 2007+)
        - Cannot process password-protected documents
        - Requires poppler-utils to be installed on the system
        - Page numbers must exist in the document
        - Complex animations or interactive elements won't be captured

    Example:
        get_document_page_images("report.docx", [1, 3], "temp_images", "png", 300)
        # Generates PNG images of pages 1 and 3 at 300 DPI
    """
    filename = ensure_docx_extension(filename)

    # Validate inputs
    if error := _validate_file_exists(filename):
        return json.dumps({"error": error})

    if error := _validate_page_numbers(page_numbers):
        return json.dumps({"error": error})

    if error := _validate_image_format(image_format):
        return json.dumps({"error": error})

    if error := _validate_dpi(dpi):
        return json.dumps({"error": error})

    # Create output directory if it doesn't exist
    try:
        os.makedirs(output_directory, exist_ok=True)
    except Exception as e:
        return json.dumps({"error": f"Cannot create output directory '{output_directory}': {str(e)}"})

    # Convert DOCX to temporary PDF
    success, result = convert_docx_to_pdf_temp(filename)
    if not success:
        return json.dumps({"error": f"Failed to convert DOCX to PDF: {result}"})

    temp_pdf_path = result
    generated_images = {}
    errors = []

    try:
        # Import pdf2image for PDF to image conversion
        try:
            from pdf2image import convert_from_path
        except ImportError:
            return json.dumps({
                "error": "pdf2image library is not available. Please install it using: pip install pdf2image\n"
                         "Note: This also requires poppler-utils to be installed on your system."
            })

        # Get base filename for image naming
        base_filename = os.path.splitext(os.path.basename(filename))[0]

        # Convert each requested page to image
        for page_num in page_numbers:
            try:
                # Convert single page from PDF (pdf2image uses 1-indexed pages)
                images = convert_from_path(
                    temp_pdf_path,
                    dpi=dpi,
                    first_page=page_num,
                    last_page=page_num,
                    fmt=image_format.lower()
                )

                if not images:
                    errors.append(f"Page {page_num} could not be converted to image (page may not exist)")
                    continue

                # Save the image
                image = images[0]  # Should only be one image since we specified a single page
                image_filename = f"{base_filename}_page_{page_num}.{image_format.lower()}"
                image_path = os.path.join(output_directory, image_filename)

                # Save image to file
                image.save(image_path, image_format.upper())
                
                # Store the absolute path for the response
                generated_images[f"page_{page_num}"] = os.path.abspath(image_path)

            except Exception as e:
                errors.append(f"Failed to generate image for page {page_num}: {str(e)}")

    except Exception as e:
        # Clean up temporary PDF
        cleanup_temp_file(temp_pdf_path)
        return json.dumps({"error": f"Image generation failed: {str(e)}"})

    finally:
        # Always clean up the temporary PDF file
        cleanup_temp_file(temp_pdf_path)

    # Prepare response
    if generated_images:
        response = {
            "success": True,
            "image_paths": generated_images,
            "message": f"Successfully generated {len(generated_images)} image(s) for the specified page(s)."
        }
        
        if errors:
            response["warnings"] = errors
            response["message"] += f" {len(errors)} page(s) had errors."
    else:
        error_summary = "; ".join(errors) if errors else "No images were generated"
        response = {"error": f"Failed to generate any images: {error_summary}"}

    return json.dumps(response, indent=2) 