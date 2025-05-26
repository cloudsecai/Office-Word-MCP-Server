"""
Conversion utilities for document format transformations.

This module provides reusable conversion functions for transforming documents
between different formats, particularly for generating temporary files needed
for other operations.
"""

import os
import platform
import subprocess
import shutil
import tempfile
from typing import Optional, Tuple

from word_document_server.utils.file_utils import check_file_writeable


def convert_docx_to_pdf_temp(filename: str, temp_dir: Optional[str] = None) -> Tuple[bool, str]:
    """Convert a DOCX file to a temporary PDF file.
    
    This is a utility function that creates a temporary PDF file from a DOCX document
    using platform-appropriate conversion tools. The temporary PDF should be cleaned
    up by the caller when no longer needed.
    
    Args:
        filename: Path to the source DOCX file
        temp_dir: Optional directory for temporary files (uses system temp if None)
        
    Returns:
        Tuple of (success: bool, result: str)
        - If success=True, result contains the path to the temporary PDF
        - If success=False, result contains the error message
    """
    if not os.path.exists(filename):
        return False, f"Source document '{filename}' does not exist"

    # Create temporary PDF file
    if temp_dir:
        os.makedirs(temp_dir, exist_ok=True)
        temp_fd, temp_pdf_path = tempfile.mkstemp(suffix='.pdf', dir=temp_dir)
    else:
        temp_fd, temp_pdf_path = tempfile.mkstemp(suffix='.pdf')
    
    # Close the file descriptor since we just need the path
    os.close(temp_fd)
    
    try:
        # Check if output file can be written
        is_writeable, error_message = check_file_writeable(temp_pdf_path)
        if not is_writeable:
            os.unlink(temp_pdf_path)  # Clean up the temp file
            return False, f"Cannot create temporary PDF: {error_message}"

        # Determine platform for appropriate conversion method
        system = platform.system()

        if system == "Windows":
            # On Windows, try docx2pdf which uses Microsoft Word
            try:
                from docx2pdf import convert
                convert(filename, temp_pdf_path)
                return True, temp_pdf_path
            except (ImportError, Exception) as e:
                os.unlink(temp_pdf_path)  # Clean up on failure
                return False, f"Failed to convert document to PDF: {str(e)}\nNote: docx2pdf requires Microsoft Word to be installed."

        elif system in ["Linux", "Darwin"]:  # Linux or macOS
            # Use LibreOffice for headless conversion (preferred for server environments)
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
                        # Get the directory for output
                        output_dir = os.path.dirname(temp_pdf_path)
                        
                        # Enhanced command for headless operation
                        cmd = [
                            cmd_name,
                            "--headless",
                            "--invisible",
                            "--nodefault",
                            "--nolockcheck",
                            "--nologo",
                            "--norestore",
                            "--convert-to",
                            "pdf",
                            "--outdir",
                            output_dir,
                            filename,
                        ]

                        # Set environment variables for headless operation
                        env = os.environ.copy()
                        env.update({
                            "DISPLAY": "",  # Ensure no display is used
                            "HOME": os.path.expanduser("~"),  # Ensure HOME is set
                        })

                        result = subprocess.run(
                            cmd, 
                            capture_output=True, 
                            text=True, 
                            timeout=120,  # Increased timeout for larger files
                            env=env
                        )

                        if result.returncode == 0:
                            # LibreOffice creates the PDF with the same basename as the source
                            base_name = os.path.basename(filename)
                            pdf_base_name = os.path.splitext(base_name)[0] + ".pdf"
                            created_pdf = os.path.join(output_dir, pdf_base_name)

                            # If the created PDF is not at the desired location, move it
                            if created_pdf != temp_pdf_path and os.path.exists(created_pdf):
                                shutil.move(created_pdf, temp_pdf_path)

                            conversion_successful = True
                            break  # Exit the loop if successful
                        else:
                            errors.append(f"{cmd_name} error (returncode {result.returncode}): {result.stderr.strip()}")
                    except subprocess.TimeoutExpired:
                        errors.append(f"{cmd_name} error: Conversion timed out after 120 seconds")
                    except (subprocess.SubprocessError, FileNotFoundError) as e:
                        errors.append(f"{cmd_name} error: {str(e)}")

                if conversion_successful:
                    return True, temp_pdf_path
                else:
                    # For headless environments, don't fall back to docx2pdf (which opens GUI apps)
                    os.unlink(temp_pdf_path)  # Clean up on failure
                    error_msg = "Failed to convert document to PDF using LibreOffice (headless mode).\n"
                    error_msg += "LibreOffice errors: " + "; ".join(errors) + "\n"
                    error_msg += "For headless operation, please ensure LibreOffice is properly installed:\n"
                    error_msg += "- Linux: apt-get install libreoffice --no-install-recommends\n"
                    error_msg += "- macOS: brew install --cask libreoffice\n"
                    error_msg += "- Docker: Use an image with LibreOffice pre-installed"
                    return False, error_msg

            except Exception as e:
                os.unlink(temp_pdf_path)  # Clean up on failure
                return False, f"Failed to convert document to PDF: {str(e)}"
        else:
            os.unlink(temp_pdf_path)  # Clean up on failure
            return False, f"PDF conversion not supported on {system} platform"

    except Exception as e:
        # Clean up temporary file on any exception
        if os.path.exists(temp_pdf_path):
            os.unlink(temp_pdf_path)
        return False, f"Failed to convert document to PDF: {str(e)}"


def cleanup_temp_file(file_path: str) -> bool:
    """Safely remove a temporary file.
    
    Args:
        file_path: Path to the temporary file to remove
        
    Returns:
        True if file was successfully removed or didn't exist, False otherwise
    """
    try:
        if os.path.exists(file_path):
            os.unlink(file_path)
        return True
    except Exception:
        return False 