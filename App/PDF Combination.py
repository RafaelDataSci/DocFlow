#!/usr/bin/env python
"""
PDF Combine Script

For each January 2025 DOCX file in the specified Word origin folder:
  1. Convert the DOCX to a temporary PDF using docx2pdf.
  2. Extract a key from the DOCX file name (the part before " - ").
  3. Find a matching PDF invoice in the PDF origin folder whose name contains that key.
  4. Merge the converted PDF with the matching invoice PDF using PyPDF2.
  5. Save the merged PDF in the destination folder using the matching PDF's original file name.
  
Progress and error messages are logged so you can follow along.
"""

import logging
from pathlib import Path
import tempfile
from docx2pdf import convert
from PyPDF2 import PdfMerger

# Configure logging to show timestamps and log levels.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def combine_word_and_pdf(word_origin: Path, pdf_origin: Path, destination_folder: Path) -> None:
    """
    Process all January 2025 DOCX files in word_origin:
      - Convert each DOCX to a temporary PDF.
      - Extract a key from the DOCX filename for matching.
      - Look up a matching PDF invoice in pdf_origin.
      - Merge the two PDFs.
      - Save the merged PDF in destination_folder using the matching PDF's file name.
    """
    destination_folder.mkdir(parents=True, exist_ok=True)
    
    # Build a lookup dictionary for PDF files (keys are lower-cased file names)
    pdf_lookup = {pdf_file.name.lower(): pdf_file for pdf_file in pdf_origin.glob("*.pdf")}
    
    # Get all January 2025 DOCX files from the Word origin folder
    docx_files = sorted([f for f in word_origin.glob("*.docx") if "January 2025" in f.name])
    total_files = len(docx_files)
    if total_files == 0:
        logging.info("No January 2025 DOCX files found in the specified Word origin folder.")
        return

    logging.info(f"Found {total_files} January 2025 DOCX files.")
    files_processed = 0

    for idx, docx_file in enumerate(docx_files, start=1):
        logging.info(f"Processing file {idx}/{total_files}: {docx_file.name}")
        base_name = docx_file.stem  # File name without extension

        # Extract a key for matching. Assume the key is the part before " - "
        key = base_name.split(" - ")[0].strip().lower()
        if not key:
            key = base_name.lower()

        # Look for a matching PDF that contains the key in its filename
        matching_pdf = None
        for pdf_name, pdf_path in pdf_lookup.items():
            if key in pdf_name:
                matching_pdf = pdf_path
                break

        if not matching_pdf:
            logging.warning(f"No matching PDF found for {docx_file.name}. Skipping.")
            continue

        logging.info(f"Found matching PDF: {matching_pdf.name} for {docx_file.name}.")

        # Convert the DOCX file to a temporary PDF
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_file:
                temp_pdf_path = Path(tmp_file.name)
            convert(str(docx_file), str(temp_pdf_path))
            logging.info(f"Converted {docx_file.name} to temporary PDF: {temp_pdf_path.name}")
        except Exception as e:
            logging.error(f"Error converting {docx_file.name}: {e}")
            if 'temp_pdf_path' in locals() and temp_pdf_path.exists():
                temp_pdf_path.unlink()
            continue

        # Merge the temporary PDF and the matching invoice PDF
        try:
            merger = PdfMerger()
            merger.append(str(temp_pdf_path))
            merger.append(str(matching_pdf))
            
            # Use the matching PDF's original file name for the final merged file.
            output_pdf_path = destination_folder / matching_pdf.name
            merger.write(str(output_pdf_path))
            merger.close()
            logging.info(f"Merged PDF saved as: {output_pdf_path.name}")
            files_processed += 1
        except Exception as e:
            logging.error(f"Error merging PDF for {docx_file.name}: {e}")
        finally:
            # Clean up the temporary file
            try:
                if temp_pdf_path.exists():
                    temp_pdf_path.unlink()
                    logging.info(f"Temporary file {temp_pdf_path.name} deleted.")
            except Exception as e:
                logging.error(f"Error deleting temporary file {temp_pdf_path.name}: {e}")
    
    logging.info(f"Processing complete: {files_processed}/{total_files} files processed successfully.")

if __name__ == "__main__":
    # Prompt the user for the folder paths
    word_origin_input = input("Enter the ORIGIN folder path containing January 2025 DOCX files: ").strip()
    pdf_origin_input = input("Enter the ORIGIN folder path containing PDF invoice files: ").strip()
    destination_input = input("Enter the DESTINATION folder path for merged PDFs: ").strip()
    
    word_origin = Path(word_origin_input)
    pdf_origin = Path(pdf_origin_input)
    destination_folder = Path(destination_input)
    
    # Validate that the origin folders exist
    if not word_origin.is_dir():
        logging.error(f"Word origin folder does not exist: {word_origin}")
    elif not pdf_origin.is_dir():
        logging.error(f"PDF origin folder does not exist: {pdf_origin}")
    else:
        combine_word_and_pdf(word_origin, pdf_origin, destination_folder)
