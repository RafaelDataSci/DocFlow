#!/usr/bin/env python
import PySimpleGUI as sg
import threading
import logging
import time
from pathlib import Path
from docx import Document
import tempfile
from shutil import copyfile
from docx2pdf import convert
from PyPDF2 import PdfMerger
from queue import Queue

# ---------------------------
# GUI Setup
# ---------------------------
sg.theme('DarkBlue14')

# Layout for Word Processing
tab1_layout = [
    [sg.Text("Word Document Processing", font=("Any", 16))],
    [sg.Text("Origin Folder:"), sg.Input(key="-WP_ORIGIN-"), sg.FolderBrowse()],
    [sg.Text("Destination Folder:"), sg.Input(key="-WP_DEST-"), sg.FolderBrowse()],
    
    [sg.Text("Replacement Pairs:")],
    [sg.Table(values=[], headings=["Old Text", "New Text"], 
              key="-TABLE-", enable_events=True, auto_size_columns=True, 
              justification='left', num_rows=5)],
    
    [sg.Text("Old Text:"), sg.Input(size=(25,1), key="-OLD_TEXT-"), 
     sg.Text("New Text:"), sg.Input(size=(25,1), key="-NEW_TEXT-")],
    
    [sg.Button("Add", key="-ADD-"), sg.Button("Remove", key="-REMOVE-"), sg.Button("Clear All", key="-CLEAR-")],
    [sg.Button("Run Word Processing", key="-RUN_WP-")],
    [sg.ProgressBar(100, orientation='h', size=(40, 20), key='-WP_PROGRESS-')]
]

# Layout for PDF Combine
tab2_layout = [
    [sg.Text("PDF Combine", font=("Any", 16))],
    [sg.Text("DOCX (Word) Origin Folder:"), sg.Input(key="-PC_WORD-"), sg.FolderBrowse()],
    [sg.Text("PDF Invoice Origin Folder:"), sg.Input(key="-PC_PDF-"), sg.FolderBrowse()],
    [sg.Text("Destination Folder:"), sg.Input(key="-PC_DEST-"), sg.FolderBrowse()],
    [sg.Button("Run PDF Combine", key="-RUN_PC-")],
    [sg.ProgressBar(100, orientation='h', size=(40, 20), key='-PC_PROGRESS-')]
]

layout = [
    [sg.TabGroup([[sg.Tab("Word Processing", tab1_layout), sg.Tab("PDF Combine", tab2_layout)]])],
    [sg.Text("Output:")],
    [sg.Multiline(size=(100, 20), key="-OUTPUT-", autoscroll=True, disabled=True)]
]

window = sg.Window("Document Processing GUI", layout, finalize=True)

# ---------------------------
# Thread-Safe Logging & Progress Bar
# ---------------------------
log_queue = Queue()
progress_queue = Queue()

def log_message(msg):
    """Put log messages into the queue."""
    log_queue.put(msg)

def update_progress(progress_key, value):
    """Put progress updates into the queue."""
    progress_queue.put((progress_key, value))

# ---------------------------
# File Processing Functions
# ---------------------------
def process_word_file(file_path, replacements, destination_folder):
    """Process a Word file, apply text replacements, and save it with a new name."""
    try:
        # Use a temporary directory to avoid Microsoft Word temp file issues
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_file_path = Path(temp_dir) / file_path.name

            # Copy the original file to temp directory
            copyfile(file_path, temp_file_path)

            doc = Document(temp_file_path)

            # Apply text replacements inside the document
            for paragraph in doc.paragraphs:
                for old_text, new_text in replacements:
                    if old_text in paragraph.text:
                        paragraph.text = paragraph.text.replace(old_text, new_text)

            # Rename file in destination
            new_filename = file_path.name
            for old_text, new_text in replacements:
                new_filename = new_filename.replace(old_text, new_text)

            destination_folder.mkdir(parents=True, exist_ok=True)
            new_file_path = destination_folder / new_filename

            # Save the updated document to destination
            doc.save(new_file_path)

            log_message(f"Processed and renamed: {file_path.name} → {new_file_path.name}")

    except Exception as e:
        log_message(f"Error processing {file_path}: {e}")

def process_directory(origin_folder, destination_folder, replacements):
    """Process all Word files in the directory and rename them correctly."""
    files = list(origin_folder.glob("*.docx"))
    total_files = len(files)
    
    if total_files == 0:
        log_message("No DOCX files found.")
        return

    destination_folder.mkdir(parents=True, exist_ok=True)

    for idx, file in enumerate(files, start=1):
        process_word_file(file, replacements, destination_folder)
        update_progress("-WP_PROGRESS-", (idx / total_files) * 100)
        time.sleep(0.1)  # Simulate processing delay

    log_message(f"Word Processing complete! {total_files} files updated.")

def combine_word_and_pdf(word_origin, pdf_origin, destination_folder):
    """Combine DOCX files converted to PDF with matching invoice PDFs."""
    word_files = list(word_origin.glob("*.docx"))
    pdf_lookup = {pdf.name.lower(): pdf for pdf in pdf_origin.glob("*.pdf")}
    
    total_files = len(word_files)
    
    for idx, docx_file in enumerate(word_files, start=1):
        base_name = docx_file.stem.split(" - ")[0].strip().lower()
        matching_pdf = next((pdf for name, pdf in pdf_lookup.items() if base_name in name), None)

        if not matching_pdf:
            log_message(f"No matching PDF for {docx_file.name}. Skipping.")
            continue

        try:
            temp_pdf = Path(tempfile.mktemp(suffix=".pdf"))
            convert(str(docx_file), str(temp_pdf))
            merger = PdfMerger()
            merger.append(str(temp_pdf))
            merger.append(str(matching_pdf))
            output_pdf_path = destination_folder / matching_pdf.name
            merger.write(str(output_pdf_path))
            merger.close()
            temp_pdf.unlink()
            log_message(f"Merged PDF saved as: {output_pdf_path.name}")
        except Exception as e:
            log_message(f"Error merging {docx_file.name}: {e}")

        update_progress("-PC_PROGRESS-", (idx / total_files) * 100)

# ---------------------------
# Event Loop & Threading
# --------------------------- 
replacement_data = []

while True:
    event, values = window.read(timeout=100)

    # Process logs and progress updates
    while not log_queue.empty():
        msg = log_queue.get_nowait()
        window["-OUTPUT-"].update(msg + "\n", append=True)

    while not progress_queue.empty():
        progress_key, progress_value = progress_queue.get_nowait()
        window[progress_key].update(progress_value)

    if event == sg.WIN_CLOSED:
        break

    elif event == "-ADD-":
        old_text, new_text = values["-OLD_TEXT-"].strip(), values["-NEW_TEXT-"].strip()
        if old_text and new_text:
            replacement_data.append([old_text, new_text])
            window["-TABLE-"].update(values=replacement_data)
            window["-OLD_TEXT-"].update("")
            window["-NEW_TEXT-"].update("")
        else:
            sg.popup_error("Enter both old and new text.")

    elif event == "-RUN_WP-":
        replacements = [(row[0], row[1]) for row in replacement_data]
        thread = threading.Thread(target=process_directory, args=(Path(values["-WP_ORIGIN-"]), Path(values["-WP_DEST-"]), replacements), daemon=True)
        thread.start()

window.close()
