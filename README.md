# DocFlow
A Python-based document processing GUI for text replacement, PDF merging, and font editing.

📌 Project Overview

This project is a GUI-based document processing tool that automates various tasks such as:

📄 Word Processing: Replace specific text in Word documents and update filenames accordingly.

📑 PDF Combine: Merge Word-generated PDFs with invoice PDFs based on filename matching.

🎨 Change Font: Change the font of all text in Word documents to Calibri 12pt.

The tool is built using Python, PySimpleGUI, python-docx, docx2pdf, and PyPDF2, making it user-friendly and efficient for bulk document processing.

1️⃣ Intent of the Project

I created this project to simplify and automate repetitive document processing tasks, reducing the time and effort required for users who need to edit and manage multiple documents at once.

The main goals were:

Provide an easy-to-use Graphical User Interface (GUI) for document processing.

Automate text replacements, PDF merging, and font standardization.

Ensure smooth and efficient performance for large batches of files.

2️⃣ Challenges Faced

🔹 Git & GitHub Challenges

One of the biggest challenges I faced was not integrating Git and GitHub from the beginning of the project. This caused difficulties in version control, tracking changes, and syncing my work across devices. I had to learn Git commands, resolve merge conflicts, and properly push multiple versions of my project while keeping my repository clean and structured.

During development, I encountered several technical challenges:

🔹 GUI Issues

The text replacement table was not updating when new entries were added.

Progress bars were stuck at 0%, not reflecting file processing progress.

🔹 Performance Bottlenecks

Processing large batches of files was slow due to sequential execution.

High memory consumption when handling large DOCX and PDF files.

🔹 File Handling Problems

Filename replacements were inconsistent, leading to duplicate names.

Temporary PDF files were not being deleted, causing storage issues.

🔹 Dependency Management

Users faced crashes if required dependencies were missing.

3️⃣ How I Fixed the Issues

✅ GUI Fixes

Fixed the "Add" button by properly updating the replacement text table:

replacement_data.append([old_text, new_text])
window["-TABLE-"].update(values=replacement_data)

Enabled progress bar updates by running file processing in a background thread:

def update_progress(progress_key, value):
    progress_queue.put((progress_key, value))

✅ Performance Improvements

Implemented multiprocessing to speed up Word and PDF processing:

with ProcessPoolExecutor() as executor:
    futures = {executor.submit(process_word_file, file, replacements, destination_folder): file for file in files}

Used in-memory file handling (`BytesIO`) to reduce disk I/O.


✅ File Handling Fixes

Ensured filename replacements apply correctly to both content and filenames:

new_filename = file_path.name
for old_text, new_text in replacements:
    new_filename = new_filename.replace(old_text, new_text)

Ensured temporary PDF files are deleted after merging:

try:
    convert(str(docx_file), str(temp_pdf_path))
finally:
    if temp_pdf_path.exists():
        temp_pdf_path.unlink()

✅ Dependency Management

Automatically installed missing libraries before running the script:

REQUIRED_LIBS = ["PySimpleGUI", "python-docx", "docx2pdf", "PyPDF2"]
def check_and_install_dependencies():
    for lib in REQUIRED_LIBS:
        try:
            __import__(lib)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

4️⃣ Lessons Learned

This project reinforced several key development lessons:

🔸 GUI & User Experience

User feedback is critical for refining functionality.

Thread-safe UI updates improve responsiveness in long-running processes.

🔸 Performance Optimization

Parallel processing significantly improves execution time.

Efficient file handling prevents high memory usage and slow processing.

🔸 Error Handling & Debugging

Logging every action helps users identify issues in file processing.

Validating user inputs prevents common runtime errors.

🔸 Maintainability & Scalability

🔸 Git & GitHub Workflow

I learned the importance of integrating Git and GitHub from the beginning of the project.

Setting up version control early helps with tracking changes, managing multiple versions, and avoiding conflicts.

Resolving Git merge issues, using branches, and pushing structured commits provided valuable hands-on experience.

Modular functions made debugging and future updates much easier.

Dependency management automation prevents installation issues for users.

🚀 Conclusion

This DocFlow successfully automates common document tasks with an intuitive interface. The optimizations I made ensure it runs efficiently without freezing or crashing. The project is now ready for deployment on GitHub.

📌 Next Steps

✅ Upload the final version of the script ( ocFlow.py) to GitHub.✅ Include this project history in the README.md file.✅ Optionally, add screenshots or a video demo of the application in action.

🛠 Installation Instructions

1️⃣ Install Git and Python (if not already installed)

Download and install Git: Git Official Website

Download and install Python: Python Official Website

2️⃣ Clone the repository

git clone https://github.com/RafaelDataSci/DocFlow.git
cd DocFlow

3️⃣ Install dependencies

pip install -r requirements.txt

4️⃣ Run the application

python DocFlow.py

1️⃣ Clone the repository

git clone https://github.com/your-username/DocFlow.git
cd DocFlow

D

2️⃣ Install dependencies

pip install -r requirements.txt

3️⃣ Run the application

python DocDocFlow






