# 📄 DocFlow

*A Python-based document processing GUI for text replacement, PDF merging, and font editing.*

---

## 📌 Project Overview
This project is a **GUI-based document processing tool** that automates various tasks, making document handling more efficient and user-friendly. 

### **Features:**
✅ **📄 Word Processing** – Replace text in Word documents and update filenames accordingly.
✅ **📑 PDF Combine** – Merge Word-generated PDFs with invoice PDFs based on filename matching.
✅ **🎨 Change Font** – Modify all text in Word documents to **Calibri 12pt** for consistency.

This tool is built using **Python**, **PySimpleGUI**, **python-docx**, **docx2pdf**, and **PyPDF2**, offering an intuitive and accessible experience for non-technical users handling bulk document processing.

---

## 1️⃣ Intent of the Project
I created this project to **automate repetitive document processing tasks**, reducing time and effort for users who need to edit and manage multiple files simultaneously.

### **Goals:**
- 🖥️ Provide an **easy-to-use** **Graphical User Interface (GUI)** for document processing.
- 🔄 Automate **text replacements, PDF merging, and font standardization**.
- ⚡ Ensure **high performance** even for large batches of files.

---

## 2️⃣ Challenges Faced

### 🔹 **Git & GitHub Challenges**
One of the biggest challenges was **not integrating Git and GitHub from the beginning** of the project. This led to difficulties in:
- 📌 Version control and tracking changes.
- 🔄 Syncing work across devices.
- 💻 Learning Git commands, resolving merge conflicts, and properly structuring commits.

### 🔹 **GUI Issues**
- The **text replacement table was not updating** when new entries were added.
- **Progress bars were stuck at 0%**, not reflecting file processing progress.

### 🔹 **Performance Bottlenecks**
- **Slow processing** for large document batches due to sequential execution.
- **High memory usage** when handling DOCX and PDF files.

### 🔹 **File Handling Problems**
- **Inconsistent filename replacements** causing duplicate names.
- **Temporary PDF files not being deleted**, leading to storage issues.

### 🔹 **Dependency Management**
- Users faced **crashes** if required dependencies were missing.

---

## 3️⃣ How I Fixed the Issues

### **✅ GUI Fixes**
- **Fixed the "Add" button** to properly update the replacement text table:
  ```python
  replacement_data.append([old_text, new_text])
  window["-TABLE-"].update(values=replacement_data)
  ```
- **Enabled progress bar updates** using a **background thread**:
  ```python
  def update_progress(progress_key, value):
      progress_queue.put((progress_key, value))
  ```

### **✅ Performance Improvements**
- **Implemented multiprocessing** to improve speed:
  ```python
  with ProcessPoolExecutor() as executor:
      futures = {executor.submit(process_word_file, file, replacements, destination_folder): file for file in files}
  ```
- **Used in-memory file handling (`BytesIO`)** to reduce disk I/O.

### **✅ File Handling Fixes**
- **Ensured filename replacements apply correctly**:
  ```python
  new_filename = file_path.name
  for old_text, new_text in replacements:
      new_filename = new_filename.replace(old_text, new_text)
  ```
- **Ensured temporary PDF files are deleted after merging**:
  ```python
  try:
      convert(str(docx_file), str(temp_pdf_path))
  finally:
      if temp_pdf_path.exists():
          temp_pdf_path.unlink()
  ```

### **✅ Dependency Management**
- **Automatically installed missing libraries** at runtime:
  ```python
  REQUIRED_LIBS = ["PySimpleGUI", "python-docx", "docx2pdf", "PyPDF2"]
  def check_and_install_dependencies():
      for lib in REQUIRED_LIBS:
          try:
              __import__(lib)
          except ImportError:
              subprocess.check_call([sys.executable, "-m", "pip", "install", lib])
  ```

---

## 4️⃣ Lessons Learned

### 🔹 **GUI & User Experience**
- **User feedback is critical** for refining functionality.
- **Thread-safe UI updates** improve responsiveness in long-running processes.

### 🔹 **Performance Optimization**
- **Parallel processing** significantly improves execution time.
- **Efficient file handling** prevents high memory usage and slow processing.

### 🔹 **Error Handling & Debugging**
- **Logging every action** helps identify issues in file processing.
- **Validating user inputs** prevents common runtime errors.

### 🔹 **Git & GitHub Workflow**
- I learned the importance of **integrating Git and GitHub from the start**.
- Setting up **version control early** makes tracking changes and managing versions easier.
- **Resolving Git merge conflicts, structuring commits, and using branches** provided valuable experience.
- **Automating dependency management** prevents future installation issues.

---

## 🚀 Conclusion
This **DocFlow** project successfully automates document processing while ensuring high performance and usability. The **optimizations** I implemented prevent freezing, improve speed, and simplify user interaction. Now, the project is **ready for deployment on GitHub**.

---

## 🛠 Installation Instructions

### 1️⃣ **Install Git and Python** (if not already installed)
- **Download Git:** [Git Official Website](https://git-scm.com/downloads)
- **Download Python:** [Python Official Website](https://www.python.org/downloads/)

### 2️⃣ **Clone the repository**
```bash
git clone https://github.com/RafaelDataSci/DocFlow.git
cd DocFlow
```

### 3️⃣ **Install dependencies**
```bash
pip install -r requirements.txt
```

### 4️⃣ **Run the application**
```bash
python DocFlow.py
```

![image](https://github.com/user-attachments/assets/dfba0be2-d439-443e-83bf-bf60f0949c99)


