# 🧩 Office ⇄ PDF Converter Pro

A modern **desktop application (PyQt6)** that converts between Microsoft Office files and PDFs — including **Word, PowerPoint, Excel, and PDF** formats.  
It features a **beautiful dark UI**, smooth animations, and live conversion progress tracking.

---

## 🚀 Features

- 🔁 **Convert both ways**:  
  - Word → PDF  
  - PDF → Word  
  - PowerPoint → PDF  
  - Excel → PDF  
- 🎨 **Dark Modern Interface** with smooth animations  
- 🧩 **Multi-file support** (Add, remove, or clear all files easily)  
- ⚡ **Progress tracking** with live percentage updates  
- 💬 **Error handling** for invalid files or failed conversions  
- 🧵 **Multithreaded conversion** (no freezing UI)

---

## 🧠 Tech Stack

| Library | Description |
|----------|-------------|
| **Python 3.10+** | Core language |
| **PyQt6** | GUI framework |
| **docx2pdf** | Convert Word → PDF |
| **pdf2docx** | Convert PDF → Word |
| **python-pptx** | Handle PowerPoint files |
| **openpyxl** | Handle Excel files |

---

## ⚙️ Installation

### 1️⃣ Clone this repository
```bash 
git clone https://github.com/yourusername/OfficePDFConverterPro.git cd OfficePDFConverterPro
```
### 2️⃣ Create a virtual environment (optional but recommended)
```bash
python -m venv venv
venv\Scripts\activate      # On Windows (PowerShell/CMD)
source venv/bin/activate   # On macOS/Linux
```
### 3️⃣ Install dependencies
```bash
pip install PyQt6 docx2pdf pdf2docx python-pptx openpyxl
```
### ▶️ Run the App
```bash
python OfficePDFConverterPro.py
```

## 🧭 Usage Guide

1. Choose the conversion mode from the dropdown menu  
2. Click **“➕ Add Files”** to upload supported files  
3. Click **“⚡ Convert Now”** to start conversion  
4. Select the output folder when prompted  
5. Watch the progress bar and messages as files convert

---

## 🧹 UI Controls

| Button | Description |
| --- | --- |
| ➕ **Add Files** | Add files according to selected mode |
| ⚡ **Convert Now** | Start converting all listed files |
| 🗑 **Clear All** | Remove all files from the list |
| 🗙 **Delete File** | Remove an individual file |

---

## 🔐 Notes

- **Word → PDF** requires **Microsoft Word** (Windows) or **LibreOffice** installed.  
- Supported formats: `.docx`, `.pdf`, `.pptx`, `.xlsx`  
- Conversion time depends on file size and system performance.  

---

## 🧩 Future Enhancements

- Add drag & drop support  
- Add PDF → PowerPoint / Excel support  
- Add light/dark theme toggle  
- Add batch rename & file merge  

---

## 👨‍💻 Author

**Mina Robir**  
Software Engineer 
🌐 GitHub: [minarob23](https://github.com/minarob23)
