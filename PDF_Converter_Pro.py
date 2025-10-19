import sys, os, time
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QListWidget, QListWidgetItem, QMessageBox,
    QComboBox, QFrame, QProgressBar
)
from PyQt6.QtCore import Qt, QPropertyAnimation, QRect, QThread, pyqtSignal
from docx2pdf import convert as docx_to_pdf
from pdf2docx import Converter
from pptx import Presentation
import openpyxl


# ---------------------------
# Worker Thread for Conversion
# ---------------------------
class ConverterThread(QThread):
    progress = pyqtSignal(int, str)   # (percentage, filename)
    done = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, files, mode, output_folder):
        super().__init__()
        self.files = files
        self.mode = mode
        self.output_folder = output_folder

    def run(self):
        try:
            total = len(self.files)
            for i, file in enumerate(self.files):
                filename = os.path.basename(file)
                self.progress.emit(int((i / total) * 100), f"Starting {filename}...")
                time.sleep(0.3)  # Just for visualization
                self.convert_file(file)
                percent = int(((i + 1) / total) * 100)
                self.progress.emit(percent, f"Converted: {filename}")
            self.done.emit()
        except Exception as e:
            self.error.emit(str(e))

    def convert_file(self, file):
        mode = self.mode
        filename = os.path.basename(file)
        name, ext = os.path.splitext(filename)
        output_path = os.path.join(self.output_folder, f"{name}_converted")

        if mode == "Word → PDF":
            docx_to_pdf(file, self.output_folder)

        elif mode == "PDF → Word":
            cv = Converter(file)
            cv.convert(f"{output_path}.docx")
            cv.close()

        elif mode == "PowerPoint → PDF":
            prs = Presentation(file)
            prs.save(f"{output_path}.pdf")

        elif mode == "Excel → PDF":
            wb = openpyxl.load_workbook(file)
            wb.save(f"{output_path}.pdf")


# ---------------------------
# Main Window
# ---------------------------
class OfficePDFConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Office ⇄ PDF Converter Pro")
        self.setGeometry(300, 200, 900, 600)
        self.files = []
        self.output_folder = ""
        self.init_ui()

    def init_ui(self):
        # ---- Modern Dark Glass Theme ----
        self.setStyleSheet("""
            QWidget {
                background-color: #0d1117;
                color: #e6edf3;
                font-family: 'Segoe UI';
            }
            QPushButton {
                background-color: #238636;
                color: white;
                font-size: 15px;
                border-radius: 8px;
                padding: 10px 18px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2ea043;
            }
            QComboBox {
                background-color: #161b22;
                color: #c9d1d9;
                border: 1px solid #30363d;
                border-radius: 8px;
                padding: 8px 12px;
            }
            QListWidget {
                background-color: #161b22;
                border-radius: 12px;
                padding: 10px;
                border: 1px solid #30363d;
            }
            QFrame {
                background-color: #1e2329;
                border: 1px solid #2f353d;
                border-radius: 12px;
            }
            QLabel {
                font-size: 14px;
            }
            QProgressBar {
                border: 1px solid #2f353d;
                border-radius: 8px;
                text-align: center;
                background-color: #161b22;
                color: #fff;
            }
            QProgressBar::chunk {
                background-color: #1f6feb;
                border-radius: 8px;
            }
        """)

        layout = QVBoxLayout()

        # Header
        header = QLabel("📂 Office ⇄ PDF Converter Pro")
        header.setStyleSheet("font-size: 24px; font-weight: bold; color: #58a6ff;")
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Mode combo
        self.combo_mode = QComboBox()
        self.combo_mode.addItems([
            "Word → PDF", "PDF → Word",
            "PowerPoint → PDF", "Excel → PDF"
        ])
        self.combo_mode.currentIndexChanged.connect(self.clear_files)

        # File list
        self.list_files = QListWidget()
        self.list_files.setSpacing(12)

        # Buttons layout
        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("➕ Add Files")
        self.btn_add.clicked.connect(self.add_files)

        self.btn_convert = QPushButton("⚡ Convert Now")
        self.btn_convert.clicked.connect(self.start_conversion)

        self.btn_clear_all = QPushButton("🗑 Clear All")
        self.btn_clear_all.setStyleSheet("""
            QPushButton {
                background-color: #da3633;
                color: white;
                border-radius: 8px;
                padding: 10px 18px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #f85149;
            }
        """)
        self.btn_clear_all.clicked.connect(self.clear_files)

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_convert)
        btn_layout.addWidget(self.btn_clear_all)

        # Progress Bar and Label
        self.progress_label = QLabel("")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        self.progress_label.hide()

        layout.addWidget(header)
        layout.addWidget(self.combo_mode)
        layout.addWidget(self.list_files)
        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    # ------------------------
    # File Handling
    # ------------------------
    def get_allowed_extensions(self):
        mode = self.combo_mode.currentText()
        if "Word → PDF" in mode:
            return ["docx"]
        elif "PDF → Word" in mode:
            return ["pdf"]
        elif "PowerPoint" in mode:
            return ["pptx"]
        elif "Excel" in mode:
            return ["xlsx"]
        return []

    def add_files(self):
        allowed_ext = self.get_allowed_extensions()
        ext_str = " ".join([f"*.{ext}" for ext in allowed_ext])
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", f"Files ({ext_str})")

        for file in files:
            ext = file.split(".")[-1].lower()
            if ext not in allowed_ext:
                QMessageBox.warning(self, "Invalid File", f"Only {', '.join(allowed_ext)} files are allowed in this mode.")
                continue
            if file not in self.files:
                self.files.append(file)
                self.add_file_card(file)

    def add_file_card(self, filepath):
        filename = os.path.basename(filepath)
        item = QListWidgetItem()
        frame = QFrame()
        hbox = QHBoxLayout()

        icon = QLabel("📄")
        icon.setStyleSheet("font-size: 20px; margin-right: 8px;")

        label = QLabel(filename)
        label.setStyleSheet("font-size: 15px; font-weight: bold; color: #e6edf3;")

        btn_remove = QPushButton("🗙")
        btn_remove.setFixedSize(32, 32)
        btn_remove.setStyleSheet("""
            QPushButton {
                background-color: #8b0000;
                border-radius: 8px;
                color: white;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #a40000;
            }
        """)
        btn_remove.clicked.connect(lambda: self.remove_file(filepath, item))

        hbox.addWidget(icon)
        hbox.addWidget(label)
        hbox.addStretch()
        hbox.addWidget(btn_remove)
        frame.setLayout(hbox)
        item.setSizeHint(frame.sizeHint())
        self.list_files.addItem(item)
        self.list_files.setItemWidget(item, frame)

        # Animation
        anim = QPropertyAnimation(frame, b"geometry")
        anim.setDuration(300)
        anim.setStartValue(QRect(0, 0, 0, 0))
        anim.setEndValue(QRect(0, 0, 500, 50))
        anim.start()

    def remove_file(self, filepath, item):
        if filepath in self.files:
            self.files.remove(filepath)
        row = self.list_files.row(item)
        self.list_files.takeItem(row)

    def clear_files(self):
        self.files = []
        self.list_files.clear()

    # ------------------------
    # Conversion Handling
    # ------------------------
    def start_conversion(self):
        if not self.files:
            QMessageBox.warning(self, "No Files", "Please add at least one file to convert.")
            return

        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not folder:
            return

        self.output_folder = folder
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.progress_label.show()
        self.progress_label.setText("Preparing conversion...")

        self.thread = ConverterThread(self.files, self.combo_mode.currentText(), folder)
        self.thread.progress.connect(self.update_progress)
        self.thread.done.connect(self.conversion_done)
        self.thread.error.connect(self.conversion_error)
        self.thread.start()

    def update_progress(self, value, filename):
        self.progress_bar.setValue(value)
        self.progress_label.setText(f"🔄 {filename}  ({value}%)")

    def conversion_done(self):
        self.progress_bar.hide()
        self.progress_label.hide()
        QMessageBox.information(self, "✅ Success", "All files have been converted successfully!")

    def conversion_error(self, msg):
        self.progress_bar.hide()
        self.progress_label.hide()
        QMessageBox.critical(self, "❌ Error", msg)


# ---------------------------
# Run App
# ---------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OfficePDFConverter()
    window.show()
    sys.exit(app.exec())
