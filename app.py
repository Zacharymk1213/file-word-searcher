import os
import fnmatch
import shutil
import subprocess
import platform
import io
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel, QFileDialog, QTreeWidget, QTreeWidgetItem, QCheckBox, QMessageBox
from PyQt5.QtCore import Qt
from docx import Document
from odf.opendocument import load as load_odt
from odf.text import P
from PyPDF2 import PdfReader
import re




class FileSearcherApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Searcher")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.create_widgets()

    def create_widgets(self):
        # Folder selection
        folder_layout = QHBoxLayout()
        self.folder_input = QLineEdit()
        folder_layout.addWidget(QLabel("Folder:"))
        folder_layout.addWidget(self.folder_input)
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self.browse_folder)
        folder_layout.addWidget(browse_button)
        self.layout.addLayout(folder_layout)

        # Search text
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        search_layout.addWidget(QLabel("Search Text:"))
        search_layout.addWidget(self.search_input)
        self.layout.addLayout(search_layout)

        # Options
        options_layout = QHBoxLayout()
        self.recursive_check = QCheckBox("Recursive")
        self.case_sensitive_check = QCheckBox("Case Sensitive")
        self.exact_match_check = QCheckBox("Exact Match")
        options_layout.addWidget(self.recursive_check)
        options_layout.addWidget(self.case_sensitive_check)
        options_layout.addWidget(self.exact_match_check)
        self.layout.addLayout(options_layout)

        # Search button
        search_button = QPushButton("Search")
        search_button.clicked.connect(self.search_files)
        self.layout.addWidget(search_button)

        # Results tree
        self.result_tree = QTreeWidget()
        self.result_tree.setHeaderLabels(["File Path", "File Name"])
        self.result_tree.itemDoubleClicked.connect(self.on_double_click)
        self.result_tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        self.layout.addWidget(self.result_tree)

        # Action buttons
        action_layout = QHBoxLayout()
        copy_button = QPushButton("Copy Selected")
        copy_button.clicked.connect(self.copy_selected)
        move_button = QPushButton("Move Selected")
        move_button.clicked.connect(self.move_selected)
        action_layout.addWidget(copy_button)
        action_layout.addWidget(move_button)
        self.layout.addLayout(action_layout)

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.folder_input.setText(folder)

    def search_files(self):
        folder = self.folder_input.text()
        search_text = self.search_input.text()
        recursive = self.recursive_check.isChecked()
        case_sensitive = self.case_sensitive_check.isChecked()
        exact_match = self.exact_match_check.isChecked()

        if not folder or not search_text:
            QMessageBox.warning(self, "Input Error", "Please provide both folder and search text.")
            return

        self.result_tree.clear()

        for root, dirs, files in os.walk(folder):
            for file in files:
                if fnmatch.fnmatch(file, "*.doc") or fnmatch.fnmatch(file, "*.docx") or fnmatch.fnmatch(file, "*.txt") or fnmatch.fnmatch(file, "*.odt") or fnmatch.fnmatch(file, "*.pdf"):
                    file_path = os.path.join(root, file)
                    if self.file_matches(file_path, search_text, case_sensitive, exact_match):
                        QTreeWidgetItem(self.result_tree, [file_path, file])

            if not recursive:
                break

    def file_matches(self, file_path, search_text, case_sensitive, exact_match):
        try:
            if fnmatch.fnmatch(file_path, "*.docx"):
                doc = Document(file_path)
                return self.search_in_text("\n".join([para.text for para in doc.paragraphs]), search_text, case_sensitive, exact_match)
            elif fnmatch.fnmatch(file_path, "*.odt"):
                doc = load_odt(file_path)
                return self.search_in_text("\n".join([str(para) for para in doc.getElementsByType(P)]), search_text, case_sensitive, exact_match)
            elif fnmatch.fnmatch(file_path, "*.pdf"):
                with open(file_path, "rb") as f:
                    reader = PdfReader(f)
                    for page in reader.pages:
                        if self.search_in_text(page.extract_text(), search_text, case_sensitive, exact_match):
                            return True
                return False
            else:
                return self.search_in_file(file_path, search_text, case_sensitive, exact_match)
        except Exception as e:
            print(f"Error reading file {file_path}: {e}")
            return False

    def search_in_file(self, file_path, search_text, case_sensitive, exact_match):
        chunk_size = 4096  # Read in 4KB chunks
        with io.open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            if not case_sensitive:
                search_text = search_text.lower()
            
            leftovers = ''
            while True:
                chunk = file.read(chunk_size)
                if not chunk:
                    # End of file
                    return self.search_in_text(leftovers, search_text, case_sensitive, exact_match)
                
                text_to_search = leftovers + chunk
                if self.search_in_text(text_to_search, search_text, case_sensitive, exact_match):
                    return True
                
                # Keep the last len(search_text) - 1 characters for the next iteration
                leftovers = text_to_search[-(len(search_text)-1):]
        
        return False

    # def search_in_text(self, text, search_text, case_sensitive, exact_match):
    #     if not case_sensitive:
    #         text = text.lower()
    #         search_text = search_text.lower()
        
    #     if exact_match:
    #         return f" {search_text} " in f" {text} "
    #     else:
    #         return search_text in text

    
    def search_in_text(self, text, search_text, case_sensitive, exact_match):
        if not case_sensitive:
            text = text.lower()
            search_text = search_text.lower()
    
        if exact_match:
            # Match even if it's part of a word
            return search_text in text
        else:
            # Match only whole words
            pattern = r'\b' + re.escape(search_text) + r'\b'
            return bool(re.search(pattern, text))

    def on_double_click(self, item, column):
        file_path = item.text(0)
        self.open_file(file_path)

    def open_file(self, file_path):
        try:
            if platform.system() == 'Darwin':       # macOS
                subprocess.call(('open', file_path))
            elif platform.system() == 'Windows':    # Windows
                os.startfile(file_path)
            else:                                   # linux variants
                subprocess.call(('xdg-open', file_path))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error opening file {file_path}: {e}")

    def copy_selected(self):
        self.move_or_copy_selected("copy")

    def move_selected(self):
        self.move_or_copy_selected("move")

    def move_or_copy_selected(self, action):
        selected_items = self.result_tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Selection Error", "Please select files to " + action + ".")
            return

        target_folder = QFileDialog.getExistingDirectory(self, f"Select Target Folder for {action.capitalize()}")
        if not target_folder:
            return

        for item in selected_items:
            file_path = item.text(0)
            try:
                if action == "copy":
                    shutil.copy(file_path, target_folder)
                else:
                    shutil.move(file_path, target_folder)
                    # Remove the item from the tree if it was moved
                    self.result_tree.takeTopLevelItem(self.result_tree.indexOfTopLevelItem(item))
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error {action}ing file {file_path}: {e}")

        QMessageBox.information(self, "Success", f"Selected files {action}ied successfully.")

if __name__ == "__main__":
    app = QApplication([])
    window = FileSearcherApp()
    window.show()
    app.exec_()
