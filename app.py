import os
import fnmatch
import shutil
import subprocess
import platform
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
from odf.opendocument import load as load_odt
from odf.text import P
from PyPDF2 import PdfReader
import re
import win32com.client

class FileSearcherApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("File Searcher")
        self.geometry("800x600")
        self.create_widgets()

    def create_widgets(self):
        # Folder selection
        folder_frame = ttk.Frame(self)
        folder_frame.pack(fill=tk.BOTH, padx=10, pady=5)
        ttk.Label(folder_frame, text="Folder:").pack(side=tk.LEFT)
        self.folder_input = ttk.Entry(folder_frame)
        self.folder_input.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, padx=(5, 5))
        ttk.Button(folder_frame, text="Browse", command=self.browse_folder).pack(side=tk.RIGHT)

        # Search text
        search_frame = ttk.Frame(self)
        search_frame.pack(fill=tk.BOTH, padx=10, pady=5)
        ttk.Label(search_frame, text="Search Text:").pack(side=tk.LEFT)
        self.search_input = ttk.Entry(search_frame)
        self.search_input.pack(side=tk.LEFT, expand=tk.YES, fill=tk.X, padx=(5, 0))

        # Options
        options_frame = ttk.Frame(self)
        options_frame.pack(fill=tk.BOTH, padx=10, pady=5)
        self.recursive_var = tk.BooleanVar()
        self.case_sensitive_var = tk.BooleanVar()
        self.exact_match_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Recursive", variable=self.recursive_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(options_frame, text="Case Sensitive", variable=self.case_sensitive_var).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(options_frame, text="Exact Match", variable=self.exact_match_var).pack(side=tk.LEFT)

        # Search button
        ttk.Button(self, text="Search", command=self.search_files).pack(pady=10)

        # Results tree
        self.result_tree = ttk.Treeview(self, columns=("path", "name"), show="headings")
        self.result_tree.heading("path", text="File Path")
        self.result_tree.heading("name", text="File Name")
        self.result_tree.pack(expand=tk.YES, fill=tk.BOTH, padx=10, pady=5)
        self.result_tree.bind("<Double-1>", self.on_double_click)

        # Scrollbar for tree
        tree_scroll = ttk.Scrollbar(self, orient="vertical", command=self.result_tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.configure(yscrollcommand=tree_scroll.set)

        # Action buttons
        action_frame = ttk.Frame(self)
        action_frame.pack(fill=tk.BOTH, padx=10, pady=10)
        ttk.Button(action_frame, text="Copy Selected", command=self.copy_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(action_frame, text="Move Selected", command=self.move_selected).pack(side=tk.LEFT)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_input.delete(0, tk.END)
            self.folder_input.insert(0, folder)

    def search_files(self):
        folder = self.folder_input.get()
        search_text = self.search_input.get()
        recursive = self.recursive_var.get()
        case_sensitive = self.case_sensitive_var.get()
        exact_match = self.exact_match_var.get()

        if not folder or not search_text:
            messagebox.showwarning("Input Error", "Please provide both folder and search text.")
            return

        self.result_tree.delete(*self.result_tree.get_children())

        for root, dirs, files in os.walk(folder):
            for file in files:
                if fnmatch.fnmatch(file, "*.doc") or fnmatch.fnmatch(file, "*.docx") or fnmatch.fnmatch(file, "*.txt") or fnmatch.fnmatch(file, "*.odt") or fnmatch.fnmatch(file, "*.pdf"):
                    file_path = os.path.join(root, file)
                    if self.file_matches(file_path, search_text, case_sensitive, exact_match):
                        self.result_tree.insert("", "end", values=(file_path, file))

            if not recursive:
                break

    def file_matches(self, file_path, search_text, case_sensitive, exact_match):
        try:
            if fnmatch.fnmatch(file_path, "*.docx"):
                doc = Document(file_path)
                return self.search_in_text("\n".join([para.text for para in doc.paragraphs]), search_text, case_sensitive, exact_match)
            elif fnmatch.fnmatch(file_path, "*.doc"):
                if platform.system() == 'Windows':
                    word = win32com.client.Dispatch("Word.Application")
                    word.visible = False
                    doc = word.Documents.Open(file_path)
                    text = doc.Range().Text
                    doc.Close(False)
                    word.Quit()
                    return self.search_in_text(text, search_text, case_sensitive, exact_match)
                else:
                    # Alternative approach for non-Windows platforms
                    # You can use a library like `antiword` or `textract` to extract text from .doc files
                    # or handle the case where .doc files are not supported
                    print(f"Warning: .doc files are not supported on this platform ({platform.system()})")
                    return False
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

    def on_double_click(self, event):
        item = self.result_tree.selection()[0]
        file_path = self.result_tree.item(item, "values")[0]
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
            messagebox.showerror("Error", f"Error opening file {file_path}: {e}")

    def copy_selected(self):
        self.move_or_copy_selected("copy")

    def move_selected(self):
        self.move_or_copy_selected("move")

    def move_or_copy_selected(self, action):
        selected_items = self.result_tree.selection()
        if not selected_items:
            messagebox.showwarning("Selection Error", f"Please select files to {action}.")
            return

        target_folder = filedialog.askdirectory(title=f"Select Target Folder for {action.capitalize()}")
        if not target_folder:
            return

        for item in selected_items:
            file_path = self.result_tree.item(item, "values")[0]
            try:
                if action == "copy":
                    shutil.copy(file_path, target_folder)
                else:
                    shutil.move(file_path, target_folder)
                    self.result_tree.delete(item)
            except Exception as e:
                messagebox.showerror("Error", f"Error {action}ing file {file_path}: {e}")

        messagebox.showinfo("Success", f"Selected files {action}ied successfully.")

if __name__ == "__main__":
    app = FileSearcherApp()
    app.mainloop()
