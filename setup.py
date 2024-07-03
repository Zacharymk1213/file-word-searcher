import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "packages": ["os", "fnmatch", "shutil", "subprocess", "platform", "io", "re", "docx", "odf", "PyPDF2", "PyQt5"],
    "excludes": [],
    "include_files": []
}

# GUI applications require a different base on Windows (the default is for a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="FileSearcher",
    version="0.1",
    description="File Searcher Application",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py", base=base)]
)

