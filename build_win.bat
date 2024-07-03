@echo off

python -m nuitka ^
    --follow-imports ^
    --standalone ^
    --enable-plugin=tk-inter ^
    --include-package=docx ^
    --include-package=odf ^
    --include-package=PyPDF2 ^
    --include-package=tkinter ^
    --include-package=_tkinter ^
    --windows-console-mode=disable ^
    --output-dir=dist ^
    app.py

echo Build complete. Executable is in the dist folder.

pause

