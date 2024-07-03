#!/bin/bash

# Find the Tcl/Tk directory
TCL_TK_DIR=$(dirname $(python3 -c "import tkinter; print(tkinter.__file__)"))

nuitka3 \
    --follow-imports \
    --standalone \
    --enable-plugin=tk-inter \
    --include-package=docx \
    --include-package=odf \
    --include-package=PyPDF2 \
    --include-package=tkinter \
    --include-package=_tkinter \
    --include-data-dir="$TCL_TK_DIR"=tcl \
    --output-dir=dist \
    app.py

echo "Build complete. Executable is in the dist folder."

