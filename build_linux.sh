#!/bin/bash

nuitka3 \
    --follow-imports \
    --standalone \
    --enable-plugin=tk-inter \
    --include-package=ttkbootstrap \
    --include-package=docx \
    --include-package=odf \
    --include-package=PyPDF2 \
    --output-dir=dist \
    app.py

echo "Build complete. Executable is in the dist folder."

