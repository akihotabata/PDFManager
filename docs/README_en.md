# PDFManager (English)

A Windows desktop app to merge, split, and edit PDF files locally.  
Built with PySide6 (Qt), using pypdf for merge/split and PyMuPDF for preview.

## Features
- Merge PDFs with sorting and drag-reorder
- Split by page, by N pages, or custom ranges (e.g., 1-3,5,7-10)
- Page operations: delete, rotate, extract, duplicate, insert
- Preview (PyMuPDF) with zoom, fit-to-width, fit-to-page
- Progress bar, logs, Japanese path support

## Requirements
- Windows 10/11, Python 3.10+  
- `pip install -r tools/requirements.txt`

## Run
```bat
tools\run.bat
```

## Build EXE (optional)
```bat
tools\build_exe.bat
```

License: MIT
