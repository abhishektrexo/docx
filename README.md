# DOCX Page Finder

A Python tool to find the page numbers of specific text occurrences in `.docx` files. It handles DOCX pagination by rendering the document to PDF using LibreOffice (headless) and then searching the text layer.

## Features

- **Accurate Pagination**: Uses LibreOffice rendering engine for layout accuracy (~95% match with MS Word).
- **Batch Processing**: Scan a single file or an entire directory of DOCX files.
- **Context Snippets**: Shows the text surrounding each match.
- **Case Sensitivity**: Optional case-sensitive search.
- **PDF Retention**: Option to keep the intermediate PDF for debugging.

## Prerequisites

- **Python 3.6+**
- **LibreOffice**: Must be installed and accessible via command line (`libreoffice`).
- **Python Dependencies**:
  ```bash
  pip install pymupdf
  ```

## Usage

### 1. Search a Single File
```bash
python3 find_text_pages.py document.docx "search text"
```

### 2. Search a Directory (Batch Mode)
```bash
python3 find_text_pages.py ./docs_folder "search text"
```

### Options
| Flag | Description |
|------|-------------|
| `-v`, `--verbose` | Show text snippets around matches |
| `-c`, `--case-sensitive` | Enable case-sensitive search |
| `--keep-pdf` | Keep the generated PDF file (for debugging) |
| `-o DIRECTORY` | Specify output directory for PDFs |

## Example Output

```text
Processing 2 files in './docs'...

Scanning: Specification.docx... FOUND (3 matches)
  Pages: 1, 5, 12
  Details:
    Page 1 (1 matches):
      ...The present invention relates generally to...

Scanning: Old_Draft.docx... No matches.

Summary: Found 3 total occurrences in 1 of 2 files.
```

## How It Works

1. **Convert**: The script uses `libreoffice --headless --convert-to pdf` to render the DOCX to a temporary PDF.
2. **Search**: It uses `pymupdf` (fitz) to search the PDF's text layer.
3. **Map**: It maps the text occurrences to their page numbers in the rendered PDF.
