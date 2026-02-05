---
name: docx-to-markdown
description: "Convert Microsoft Word documents (.docx) to Markdown format with image extraction and embedded Excel table conversion. Use when (1) Converting DOCX files to Markdown, (2) Extracting images from Word documents, (3) Converting embedded Excel spreadsheets to Markdown tables, (4) Batch processing multiple DOCX files, (5) Converting Markdown to PDF with Chinese font support."
---

# DOCX to Markdown Converter

Convert Word documents to Markdown with full support for images, embedded Excel tables, and batch processing.

## Workflow Overview

```
1. Single file conversion    → Run convert_docx.py
2. Batch conversion          → Run batch_convert.py  
3. Markdown to PDF           → Run md_to_pdf.py
```

## Quick Start

### Single File Conversion

```bash
python scripts/convert_docx.py <input.docx> <output_directory>
```

Output structure (auto-creates subfolder named after the document):
```
output_directory/
└── document_name/        # Auto-created folder
    ├── document_name.md  # Markdown file
    └── assets/           # Extracted images
        ├── image1.png
        └── image2.jpeg
```

### Batch Conversion

```bash
python scripts/batch_convert.py <source_dir> <output_dir>
```

Each DOCX creates a separate folder with its MD file and assets.

### Markdown to PDF

Edit `md_to_pdf.py` to set input/output paths, then run:

```bash
python scripts/md_to_pdf.py
```

## Key Features

### Embedded Excel Conversion

Automatically detects Excel spreadsheets embedded in DOCX and converts them to Markdown tables:

- Parses XML relationship files to map Excel files to preview images
- Extracts Excel data using pandas
- Replaces preview images with formatted Markdown tables

### Image Handling

- Extracts all images from `word/media/`
- Auto-detects true image format (PNG/JPEG/GIF/WEBP/BMP) regardless of extension
- Saves with corrected extensions
- Uses relative paths (`assets/image.png`) in Markdown

### Format Support

| Element | Support |
|---------|---------|
| Headings (H1-H6) | ✅ |
| Bold/Italic | ✅ |
| Lists (bullet/numbered) | ✅ |
| Tables | ✅ |
| Images | ✅ |
| Hyperlinks | ✅ |
| Embedded Excel | ✅ → Markdown tables |

## Dependencies

Install before use:

```bash
pip install mammoth pandas openpyxl markdown reportlab
```

## Scripts Reference

| Script | Purpose |
|--------|---------|
| `convert_docx.py` | Core converter: DOCX → Markdown + images |
| `batch_convert.py` | Batch process directory of DOCX files |
| `md_to_pdf.py` | Convert Markdown to PDF (Chinese support) |

For detailed API and customization, see [references/usage-guide.md](references/usage-guide.md).
