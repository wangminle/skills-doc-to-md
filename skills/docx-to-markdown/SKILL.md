---
name: docx-to-markdown
description: "Convert Microsoft Word documents (.docx) to Markdown format with image extraction and embedded Excel table conversion. Use when (1) Converting DOCX files to Markdown, (2) Extracting images from Word documents, (3) Converting embedded Excel spreadsheets to Markdown tables, (4) Batch processing multiple DOCX files, (5) Optionally converting Markdown to PDF with Chinese font support."
---

# DOCX to Markdown Converter

Convert Word documents to Markdown with full support for images, embedded Excel tables, and batch processing.

## When to Use This Skill vs Alternatives

| Scenario | Recommended Approach |
|----------|---------------------|
| DOCX contains **embedded Excel tables** | **Use this skill** (unique capability) |
| DOCX has images, need local files + relative paths | Use this skill, or `pandoc --extract-media` |
| Simple DOCX, text only | `pandoc input.docx -o output.md` (zero dependencies) |
| Need to edit/redline DOCX (not convert) | Use the `docx` skill instead |

**Core value**: This skill handles the case that pandoc and markitdown cannot — converting embedded Excel spreadsheets into Markdown tables while extracting images locally.

## Workflow Overview

```
1. Single file conversion    → Run convert_docx.py
2. Batch conversion          → Run batch_convert.py  
3. Markdown to PDF (optional)→ Run md_to_pdf.py
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

When user mentions converting multiple DOCX files, use batch conversion:

```bash
python scripts/batch_convert.py <source_dir> <output_dir>
```

Each DOCX creates a separate folder with its MD file and assets.

### Markdown to PDF (Optional)

```bash
python scripts/md_to_pdf.py <input.md> [output.pdf] [--engine auto|pandoc|python]
```

If output path is omitted, PDF is saved in the same directory as the input file.

`md_to_pdf.py` is standalone and works independently from this skill:
- `--engine auto` (default): prefer system `pandoc`, fallback to Python renderer
- `--engine pandoc`: force pandoc
- `--engine python`: force Python renderer (`pip install markdown reportlab`)

> If pandoc is available, it often produces better results.

## Key Features

### Embedded Excel Conversion

Automatically detects Excel spreadsheets embedded in DOCX and converts them to Markdown tables:

- Parses `document.xml` OLE object references to find Excel-to-preview-image mappings (robust)
- Uses relationship ID adjacency heuristic to supplement mappings not covered by OLE parsing
- Extracts Excel data using openpyxl (lightweight, no pandas needed)
- Replaces preview images with formatted Markdown tables, with repeat-safe placeholder handling

### Image Handling

- Extracts all images from `word/media/`
- Auto-detects true image format (PNG/JPEG/GIF/WEBP/BMP) regardless of extension
- Saves with corrected extensions
- Prevents overwrite on corrected-name collisions by appending short hash suffix
- Uses relative paths (`assets/image.png`) in Markdown

### Output Naming Safety

- Cleans invalid filename characters and quote variants
- For very long document names, truncates safely and appends a short hash suffix to avoid directory collisions

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

### Core (required)

```bash
pip install -r requirements.txt
# Installs: mammoth, openpyxl
```

## Scripts Reference

| Script | Purpose | Dependencies |
|--------|---------|-------------|
| `convert_docx.py` | Core converter: DOCX → Markdown + images | `requirements.txt` |
| `batch_convert.py` | Batch process directory of DOCX files | `requirements.txt` |
| `md_to_pdf.py` | Standalone Markdown → PDF (Chinese support) | `pandoc` (recommended) OR `markdown` + `reportlab` |

For detailed API and customization, see [references/usage-guide.md](references/usage-guide.md).
