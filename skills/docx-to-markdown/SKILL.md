---
name: docx-to-markdown
description: "Convert DOCX to Markdown with embedded Excel table conversion and image extraction. Use when (1) converting .docx files to Markdown, especially those containing embedded Excel spreadsheets, (2) extracting images from Word documents to local files with relative paths, (3) batch processing multiple DOCX files, (4) optionally converting Markdown to PDF with Chinese font support. 中文场景：Word 转 Markdown、提取图片、嵌入 Excel 表格转 Markdown、批量转换 docx、Markdown 转 PDF。"
---

# DOCX to Markdown Converter

![Version](https://img.shields.io/badge/version-0.1.6-blue)

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

All commands below assume the working directory is the **skill root** (`skills/docx-to-markdown/`).
Install dependencies first: `pip install -r requirements.txt`

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

# Per-document timeout (default 300s; <=0 disables; skipped on Windows)
python scripts/batch_convert.py <source_dir> <output_dir> --timeout 120

# Force re-convert even if output already exists
python scripts/batch_convert.py <source_dir> <output_dir> --force
```

Each DOCX creates a separate folder with its MD file and assets.

**Skip semantics (SHA-256 sentinel)**: a document is skipped only when the output
folder + MD + valid `.converted` sentinel are all present AND the sentinel's
recorded source SHA-256 matches the current source file. Changed sources are
re-converted automatically — no `--force` needed. Legacy plain-text sentinels are
treated as invalid (re-converted). Failed conversions (including timeouts and
security rejections) clean up the half-finished output folder.

### Markdown to PDF (Optional)

```bash
python scripts/md_to_pdf.py <input.md> [output.pdf] [--engine auto|pandoc|python]
```

If output path is omitted, PDF is saved in the same directory as the input file.

`md_to_pdf.py` is standalone and works independently from this skill:
- `--engine auto` (default): prefer system `pandoc`; if pandoc fails (e.g. default pdflatex cannot render Chinese), automatically retries with `xelatex` + CJK fonts, then falls back to Python renderer
- `--engine pandoc`: force pandoc (raises on failure, no fallback)
- `--engine python`: force Python renderer (`pip install markdown reportlab`)

> If pandoc is available, it often produces better results.
> The Python renderer escapes all text before passing to reportlab, so content with `<`, `>`, `&` is safe.

## Key Features

### Heading Restoration

- Prefer DOCX `heading_*` bookmarks + paragraph style mapping to recover Markdown heading levels (style/depth directly map to H1/H2/H3... without +1 offset)
- Keep original numbering (no automatic renumbering for reset-style sections like repeated `1.` / `1.1`) to avoid mutating bilingual or manually numbered headings
- Conservatively promote a leading full-line bold paragraph to H1 only when later numbered section headings exist

### Embedded Excel Conversion

Automatically detects Excel spreadsheets embedded in DOCX and converts them to Markdown tables:

- Parses `document.xml` OLE object references to find Excel-to-preview-image mappings (robust)
- Uses relationship ID adjacency heuristic to supplement mappings not covered by OLE parsing
- Extracts Excel data using openpyxl (lightweight, no pandas needed)
- Replaces preview images with formatted Markdown tables, with repeat-safe placeholder handling
- Expands merged cells to explicit Markdown grid values for parser/LLM-friendly output:
  - Vertical single-column merge (`rowspan`) → fill down all rows with anchor value
  - Horizontal single-row merge (`colspan`) → fill right all columns with anchor value
  - `n x m` rectangular merge → fill the entire merged rectangle with anchor value
- Adds a blockquote note above each merged-table:
  - `> merge_ranges: A1:B2, C3:C5, ...`
  - Preserves original merge scope metadata for downstream parsers

### Image Handling

- Extracts all images from `word/media/`
- Auto-detects true image format (PNG/JPEG/GIF/WEBP/BMP/TIFF/WMF/EMF) regardless of extension
- Saves with corrected extensions; unrecognized formats keep their original extension instead of being mislabeled as PNG
- Prevents overwrite on corrected-name collisions by appending short hash suffix
- Uses relative paths (`assets/image.png`) in Markdown

### Output Naming Safety

- Cleans invalid filename characters and quote variants
- When cleaning actually replaces/removes characters (e.g. `a:b` → `a_b`) or truncates over-long names, a short hash of the original name is appended so different source names never share one output folder
- Pure NFKC normalization (full-width → half-width, common in Chinese documents) does not append a hash; its rare collisions are covered by the sentinel hash check

### Security: Resource-Exhaustion Defense (Malicious DOCX)

Before decompression, the ZIP structure is validated against
`DOCX_SECURITY_LIMITS` (module-level dict, tunable):

| Limit | Default |
|-------|---------|
| Total uncompressed size | 500 MB |
| Single entry uncompressed | 100 MB |
| Single entry compression ratio | 100x |
| Total compression ratio (only judged when compressed total > 1 MB) | 100x |
| Image count (`word/media/`) | 500 |
| Single image file size | 20 MB |
| Single image pixels (decompression-bomb check via header parsing, no decode) | 50,000,000 |
| Embedded Excel size | 50 MB |

Violations raise `DocxSecurityError(ValueError)`. It subclasses `ValueError`
for compatibility, but callers can catch it precisely: a security rejection
means the input is malicious/abnormal and must NOT be retried or degraded to
a fallback path. Actual reads are additionally bounded at decompression time
(`read_zip_entry_bounded`) to defend against lying ZIP metadata.

XML parsing goes through `defusedxml` when installed (falls back to stdlib
`xml.etree` otherwise); `defusedxml` is listed as an optional dependency.

### Batch Reliability

- **Per-document timeout**: `--timeout` (default 300s) via POSIX `signal.alarm`;
  `<=0` disables. Windows (no SIGALRM) and non-main threads degrade to no timeout.
  The previous signal handler is always restored.
- **SHA-256 `.converted` sentinel** (atomic tmp+rename JSON) — see Batch Conversion above
- **Half-finished cleanup**: failed conversions and untrusted pre-existing outputs
  (missing/invalid/mismatched sentinel) are removed and re-converted

### Additional Enhancements

- **Excel date/number formatting**: `datetime` with zero time → `YYYY-MM-DD`; integer `float` → no `.0`
- **Footnotes**: mammoth footnote HTML → Markdown `[^N]` / `[^N]: text` syntax
- **Text boxes**: Extracts `<w:txbxContent>` content ignored by mammoth, appended as blockquote
- **Math formulas**: Extracts OMML text nodes, wraps in `$$ ... $$` (basic detection, not full LaTeX)
- **Residual text cleanup**: Removes "点击图片可查看完整电子表格" after table replacement

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
| Footnotes | ✅ → `[^N]` syntax |
| Text boxes | ✅ → Appended blockquote |
| Math (OMML) | ⚠️ Text extraction only |

## Dependencies

```bash
pip install -r requirements.txt
# Installs: mammoth, openpyxl (core) + markdown, reportlab (optional PDF engine)
#           + defusedxml (optional hardened XML parsing)
```

- **mammoth** + **openpyxl**: DOCX→Markdown 核心转换（必需）
- **markdown** + **reportlab**: Python 内置 PDF 渲染引擎（仅 `md_to_pdf.py --engine python` 时使用；若系统有 pandoc 则可不装）
- **defusedxml**: 增强解析 DOCX 内 XML 的安全性（防实体膨胀/外部实体；未安装自动回退标准库，功能不受影响）

## Scripts Reference

| Script | Purpose | Dependencies |
|--------|---------|-------------|
| `convert_docx.py` | Core converter: DOCX → Markdown + images | `requirements.txt` |
| `batch_convert.py` | Batch process directory of DOCX files | `requirements.txt` |
| `md_to_pdf.py` | Standalone Markdown → PDF (Chinese support) | `pandoc` (recommended) OR `markdown` + `reportlab` |

For detailed API and customization, see [references/usage-guide.md](references/usage-guide.md).
