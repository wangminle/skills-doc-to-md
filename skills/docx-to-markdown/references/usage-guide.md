# DOCX to Markdown 使用指南

## 目录

1. [convert_docx.py 详解](#convert_docxpy-详解)
2. [batch_convert.py 详解](#batch_convertpy-详解)
3. [md_to_pdf.py 详解](#md_to_pdfpy-详解)
4. [常见问题](#常见问题)
5. [自定义扩展](#自定义扩展)

---

## convert_docx.py 详解

> **执行目录前提**：以下所有 `python scripts/...` 命令均假设当前工作目录为 skill 根目录 `skills/docx-to-markdown/`。
> 若从仓库根目录执行，需加上路径前缀：`python skills/docx-to-markdown/scripts/...`

### 核心功能

将单个 DOCX 文档转换为 Markdown，同时提取图片和转换嵌入的 Excel 表格。

### 命令行用法

```bash
# 在 skills/docx-to-markdown/ 目录下执行
python scripts/convert_docx.py <docx文件路径> <输出目录>
```

**示例：**
```bash
python scripts/convert_docx.py report.docx ./output
```

**输出（自动创建以文件名命名的子文件夹）：**
```
output/
└── report/           # 自动创建的子文件夹
    ├── report.md
    └── assets/
        ├── image1.png
        ├── image2.jpeg
        └── ...
```

### 核心函数

#### `convert_docx_to_markdown(docx_path, output_dir, create_subfolder=True)`

主入口函数，执行完整的转换流程。

**参数：**
- `docx_path`: DOCX 文件路径
- `output_dir`: 输出目录路径
- `create_subfolder`: 是否在输出目录下创建以文件名命名的子文件夹（默认 True）

**返回：** 生成的 Markdown 文件路径

**异常：**
- 当输入文件不是有效 DOCX/ZIP，或缺少 `word/document.xml` 时抛出 `ValueError`

**输出结构：**
- 当 `create_subfolder=True` 时：`output_dir/文件名/文件名.md` + `assets/`
- 当 `create_subfolder=False` 时：`output_dir/文件名.md` + `assets/`

#### `parse_relationships(docx_path)`

解析 DOCX 内部的关系文件，识别 Excel 嵌入与预览图的映射关系。

**策略（双重保险）：**
1. **优先**：解析 `document.xml` 中的 `<w:object>` 节点，从 `<o:OLEObject>` 和 `<v:imagedata>` 提取真实的 rId 配对（可靠，不依赖 ID 排列顺序）
2. **补全**：对方法 1 未覆盖的 Excel 项，使用 rId 相邻启发式补全映射（兼容非标准生成器）

**返回：** `(excel_to_preview, preview_to_excel, ordered_pairs)` 三元组
- `excel_to_preview`: Excel 文件 → 预览图文件 映射
- `preview_to_excel`: 预览图文件 → Excel 文件 映射
- `ordered_pairs`: `[(Excel路径, 预览图路径), ...]`，按文档出现顺序排列（用于队列消费）

#### `excel_to_markdown(xlsx_data)`

将 Excel 二进制数据转换为 Markdown 表格（仅依赖 openpyxl，无需 pandas）。

自动清理全空行和全空列，补齐短行。

**参数：** `xlsx_data` - Excel 文件的二进制内容

**返回：** Markdown 表格字符串，失败返回 None

#### `detect_image_format(image_data)`

通过文件头检测图片真实格式。

**支持格式：**
- PNG (magic: `\x89PNG\r\n\x1a\n`)
- JPEG (magic: `\xff\xd8`)
- GIF (magic: `GIF87a` / `GIF89a`)
- WEBP (magic: `RIFF...WEBP`)
- BMP (magic: `BM`)

#### `html_to_markdown(html, heading_level_map=None)`

将 HTML 转换为 Markdown，支持：

| HTML 元素 | Markdown 输出 |
|----------|--------------|
| `<h1>-<h6>` | `#` - `######` |
| `<strong>`, `<b>` | `**text**` |
| `<em>`, `<i>` | `*text*` |
| `<img>` | `![](path)` |
| `<a>` | `[text](url)` |
| `<ul>/<li>` | `- item` |
| `<table>` | Markdown 表格 |

补充说明：
- `heading_level_map`（可选）用于按 DOCX 原始 heading 样式覆盖标题层级
- 标题层级映射规则：段落样式编号/文本编号深度与 Markdown `#` 层级直接对应（不再 `+1` 偏移）
- 编号标题默认保持原始编号，不进行自动重排
- 文首整行加粗仅在后续存在“编号章节标题”时提升为一级标题
- `<img src=...>` 和 `<a href=...>` 支持双引号、单引号、无引号三种写法
- HTML 实体通过 `html.unescape()` 统一解码，覆盖所有命名和数字实体

#### `_convert_footnotes(html)`

将 mammoth 生成的脚注 HTML 转换为 Markdown 脚注语法 `[^N]` / `[^N]: text`。

**mammoth 输出格式：**
- 正文引用: `<sup><a href="#footnote-N">[N]</a></sup>`
- 文末列表: `<li id="footnote-N"><p>text ↑</p></li>`

**转换结果：**
- 正文引用 → `[^N]`
- 文末 → `[^N]: text`（以 `---` 分隔线分隔）

#### `extract_textbox_content(docx_path)`

从 DOCX 的 `document.xml` 中提取文本框 `<w:txbxContent>` 的纯文本内容。

mammoth 通常忽略文本框/形状中的内容，此函数作为补充提取。返回非空文本块列表。

#### `extract_math_text(docx_path)`

从 DOCX 的 `document.xml` 中提取 OMML 数学公式 `<m:oMath>` 的纯文本。

完整的 OMML→LaTeX 转换极为复杂，此函数仅提取公式中的文本节点，用 `$$ ... $$` 包裹作为占位标记。

#### `_format_cell_value(cell)`

将 openpyxl 单元格值转为友好字符串：
- `datetime` 仅含日期 → `YYYY-MM-DD`（不输出 `00:00:00`）
- `datetime` 含时间 → `YYYY-MM-DD HH:MM:SS`
- `date` → `YYYY-MM-DD`
- `time` → `HH:MM:SS`
- 整数 `float` → `int`（如 `3.0` → `3`）

---

## batch_convert.py 详解

### 核心功能

批量转换目录下所有 DOCX 文件，每个文件生成独立的输出文件夹。

### 命令行用法

```bash
python scripts/batch_convert.py [源目录] [输出目录] [--force]
```

**默认值：**
- 源目录: `1-Reference`
- 输出目录: `2-Temp`

**示例：**
```bash
python scripts/batch_convert.py ./documents ./markdown_output

# 强制重新转换已存在的输出目录
python scripts/batch_convert.py ./documents ./markdown_output --force
```

### 核心函数

#### `batch_convert(source_dir, output_dir, force=False)`

**参数：**
- `source_dir`: 源文件目录
- `output_dir`: 输出目录
- `force`: 为 `True` 时强制重新转换已存在的输出目录（删除旧目录后重新生成）

### 特性

1. **自动跳过** - 已存在的输出目录会被跳过（使用 `--force` 可强制重新转换）
2. **`--force` 模式** - 删除已有输出目录后重新转换，适合文档更新后需要重新生成的场景
3. **进度显示** - 显示 `[当前/总数]` 进度
4. **统计汇总** - 结束时显示成功/失败数量
5. **文件名清理与防冲突** - 自动清理非法字符；超长文件名会附加短 hash
6. **大小写去重** - macOS 等大小写不敏感文件系统上自动去重 `.docx`/`.DOCX`

### 输出结构

```
output_dir/
├── Document1/
│   ├── Document1.md
│   └── assets/
├── Document2/
│   ├── Document2.md
│   └── assets/
└── ...
```

---

## md_to_pdf.py 详解（可选功能）

> **独立运行**：该脚本可独立于本 skill 使用。
> **推荐模式**：若系统已安装 pandoc，优先走 pandoc 引擎，效果通常更好。

### 核心功能

将 Markdown 文件转换为 PDF，支持中文字体。

### 命令行用法

```bash
python scripts/md_to_pdf.py <markdown文件路径> [pdf输出路径] [--engine auto|pandoc|python]
```

**示例：**
```bash
python scripts/md_to_pdf.py document.md                           # 默认 auto（优先 pandoc）
python scripts/md_to_pdf.py document.md output.pdf --engine auto
python scripts/md_to_pdf.py document.md output.pdf --engine pandoc
python scripts/md_to_pdf.py document.md output.pdf --engine python
```

### 核心函数

#### `convert_md_to_pdf(input_file, output_file=None)`

**参数：**
- `input_file`: Markdown 文件路径
- `output_file`: PDF 输出路径（可选，默认为输入文件同目录下同名 .pdf 文件）

**返回：** 生成的 PDF 文件路径字符串，失败抛出异常

#### `run_conversion(input_file, output_file, engine)`

**引擎策略：**
1. `auto`：先尝试 `pandoc`，不可用则回退 Python 渲染
2. `pandoc`：强制使用 pandoc
3. `python`：强制使用 Python 渲染（需安装 `markdown` + `reportlab`）

### 中文字体支持

自动检测系统字体：

| 系统 | 字体路径 |
|-----|---------|
| macOS | `/System/Library/Fonts/PingFang.ttc` |
| Linux | `/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf` |
| Windows | `C:/Windows/Fonts/msyh.ttc` |

如果系统字体无法注册，会自动回退到 ReportLab 内置的 CID 字体（如 `STSong-Light`）。

### 样式配置

| 元素 | 字体大小 | 颜色 |
|-----|---------|-----|
| 标题 (H1) | 18pt | #1a5490 |
| 标题 (H2) | 14pt | #1a5490 |
| 标题 (H3) | 12pt | #2c3e50 |
| 正文 | 10pt | 黑色 |
| 列表 | 10pt | 黑色，缩进 20pt |

### 页面设置

- 纸张: A4
- 页边距: 2cm (上下左右)

---

## 常见问题

### Q: 图片没有正确提取？

检查 DOCX 文件结构，确保图片在 `word/media/` 目录下。某些第三方工具生成的 DOCX 可能有不同结构。

### Q: Excel 表格没有转换？

确认：
1. Excel 是嵌入对象（不是链接）
2. 安装了 `openpyxl`（`requirements.txt` 已包含）
3. 如果文档由 WPS/LibreOffice 等非 Microsoft Office 生成，OLE 引用结构可能不同，脚本会自动用启发式补全

### Q: 同一个表格预览图在文档里出现多次，会不会后面失效？

不会。脚本使用“队列 + 重复兜底”策略处理占位替换，同一预览图重复出现时会持续替换为表格，不会退化成普通图片。

### Q: 图片提取时会不会因为扩展名修正而覆盖同名文件？

默认不会。若修正扩展名后发生重名且内容不同，脚本会自动追加短 hash 后缀（如 `_a1b2c3d4`）避免覆盖。

### Q: PDF 中文显示为方块？

确保系统有支持的中文字体。脚本会依次尝试系统字体并回退到 ReportLab 内置 CID 字体（如 `STSong-Light`）。

### Q: 批量转换时某些文件失败？

查看控制台输出的错误信息，常见原因：
- 文件损坏
- 密码保护
- 非标准 DOCX 格式

### Q: 文档更新后想重新转换，但输出已存在怎么办？

使用 `--force` 参数强制重新转换：
```bash
python scripts/batch_convert.py ./documents ./output --force
```
该模式会删除已有输出目录后重新生成。

### Q: 脚注能自动转换吗？

可以。mammoth 生成的脚注 HTML 会自动转换为 Markdown `[^N]` 脚注语法，脚注正文以 `---` 分隔线追加在文档末尾。

### Q: 文档中的文本框或数学公式会被提取吗？

- **文本框**：mammoth 通常会忽略 `<w:txbxContent>` 中的内容，脚本会自动提取并以引用块追加在文档末尾。
- **数学公式**：OMML 数学公式的纯文本会被提取并以 `$$ ... $$` 标记输出。完整的 OMML→LaTeX 转换需要额外工具（如 pandoc 的 `--mathml` 选项）。

### Q: Excel 表格中的日期显示了多余的 00:00:00？

已修复。`datetime` 值当时间部分全为零时，仅输出 `YYYY-MM-DD`；整数 `float` 如 `3.0` 输出为 `3`。

---

## 自定义扩展

### 添加新的图片格式支持

在 `detect_image_format()` 中添加新的 magic bytes 检测：

```python
elif image_data[:4] == b'新格式magic':
    return '.新扩展名'
```

### 自定义 HTML 转 Markdown 规则

在 `html_to_markdown()` 中添加新的正则替换：

```python
html = re.sub(r'<custom>(.*?)</custom>', r'自定义格式\1', html, flags=re.DOTALL)
```

### 添加新的 PDF 样式

在 `md_to_pdf.py` 中创建新的 `ParagraphStyle`：

```python
custom_style = ParagraphStyle(
    'CustomStyle',
    parent=styles['Normal'],
    fontName=chinese_font,
    fontSize=12,
    textColor=colors.HexColor('#颜色'),
)
```
