# skills-docstomd-withpics

本仓库用于发布一个 Codex Skill：`docx-to-markdown`（`skills/docx-to-markdown/SKILL.md:1`），用于将 Word 的 `.docx` 文档转换为 Markdown，并提取图片/转换嵌入的 Excel 表格。

## 核心价值

与 pandoc、markitdown 等通用方案相比，本 skill 的独特价值在于：
- **嵌入 Excel 表格 → Markdown 表格**：优先解析 DOCX 内部 OLE 对象引用，再用关系启发式补全映射，将嵌入的 Excel 数据转为可读 Markdown 表格
- **图片提取到本地文件**：输出到 `assets/`，Markdown 使用相对路径引用

> 如果你的文档不含嵌入 Excel，`pandoc --extract-media ./media input.docx -o output.md` 即可满足需求。

## 功能

- `.docx → .md`：保留常见标题、段落、列表、链接、表格
- 标题还原增强：优先使用 DOCX `heading_*` 样式映射标题层级，并对重复重置的编号小节做局部重排
- 图片提取：输出到 `assets/`，Markdown 使用相对路径引用
- 嵌入 Excel：将嵌入的 `.xlsx` 转为 Markdown 表格
- 合并单元格语义展开：将 Excel 合并格展开为 Markdown 管道表的显式网格值
- 批处理：批量转换目录下多个 DOCX
- `.md → .pdf`（可选）：独立脚本，优先走 pandoc，支持中文字体回退
- 文件名稳健：自动清理非法字符；超长文件名会附加短 hash，避免不同文件名截断后冲突

## 快速开始

在 skill 目录下执行：

```bash
cd skills/docx-to-markdown

# 安装核心依赖（mammoth + openpyxl）
pip install -r requirements.txt

# 单文件
python scripts/convert_docx.py <input.docx> <output_dir>

# 批量
python scripts/batch_convert.py <source_dir> <output_dir>

# Markdown 转 PDF（独立脚本，默认优先 pandoc）
python scripts/md_to_pdf.py <input.md> [output.pdf] [--engine auto|pandoc|python]

# 若强制使用 Python 引擎，再安装依赖
pip install markdown reportlab
```

## 输出结构

每个 DOCX 会生成一个同名目录：

```
output_dir/
└── document_name/
    ├── document_name.md
    └── assets/
        ├── image1.png
        └── ...
```

## 依赖说明

| 文件 | 包含 | 用途 |
|------|------|------|
| `requirements.txt` | mammoth, openpyxl | 核心转换功能 |
| `scripts/md_to_pdf.py` | pandoc 或 markdown+reportlab | 独立的 MD→PDF 功能 |

## 注意事项

- "嵌入 Excel → Markdown 表格"使用双策略：优先解析 `document.xml` OLE 引用，再对未覆盖项使用 rId 相邻启发式补全。不同 Office/WPS 生成的文档可能存在差异。
- Excel 合并单元格在 Markdown 中采用“全展开”语义（便于程序与 LLM 读取）：
  - 纵向单列合并（`rowspan`）: 用左上角锚点值向下填充每一行
  - 横向单行合并（`colspan`）: 用左上角锚点值向右填充每一列
  - `n x m` 矩形合并: 用左上角锚点值填满整个矩形区域
- 对包含合并单元格的 Excel 表格，Markdown 表格上方会追加引用说明：
  - `> merge_ranges: A1:B2, C3:C5, ...`
  - 用于保留原始合并范围信息，便于程序/LLM 做语义还原
- 当同一预览图在文档中重复出现时，脚本会优先持续替换为表格；不会因为队列耗尽退化为普通图片。
- 修正图片扩展名后若发生重名冲突，脚本会自动追加 hash 后缀，避免覆盖已提取文件。
- 输入文件若不是有效 DOCX/ZIP（或缺少 `word/document.xml`），会抛出明确错误信息。
- 如果只需要简单的 DOCX 转 MD（无嵌入 Excel），推荐直接使用 `pandoc --extract-media ./media input.docx -o output.md`。

## License

Apache-2.0（见 `LICENSE`）。
