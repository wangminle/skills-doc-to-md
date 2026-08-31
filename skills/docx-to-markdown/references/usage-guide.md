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
python scripts/convert_docx.py <docx文件路径> <输出目录> [选项]
```

**选项：**
- `--output-name <名称>`：自定义输出子文件夹与 `.md` 文件名（末尾 `.docx`
  自动去除，默认用源文件名）。适合 Web 上传等需要以用户原始文件名命名的场景
- `--on-limit reject|skip`：资源超限处置，默认 `reject` 整篇拒绝；
  `skip` 仅跳过超限资源继续转换（见下文「资源超限处置策略」）

**示例：**
```bash
python scripts/convert_docx.py report.docx ./output
python scripts/convert_docx.py upload1234.docx ./output --output-name 用户原始文件名
python scripts/convert_docx.py report.docx ./output --on-limit skip
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

#### `convert_docx_to_markdown(docx_path, output_dir, create_subfolder=True, output_name=None, on_limit="reject")`

主入口函数，执行完整的转换流程。

**参数：**
- `docx_path`: DOCX 文件路径
- `output_dir`: 输出目录路径
- `create_subfolder`: 是否在输出目录下创建以文件名命名的子文件夹（默认 True）
- `output_name`: 自定义输出命名（默认 None 用源文件名）。末尾 `.docx`（大小写不敏感）
  会自动去除；其他点号后缀保留，因此 `需求V2.4` 不会被截成 `需求V2`。之后经
  `sanitize_stem` 清洗后统一用于子文件夹名、`.md` 文件名与 sentinel 的
  `folder_name`，三处保持一致；sentinel 仍记录真实源文件 SHA-256。空白值抛
  `ValueError`
- `on_limit`: 可降级资源超限处置，`"reject"`（默认，保持既有行为）或
  `"skip"`（见下文「资源超限处置策略」）；其他值抛 `ValueError`

**返回：** 生成的 Markdown 文件路径

**异常：**
- 当输入文件不是有效 DOCX/ZIP，或缺少 `word/document.xml` 时抛出 `ValueError`
- 当输入触发资源耗尽防线（zip bomb / 超大资源，见下文安全防线）时抛出
  `DocxSecurityError`；其中可降级资源超限实际抛其子类
  `ResourceLimitExceeded`（`reject` 模式下处置一致）

**完成标记：** 转换全部成功后会在输出目录原子写入 `.converted` JSON
（`{"folder_name": ..., "source_sha256": ..., "on_limit": ...}`），批处理据此判断
输出完整且与当前源及资源处置策略一致。V0.1.6 缺少策略字段的 JSON sentinel
按 `reject` 兼容读取；策略变化会自动重转。

#### 安全防线（资源耗尽防御）

解压前依据 ZIP 中央目录元数据校验，实际读取时再按真实解压量兜底。阈值集中在
模块级 `DOCX_SECURITY_LIMITS` dict 中，可按需收紧或测试注入。防线分两层：

**第一层：恶意特征，无条件整篇拒绝（`on_limit` 不影响）**

重复 ZIP 条目名也会被拒绝，避免同名条目的不唯一解压语义绕过计数。

| 防线 | 默认阈值 |
|------|---------|
| ZIP 总解压上限 | 500 MB |
| 单 entry 解压上限 | 100 MB |
| 单 entry 压缩比 | 100x |
| 总压缩比（仅当压缩后总大小 > 1MB 时判定，避免小文件舍入误伤） | 100x |

**第二层：可降级资源限制（抛 `ResourceLimitExceeded`，`skip` 模式下降级跳过）**

| 防线 | 默认阈值 |
|------|---------|
| 图片数量（`word/media/`） | 500 |
| 单图文件大小 | 20 MB |
| 单图像素（解压炸弹检测） | 5000 万 |
| 嵌入 Excel 大小 | 50 MB |

`DocxSecurityError(ValueError)`：继承 `ValueError` 以兼容既有异常处理，但调用方
应精确捕获本类——安全拒绝表示输入恶意/异常，**不可降级重试**（降级会绕过防线）。
`ResourceLimitExceeded(DocxSecurityError)`：可降级资源超限专用子类，默认 `reject`
模式下与基类处置完全一致。

#### 资源超限处置策略（`on_limit`）

`convert_docx_to_markdown(..., on_limit="reject"|"skip")`、单文件 CLI
`--on-limit`、批处理 CLI `--on-limit` 三入口统一，默认均为 `reject`：

| 模式 | 行为 |
|------|------|
| `reject`（默认） | 四类可降级资源任一超限即抛 `DocxSecurityError`（子类 `ResourceLimitExceeded`）整篇拒绝，与历史行为完全一致 |
| `skip` | 仅跳过超限资源继续转换：超大嵌入 Excel 不转表格但正文保留；超限图片不落盘，原位置输出可见说明 `*【图片已跳过：单图超过大小上限】*` 等；图片超过数量配额后不再提取。转换结束汇总输出跳过清单日志 |

`skip` 模式的边界与保证：
- 第一层恶意特征（总量/单 entry/压缩比）**依旧整篇拒绝**，跳过策略只对
  “带超大附件的正常文档”生效，对恶意输入不提供任何放宽
- 所有读取路径保持实际上限（含 mammoth 图片回调的有界读取），不会为跳过
  而退化为无界读取
- mammoth 回调复用同一套大小/像素防线与图片数量配额：提取阶段跳过的超限
  图片不可能经回调兜底写盘“复活”
- 文档无超限资源时，`skip` 与 `reject` 产生的 Markdown 正文与
  `assets` 资源文件一致（`.converted` 会分别记录实际策略）

场景建议：服务端无人值守批处理用默认 `reject`（严格）；面向用户的单文档
转换（如 Web 上传）用 `skip`，避免“带一个 60MB 附表的正常文档转不出来”。

XML 解析统一走 `_safe_xml_fromstring` 入口：安装了 `defusedxml` 时防御实体膨胀/
外部实体攻击，未安装自动回退标准库 `xml.etree`（功能等价，仅防护降级）。

#### `validate_docx_zip_security(zip_ref, on_limit="reject")`

对已打开的 `zipfile.ZipFile` 执行上表安全校验（只读声明值不解压）。第一层
超限抛 `DocxSecurityError`；第二层超限在 `reject` 模式抛
`ResourceLimitExceeded`，`skip` 模式不抛（改由提取阶段按真实读取逐项跳过）。
`convert_docx_to_markdown` 在结构校验后自动调用。

#### `image_pixel_count(image_data)`

从图片头部解析宽高并返回像素数（宽×高），用于解压炸弹检测；仅读头部几十字节、
不解码像素。支持 PNG/JPEG/GIF/BMP/WEBP/TIFF；WMF/EMF 等矢量格式及未知格式返回 `None`。

#### `read_zip_entry_bounded(zip_ref, name, max_bytes, error_cls=DocxSecurityError)`

带实际上限的条目读取：边解压边计数，超过 `max_bytes` 立即抛 `error_cls`，
防御中央目录元数据与实际数据不一致的恶意构造。可降级资源（图片/嵌入
Excel）传 `error_cls=ResourceLimitExceeded` 供 `skip` 模式精确捕获。

#### `sha256_file(path)` / `write_conversion_sentinel(...)` / `read_conversion_sentinel(directory)`

完成标记相关工具：流式计算源文件 SHA-256；原子写入（tmp + rename）与读取
`.converted` JSON，记录目录名、源 SHA-256 与 `on_limit`。旧格式（纯文本）、
损坏内容或未知策略返回 `None`；V0.1.6 JSON 缺少策略字段时按 `reject` 返回。

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
调用 openpyxl 前会先校验 XLSX 内层 ZIP 的单条目解压量与压缩比，
嵌套 XLSX zip bomb 在 `skip` 模式下也会无条件拒绝整篇文档。

自动清理全空行和全空列，补齐短行。

**参数：** `xlsx_data` - Excel 文件的二进制内容

**返回：** Markdown 表格字符串，失败返回 None

#### `detect_image_format(image_data)`

通过文件头检测图片真实格式；无法识别时返回 `None`（调用方保留原扩展名，避免误标为 `.png` 损坏文件）。

**支持格式：**
- PNG (magic: `\x89PNG\r\n\x1a\n`)
- JPEG (magic: `\xff\xd8`)
- GIF (magic: `GIF87a` / `GIF89a`)
- WEBP (magic: `RIFF...WEBP`)
- BMP (magic: `BM`)
- TIFF (magic: `II*\x00` / `MM\x00*`)
- WMF (magic: `\xd7\xcd\xc6\x9a`，placeable header)
- EMF (magic: offset 40 处为 ` EMF`)

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
python scripts/batch_convert.py [源目录] [输出目录] [--force] [--timeout 秒数] [--on-limit reject|skip]
```

**默认值：**
- 源目录: `1-Reference`
- 输出目录: `2-Temp`
- 超时: `300` 秒/文档
- 资源超限处置: `reject`（整篇拒绝计失败）

**示例：**
```bash
python scripts/batch_convert.py ./documents ./markdown_output

# 单文档超时 120 秒
python scripts/batch_convert.py ./documents ./markdown_output --timeout 120

# 超大附件降级跳过（正文保留，恶意输入仍计失败）
python scripts/batch_convert.py ./documents ./markdown_output --on-limit skip

# 强制重新转换已存在的输出目录
python scripts/batch_convert.py ./documents ./markdown_output --force
```

### 核心函数

#### `batch_convert(source_dir, output_dir, force=False, timeout=300, on_limit="reject")`

**参数：**
- `source_dir`: 源文件目录
- `output_dir`: 输出目录
- `force`: 为 `True` 时强制重新转换已存在的输出目录（删除旧目录后重新生成）
- `timeout`: 单文档转换超时秒数（`<=0` 不限制；仅 POSIX 主线程生效，
  Windows 无 SIGALRM 自动跳过，非主线程安装失败降级为无超时）
- `on_limit`: 透传给 `convert_docx_to_markdown` 的资源超限处置。默认 `reject`
  超限整篇拒绝计失败；`skip` 仅跳过超限资源继续转换。批处理按源文件名命名
  输出（仅消歧时透传 `output_name`），`skip` 模式下 sentinel 与跳过语义均不受影响

**返回值：** `{"success": 成功数, "skipped": 跳过数, "failed": 失败数}`，
任意文档转换失败时 `failed > 0`，CLI 据此返回退出码 1（全部成功/跳过为 0，
空目录视为成功），供 CI/自动化判定批次结果。

### 特性

1. **SHA-256 完成标记跳过** - 跳过需满足：输出目录 + md + 有效 `.converted`
   sentinel 齐备，且 sentinel 记录的源 SHA-256 与当前源文件、`on_limit` 与当前
   请求一致；**源文件或策略变更后自动重转，无需 `--force`**。旧格式（纯文本）sentinel 视为无效，
   按半成品清理后重转
2. **`--force` 模式** - 删除已有输出目录后重新转换，适合强制全量重建
3. **单文档超时** - `--timeout`（默认 300 秒）基于 POSIX `signal.alarm` 实现，
   超时抛 `TimeoutError` 计为失败；结束后恢复原信号 handler
4. **半成品清理** - 转换失败（含超时/安全拒绝）时清理输出目录；非 `--force`
   路径下既有输出不可信（sentinel 缺失/无效/哈希不匹配）时同样清理重转
5. **进度显示** - 显示 `[当前/总数]` 进度
6. **统计汇总** - 结束时显示成功/跳过/失败数量
7. **文件名清理与防冲突** - 自动清理非法字符；清洗发生字符替换/超长截断时
   附加源文件名短 hash，防止不同原始名称映射到同一输出目录
8. **批内同名消歧** - 两个源文件名清洗后相同时（如 NFKC 归一化的 `A` 与
   全角 `Ａ`），后来者附加原始名短 hash（如 `A_30757541`），避免共用输出
   目录互相覆盖；分配只依赖文件名，跨批次稳定，不影响增量跳过
9. **扩展名大小写不敏感** - `.docx`/`.DOCX`/`.Docx`/`.doCx` 等全部匹配；
   macOS 等大小写不敏感文件系统上的重复条目自动去重；名为 `*.docx` 的
   目录不是待转文档
10. **失败退出码** - 任一文档失败时 CLI 退出码为 1，供 CI/自动化判定

> 安全拒绝（`DocxSecurityError`）在批处理中计为失败并清理输出，不降级不重试。
> `--on-limit skip` 只放宽可降级资源（超大附件）的处置；ZIP bomb 等恶意特征
> 在任何模式下都计失败。

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
1. `auto`：先尝试 `pandoc`；若 pandoc 失败（如默认 pdflatex 无法排版中文），自动用 `xelatex` 依次尝试 PingFang SC / Noto Sans CJK SC / Microsoft YaHei / SimSun 字体重试；若仍失败或无 pandoc，回退 Python 渲染引擎
2. `pandoc`：强制使用 pandoc（失败直接抛异常，不回退）
3. `python`：强制使用 Python 渲染（需安装 `markdown` + `reportlab`）

### 中文字体支持

**Python 引擎**：自动检测系统字体并注册：

| 系统 | 字体路径 |
|-----|---------|
| macOS | `/System/Library/Fonts/PingFang.ttc` |
| Linux | `/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf` |
| Windows | `C:/Windows/Fonts/msyh.ttc` |

如果系统字体无法注册，会自动回退到 ReportLab 内置的 CID 字体（如 `STSong-Light`）。

**pandoc 引擎**：默认使用 pdflatex，无法排版中文。脚本检测到失败后，若系统有 `xelatex`，会依次尝试常见 CJK 字体重试：
1. PingFang SC（macOS）
2. Noto Sans CJK SC（Linux）
3. Microsoft YaHei（Windows）
4. SimSun（Windows 通用）

所有字体均失败时，`auto` 模式回退到 Python 引擎；`pandoc` 模式抛出异常。

### 文本安全

Python 引擎在将文本传入 reportlab `Paragraph` 前统一调用 `xml.sax.saxutils.escape()` 转义 `<`、`>`、`&`，避免含特殊字符的文本（如 `A<B`、`C&D`）被误解析为 XML 标签导致崩溃。`Preformatted`（代码块/表格）不需要转义，reportlab 不对其解析 XML 标签。

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

**Python 引擎**：确保系统有支持的中文字体。脚本会依次尝试系统字体并回退到 ReportLab 内置 CID 字体（如 `STSong-Light`）。

**pandoc 引擎**：默认 pdflatex 无法排版中文。`auto` 模式下脚本会自动切换到 `xelatex` 并尝试 CJK 字体；若使用 `--engine pandoc` 且 pandoc 失败，请确认系统已安装 `xelatex`（TeX Live / MacTeX / MiKTeX），或改用 `--engine auto` 让脚本自动回退。

### Q: Markdown 含 `<` 或 `&` 等字符，转 PDF 时崩溃？

已修复。Python 引擎在传入 reportlab 前统一 `escape()`，含 `A<B`、`C&D` 等文本可正常输出 PDF。pandoc 引擎无此问题。

### Q: 批量转换时某些文件失败？

查看控制台输出的错误信息，常见原因：
- 文件损坏
- 密码保护
- 非标准 DOCX 格式

### Q: 文档更新后想重新转换，但输出已存在怎么办？

直接重跑批处理即可：转换成功时会写入 `.converted` 完成标记（记录源文件
SHA-256），源文件变更后哈希不一致会**自动清理重转**，无需任何参数。
仅当想强制重建全部输出时才需要 `--force`：
```bash
python scripts/batch_convert.py ./documents ./output --force
```

### Q: 带超大附件的正常文档转不出来怎么办？

例如文档里嵌了一个 60MB 的 Excel（默认上限 50MB）或一张超大图片，默认策略
会整篇拒绝。如果希望“丢附件保正文”，显式选择降级模式：
```bash
python scripts/convert_docx.py input.docx ./output --on-limit skip
python scripts/batch_convert.py ./documents ./output --on-limit skip
```
超限资源被跳过并在 Markdown 原位置留下 `*【图片已跳过：…】*` 可见说明，
转换结束输出跳过清单日志。注意：zip bomb 等恶意特征不受此开关影响，
依旧整篇拒绝；默认值为 `reject`，不会静默降级。

### Q: 想用上传时的原始文件名命名输出怎么办？

```bash
python scripts/convert_docx.py upload1234.docx ./output --output-name 用户原始文件名
```
或 Python API `convert_docx_to_markdown(path, out_dir, output_name="用户原始文件名")`。
命名统一经 `sanitize_stem` 清洗，输出目录、`.md` 文件名与 sentinel 的
`folder_name` 三处一致；sentinel 仍记录真实源文件 SHA-256。

### Q: 恶意/异常 DOCX 会把进程拖死吗？

不会。解压前会依据 ZIP 元数据做资源耗尽防线（总解压 500MB、单 entry 100MB、
压缩比 100x、图片数量/大小/像素上限等，见上文「安全防线」），实际读取时再按
真实解压量兜底。超限抛 `DocxSecurityError`——批处理计为失败并清理输出；
它继承 `ValueError` 以兼容既有处理，但调用方应精确捕获本类且**不可降级重试**。

### Q: 批处理时单个文档卡死怎么办？

用 `--timeout`（默认 300 秒）限制单文档转换时长，超时计为失败并清理该文档的
半成品输出。该机制基于 POSIX `signal.alarm`，Windows 上自动跳过（无超时保护）。

### Q: 需要额外安装 defusedxml 吗？

可选。安装后 DOCX 内 XML 解析启用实体膨胀/外部实体防御；未安装自动回退标准库，
功能不受影响。`pip install defusedxml` 即可启用。

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
