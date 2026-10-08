# 2026-10-08 修复验收

验收对象：HEAD `e9aff1c` 上当前未提交的转换器、测试和 SKILL.md 改动，针对上一轮 BUG-019～BUG-026。全部实验只使用临时 DOCX、XLSX 和输出目录。本轮未修改实现与测试，未提交、发布或改变 GitHub issue 状态。

**结论：验收未通过，仍需继续修复。** 原报告的基础复现已修复，当前 123 项测试和 2 项子测试通过，但扩展验收复现了资源计数、超时、转义、资源文件复用与权限问题。不能据此认定所有遗留缺口已经修复。

## 原修复逐项结果

| 条目 | 验收结果 |
| --- | --- |
| BUG-019：文档输出子目录链接越界 | 通过原静态链接场景；普通目录校验在写入前拒绝输出子目录链接 |
| BUG-023：Markdown 文件链接覆盖外部文件 | 通过；随机独占临时文件和 os.replace 不跟随最终 Markdown 链接 |
| BUG-024：固定 sentinel 临时名称链接 | 通过；固定 .converted.tmp 不再用于写入，外部目标不变 |
| BUG-021：图片 src 等号后空白 | 通过；空格、制表符、换行及链接内图片均有回归覆盖 |
| BUG-022：符号链接测试权限不足 | 既有能力探测保留正确；未执行 Windows 真机测试 |
| BUG-025：Excel 超时传播 | 原 Excel 场景通过，但新增图片比较助手仍吞超时，整体未通过 |
| BUG-026：非标准媒体位置 | 主文档相对 word/custom 目标场景通过；脚注及包根绝对目标仍绕过数量配额，整体未通过 |
| BUG-020：嵌套内容和转义 | 原普通内表、实体、单个管道和三层嵌套场景通过；原文本反斜线与管道组合仍错列，字面标签在 PDF 渲染时仍成为 HTML |

## 必须继续修复的发现

### 1. [P1] 资源集合遗漏脚注关系和包根绝对目标（BUG-026）

位置：`skills/docx-to-markdown/scripts/convert_docx.py:123`～`:142`、`:1127`～`:1136`、`:1734`～`:1740`。

`_document_image_part_names()` 只读主文档 document.xml.rels，未读 Mammoth 同样处理的 footnotes.xml.rels/endnotes.xml.rels。`resolve_part_path()` 对 `/custom/image1.png` 去掉开头斜线后再加 word，得到 word/custom/image1.png；Mammoth 实际读取的是包根 custom/image1.png。

独立复现：

1. 用 test_docx_on_limit.py 的 build_docx_with_images 创建 3 张不同 tiny_png。
2. 脚注场景：将图片段落移动到 word/footnotes.xml，主文档只留下 w:footnoteReference，图片关系移到 word/_rels/footnotes.xml.rels，Target 指向 custom/imageN.png，图片存放 word/custom/。
3. 包根场景：图片条目改为 custom/imageN.png，主文档 image 关系 Target 改为 `/custom/imageN.png`。
4. 分别设置 image_count=1，调用 reject 与 skip。

两种结构、两种策略的四次转换全部成功，均落盘 3 张图片，没有数量跳过说明。reject 回调的大小/像素检查已正确，但没有独立数量检查；skip 对未识别路径仅在已知集合超额时拒绝，遗漏资源时该标记为 False。

建议按照每个实际处理的 part 的路径解析其关系目标，正确处理包根绝对目标、相对目标及 External 类型；ZIP 统计、提取和回调统一使用完整集合。补脚注/尾注和包根目标的 reject/skip 数量回归，不只检查主文档相对 custom 目标。

### 2. [P2] 新增图片比较助手仍吞掉 TimeoutError（BUG-025）

位置：`convert_docx.py:436`～`:444`。

`_same_file_content()` 的 except OSError 会捕获其子类 TimeoutError。现有 Excel 超时补丁没有覆盖这个新增捕获点。

独立端到端复现：在真实转换开始前放置已有 assets/image1.png，以触发比较；仅将 getsize 调用替换为 sleep(2)，执行真实 batch_convert(timeout=1)。SIGALRM 被助手捕获后，批处理仍返回 `{'success': 1, 'skipped': 0, 'failed': 0}`，并写入 .converted。耗时约 1.12 秒，计成功说明 alarm 被降级吞掉，而非按超时计失败和清理。

建议在 OSError 前显式重抛 TimeoutError；审查所有执行在超时包裹内的异常捕获，避免新增助手再次破坏取消信号。补已有资源比较阶段的真实批处理超时回归。

### 3. [P2] 管道保护仅检查前一字符，原始反斜线文本导致错列和丢尾格（BUG-020）

位置：`convert_docx.py:555`～`:570`，尤其 `(?<!\\)\|`。

这个函数同时接收原始 Excel/Word 文本和已生成的嵌套 Markdown，却把管道前任何反斜线都当作已有正确转义。它没有区分原文中的反斜线，也没有检查连续反斜线数量。

独立复现：有效 XLSX 两列表，表头 H1/H2，数据第一格为 Python 原始字符串 `r'A\\|B'`（A、两条反斜线、管道、B），第二格为 TAIL。HEAD 的 Python Markdown tables 渲染仍保留 TAIL；当前输出第一格为 A 加反斜线、第二格变为 B，TAIL 完全丢失。同样的文本位于嵌套 Word 表格时也可复现。

建议区分原始单元格文本和已转义嵌套产物，在统一的最终序列化层处理反斜线/管道。增加 1、2、3 条反斜线相邻管道的真实 XLSX/DOCX 渲染断言，检查所有单元格文字与末尾列。

### 4. [P2] 找回的字面标签在实际 PDF 渲染时仍被当作 HTML（同 BUG-020）

位置：`convert_docx.py:2024`；实际项目渲染器 `scripts/md_to_pdf.py:163`。

当前实体处理正确地让 Markdown 源文件包含“保留 <table> 标签”或“保留 <time> 标签”，但末尾统一 unescape 将其变成未转义的 Markdown HTML。Python Markdown 渲染保留实际 `<table>`/`<time>` 元素，而不是输出 `&lt;table&gt;`/`&lt;time&gt;` 文字：槽位文字在渲染时不可见，table 还会形成不正确的嵌套元素。

现有字面标签测试只断言 Markdown 源字符串，渲染测试只覆盖 A|B，没有交叉验证字面标签。建议在最终 Markdown 保留字面标签所需的实体或 Markdown 转义，同时区分真正用来换行的 br；渲染后提取可见文本并断言完整槽位内容。

### 5. [P2] 图片 hash 改名后的候选路径未再次验证（BUG-027）

位置：`convert_docx.py:1512`～`:1523`；回调 fallback `:1798`～`:1804` 也有相同的盲复用假设。

图片自然名冲突时改成 image1_<hash8>.png，但对新名字仅检查 lexists，没有重新检查是否链接、普通文件及内容一致，随后直接写入 image_by_hash。

独立复现：使用 2×2 tiny_png（摘要前缀 aaec581e），预置 image1.png 为不同内容，再分别让 image1_aaec581e.png 是错误普通文件、指向外部的链接、目录。三次转换均成功，Markdown 均引用该候选；当前图片字节未正确落盘，链接场景还保留指向外部的资源链接。此处没有重新发生外部写入，但输出已经错误，不能判定资源写入与结果完整性通过。

建议循环分配候选，只有普通文件且字节一致才复用；其他占用继续附加序号或使用安全替换。写入也应使用独占创建避免检查和写入之间被置换。补次级候选的链接、悬空链接、目录和不同内容测试，断言输出引用最终对应输入图片的普通文件。

### 6. [P2] 原子写入强制 0644，扩大调用者指定的私有权限（BUG-028）

位置：`convert_docx.py:426` 和 `:471`。

两个原子写入函数均无条件 chmod(0644)。独立复现：在临时作用域设置 umask(0077)，调用 Markdown 与 sentinel 写入，最终两文件权限都为 0644，而非旧普通写入遵守 umask 时的 0600。覆盖已有 0600 的文件也会把其权限扩大为 0644。

建议保留已有文件的权限，或对新文件遵守调用者 umask/保持 mkstemp 的私有权限；不要为所有输出硬编码组用户和其他用户可读。补严格 umask 与覆盖私有旧文件的测试。

## 验证证据与限制

- 当前工作区：123 passed、2 subtests passed。
- 不含本机真实私有 DOCX 的临时干净快照：115 passed、8 skipped、2 subtests passed。
- compileall、Ruff、git diff --check、台账结构校验通过。
- 原有专项测试覆盖的目录/Markdown/sentinel 链接、Excel 超时、主文档相对 custom 媒体、src 空白和普通嵌套场景通过。
- 上述失败由独立的临时输入与定向延迟注入复现；未修改任何生产输入、已有输出或安装的 skills。
- 6 份真实 DOCX 已独立比较 HEAD/当前输出：5 份 Markdown 逐字节一致；设备预约文档仅多出被旧实现吞掉的字面标签，对当前输出去除这些字面标签后与 HEAD 完全一致。该对比验证的是 Markdown 源文件，不能替代前述 PDF 可见文本检查。
- 未进行 Windows 真机测试。

建议先补资源关系覆盖和剩余超时传播，再处理表格最终序列化、资源候选循环与权限。SKILL.md 中“all image relationships / every degradation handler”等保证应在实现与相应专项回归全部通过后再作为验收结论。
