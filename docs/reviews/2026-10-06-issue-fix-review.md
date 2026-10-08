# GitHub issue 与对应修复复核

审查日期：2026-10-06（Asia/Shanghai）。审查基线：`e9aff1cdfef697d85e82b248fd908f1c0b51ff99`，与 GitHub main 最新提交一致；同时审查已有未提交的 #5、#6、#7 修复。行号均指审查时工作区版本。

结论：仍需要继续修复。#4 尚未修复，且有同类文件链接写入缺口；#1、#2 的已关闭实现存在资源防线/超时遗漏；#5、#6 的本地修复仅解决基础场景；#7 本地修复正确。本轮只审查、验证和同步台账，没有修改转换器或测试，也没有评论、关闭或重开 GitHub issue。

## 逐 issue 结论

| Issue | GitHub 状态 | 实现与验证 | 是否需要继续修复 |
| --- | --- | --- | --- |
| [#1 安全防线、sentinel 与超时](https://github.com/wangminle/skills-doc-to-md/issues/1) | CLOSED | 常规 ZIP 防线、源哈希重转、失败清理已实现；图片实际关系目标绕过资源检查，Excel 捕获吞掉超时 | 是，建议重新跟踪遗漏 |
| [#2 资源超限降级](https://github.com/wangminle/skills-doc-to-md/issues/2) | CLOSED | 标准媒体目录的大小/像素/数量及 Excel 降级、恶意 ZIP 无条件拒绝正确；其他媒体位置绕过数量配额 | 是 |
| [#3 自定义输出命名](https://github.com/wangminle/skills-doc-to-md/issues/3) | CLOSED | API、单文件 CLI、末尾 .docx 剥离、版本号保留和 sentinel 命名通过核查 | 未发现需要继续修复的功能问题 |
| [#4 输出子目录符号链接越界](https://github.com/wangminle/skills-doc-to-md/issues/4) | OPEN | HEAD 与当前工作区均可删除外部 assets 文件；输出 Markdown 和 sentinel 临时文件也有链接写入缺口 | 是，优先修复 |
| [#5 嵌套表格](https://github.com/wangminle/skills-doc-to-md/issues/5) | OPEN | 普通两层嵌套已在工作区修复；字面标签文本导致整块内层内容丢失，重复转义还影响项目 PDF 路径 | 是 |
| [#6 超链接图片](https://github.com/wangminle/skills-doc-to-md/issues/6) | OPEN | 真实 DOCX 验证基础场景修复正确；新正则不接受等号后空格，原本可转换的普通图片也消失 | 是 |
| [#7 Windows 符号链接测试](https://github.com/wangminle/skills-doc-to-md/issues/7) | OPEN | 本地修复在 OSError/NotImplementedError 时正确 skip，正常创建链接时仍执行安全断言 | 不需要追加代码修复，尚未提交发布；Windows 真机未运行 |

#3 的关闭评论误称批处理也提供 `--output-name`；当前实际是单文件 CLI 提供此选项，批处理内部使用 API 做命名消歧。该差异为评论说明问题，不属于原 issue 必须实现的功能。

## 需要修复的具体问题

### 1. [P1] 输出子目录链接导致外部删除和写入（BUG-019，#4）

位置：`skills/docx-to-markdown/scripts/convert_docx.py:1533`、`:1712`。

当前只检查 assets 本层链接，`os.makedirs(final_output_dir, exist_ok=True)` 接受指向外部的文档子目录链接。临时实验中创建 `output/evil -> victim`，victim/assets/keep.bin 预先写入 KEEP，转换普通 evil.docx 成功，keep.bin 被删除，victim/evil.md 和 victim/.converted 被创建。

建议在任何输出写入之前拒绝文档子目录符号链接，并校验最终 realpath 位于指定输出根目录内；回归断言拒绝后外部目录内容完全不变。

### 2. [P1] 输出文件链接仍可覆盖外部内容（BUG-023、BUG-024，#4 同类遗漏）

位置：`convert_docx.py:1706` 和 `:379`。

即使文档输出目录是普通目录，预置 evil.md 为指向外部 keep.txt 的链接，`open(..., 'w')` 仍把外部 KEEP 覆盖为正文。预置 .converted.tmp 为外部文件链接时，写 sentinel 会把外部内容覆盖为 JSON，并把链接重命名为 .converted。

只修复目录链接不能堵住这两条路径。建议 Markdown 使用安全的临时文件与原子替换；sentinel 使用同目录独占创建的随机临时文件，避免跟随已有固定名称链接。增加两个外部文件零改动的独立回归用例。

### 3. [P1] 非标准媒体位置绕过资源防线（BUG-026，#1/#2）

位置：`convert_docx.py:142`、`:448`、`:1278`、`:1578`、`:1597`。

资源统计、白名单和 Mammoth 媒体路径识别均只匹配 word/media/。但 DOCX 图片可以通过关系指向包内其他位置。实际复现：使用 tests/test_docx_on_limit.py 的 build_docx_with_images，重写 ZIP，把 word/media/ 改为 word/custom/，并将 document.xml.rels 中 Target="media/ 改为 Target="custom/。

- reject 模式：20000×15000（3 亿像素）的 PNG 头正常转换并落盘，未触发默认 5000 万像素上限。
- skip 模式：将 image_count 设为 1，3 张不同图片全部落盘，没有跳过说明。

reject 回调直接无界 read，skip 对未知路径仅在已识别的 word/media 超额时短路。建议按实际 relationship 目标识别资源，并让两种策略的回调都执行大小、像素和统一物理图片数量限制。

### 4. [P2] Excel 降级捕获吞掉批处理超时（BUG-025，#1）

位置：`convert_docx.py:1219`；超时抛出与失败清理在 `batch_convert.py:130`、`:233`。

有效 DOCX 内嵌有效 XLSX，mock openpyxl.load_workbook 为 time.sleep(2)，batch_convert(timeout=1) 的 SIGALRM 在 Excel 解析期间触发。日志记录“Excel转Markdown失败: 单文档转换超时（>1秒）”，但 except Exception 将其降级为 None，批处理返回 success=1、failed=0，Markdown 和 sentinel 均存在。一次性 alarm 已触发，后续处理也失去该超时保护。

建议 TimeoutError 显式重抛，并检查其他降级捕获能否吞掉取消/超时信号；增加在真实转换内部触发超时后计失败、清理输出、不写 sentinel 的用例。

### 5. [P2] 内层表格文本被重新当作 HTML，整块内容丢失（BUG-020，#5）

位置：`convert_docx.py:576`、`:543`。

HTMLParser 默认解码实体；handle_data 在收集内表 HTML 时直接追加已解码文本，再递归解析。真实 DOCX 内表第一格写“保留 <table> 标签”，第二格写“内层Y”，Mammoth 正确输出 `&lt;table&gt;`，本地转换结果却将整个外层对应单元格变为空，第一格文字、内层Y 均丢失。

建议重新序列化文本节点时做 HTML 转义，或递归遍历结构避免多次解析；增加字面标签、实体、多层嵌套及前后文本保留的用例。

### 6. [P2] 嵌套 Markdown 重复转义与项目 PDF 渲染器不兼容（同 BUG-020）

位置：`convert_docx.py:460`、`:554`；项目使用的渲染器在 `scripts/md_to_pdf.py:163`。

内层单元格含 A|B，或存在三层嵌套时，内表已有转义再次经过外层规范化，产生连续反斜线。用项目自身的 `markdown.markdown(..., extensions=['tables', 'fenced_code'])` 渲染，第二行外层B2 不再成为表格第二格，尾部内容落到表格外或丢失；一行外表甚至不再识别为表格。MarkdownIt 的 table 渲染仍保留 tail，因此此结论明确限定为项目 Python Markdown/PDF 路径。

建议统一嵌套表格的降级文本表示与转义层次，并使用真实渲染器断言单元格与尾部文字。当前 count_malformed_table_blocks 只看前一个字符是否为反斜线，无法发现该兼容性问题。

### 7. [P2] 新图片正则导致空格属性写法的图片消失（BUG-021，#6）

位置：`convert_docx.py:792`。

提取共享正则时去掉了旧实现中等号后的 `\s*`。输入 `<p><img src= "a.png" /></p>`：HEAD 输出 `![](a.png)`，工作区输出空字符串；包裹链接时输出 `[](https://example.com)`。新加的用例只覆盖 src="..."，没有覆盖该回归。

建议恢复等号后空白支持，覆盖单引号、双引号、无引号、有空白属性以及链接内外图片。

## 验证结果与边界

- 当前完整工作区：`python3 -m pytest -q` → 108 passed、2 subtests passed。
- 不含本机私有 DOCX 的临时快照，应用当前转换器和测试：100 passed、8 skipped、2 subtests passed。跳过项是可选真实 DOCX 夹具。
- 临时 HEAD 快照只应用新增测试：#5、#6 两项失败；应用当前转换器后全量通过，说明新增测试能够识别基础修复，但覆盖范围不足。
- 真实 python-docx 生成的嵌套表格和超链接图片已做端到端转换；实体内容丢失也在真实 DOCX 复现。
- #7 用 OSError（模拟 Windows 缺少创建权限）和 NotImplementedError 分别运行原测试：均 1 skipped、0 errors、0 failures；未执行 Windows 真机测试。
- `python3 -m compileall -q skills tests`、Ruff、`git diff --check` 均通过。
- 所有破坏性验证仅对新建临时目录和测试文件执行，未碰触用户现有输出或安装的 skills。

建议顺序：先修 #4 的完整输出写入边界和 #1/#2 的资源防线旁路，再修超时遗漏及 #5/#6 的回归；补充专项测试后重新复核并提交发布。#7 当前实现可以保留。
