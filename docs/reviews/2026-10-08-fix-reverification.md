# 2026-10-08 二次修复复核

复核对象：验收报告 `2026-10-08-fix-acceptance.md`（14:46，结论未通过）之后，修复会话于 15:13 完成在工作区的未提交改动（转换器、测试、SKILL.md）。

背景：用户要求逐一修复验收报告的 6 项遗留发现。本会话核实工作区已含该二次修复（BUG-020/025/026 恢复修复、BUG-027/028 新增修复、TST-008 12 项新回归、DOC-009 文档同步）后，**未改动任何实现与测试**，改为按验收报告的独立复现步骤逐项构造输入复核（不使用 `tests/` 下的 helper），并重跑全量测试与静态检查。

**结论：6 项发现全部修复并通过独立复现。** 工作区 135 项测试 + 6 项子测试通过；干净快照（`git archive HEAD` + 工作区补丁）127 项通过、8 项跳过 + 6 项子测试；compileall、Ruff、`git diff --check` 通过；SKILL.md 新增表述与实现逐句核对一致。复核中新登记 1 个既有缺口（脚注体内图片不可见，[[BUG-029]]，HEAD 即存在，非本轮修复引入）。

## 逐项复核结果

| # | 验收发现 | 结果 | 独立复现证据 |
| --- | --- | --- | --- |
| 1 | [P1] 脚注关系与包根绝对目标绕过数量配额（BUG-026） | 通过 | 按报告步骤自建 DOCX：脚注场景（图片段落移入 `word/footnotes.xml`、主文档仅 `w:footnoteReference`、关系在 `footnotes.xml.rels` Target=`custom/imageN.png`、图片在 `word/custom/`）与包根场景（`custom/imageN.png` + Target=`/custom/imageN.png`）。配额=1 时 reject 均以"图片数量超过上限"整篇拒绝；skip 均仅落盘 1 张且正文保留。包根（正文引用）场景 Markdown 含 1 个图片引用 + 2 条 `*【图片已跳过：图片数量超过上限】*` 可见说明。 |
| 2 | [P2] `_same_file_content()` 吞 `TimeoutError`（BUG-025） | 通过 | 自建 `image1.png` + `image1.jpeg`（PNG 魔数、扩展名修正后同名、内容不同）触发比较阶段；`os.path.getsize` 替换为 `sleep(2)`，真实 `batch_convert(timeout=1)` → `{'success': 0, 'skipped': 0, 'failed': 1}`，输出目录整体清理（无 sentinel）。代码走读确认 `_same_file_content`（:478）与 `_TableHTMLParser._safe_int`（:726）均显式 `except TimeoutError: raise` 置于 OSError 之前。 |
| 3 | [P2] 原文反斜线+管道导致错列丢尾格（BUG-020） | 通过 | 自建 XLSX（H1/H2 表头，数据 `A\|B`/`A\\|B`/`A\\\|B` + TAIL），经 `excel_to_markdown` + `markdown.markdown(extensions=['tables','fenced_code'])` 渲染：1/2/3 条反斜线场景原文均完整保留在单格、TAIL 尾列不丢。序列化层（`_escape_plain_cell_text` 先转反斜线后转管道、`_escape_unescaped_pipes` 按奇偶补转义、`_serialize_table_cell` 区分原始/嵌套片段）走读核对。 |
| 4 | [P2] 字面标签 PDF 渲染仍成 HTML（BUG-020） | 通过 | 嵌套表格内"槽位 `&lt;time&gt;` 文本"：Markdown 源保留实体形式（`\x00/\x01` 标记保护，:2161～:2163）；用与 `md_to_pdf.py:163` 完全相同的渲染调用断言——可见文本含"槽位 `<time>` 文本"、渲染产物无 `<time>` 元素、外层尾部内容不丢。 |
| 5 | [P2] hash 候选再次占用时盲复用（BUG-027） | 通过 | 自建 2×2 tiny_png（摘要前缀 `aaec581e`，与验收报告一致），预置 `image1.png` 为不同内容后分别用错误文件/指向外部文件的符号链接/目录占用 `image1_aaec581e.png`：三种场景转换均成功，Markdown 引用 `image1_aaec581e_2.png`——普通文件、字节与输入一致、不跟随链接（外部文件未被创建）。提取循环（:1652）与 Mammoth 回调 fallback（:1932）共用 `_allocate_asset_path` 循环分配 + `_write_asset_file_exclusive` O_EXCL 写入。 |
| 6 | [P2] 原子写入强制 0644 扩大私有权限（BUG-028） | 通过 | CLI 子进程级实测：`umask 077` 下新转换的 `.md` 与 `.converted` 权限均 0600；预置 0600 后在 `umask 022` 下 `--force` 重转，权限保持 0600。`_atomic_target_mode`（:423）对已有普通文件保留原权限位，其余按调用者 umask 计算。 |

## External 关系语义实验

验收报告建议"正确处理……External 类型"。实验证实：3 个图片关系均标 `TargetMode="External"` 且 Target 指向包内真实条目时，**mammoth 无视 TargetMode 仍从 ZIP 读取并触发回调**（配额生效、落盘受限）。因此当前实现按"目标解析为真实 ZIP 条目即计数、不按 TargetMode 过滤"是**必要语义**——若按 External 跳过计数，该实验场景即重新打开配额绕过。真外部 URL 永远解析不到条目名，自然排除；`_document_image_part_names` 文档字符串"External 链接自动排除"按此理解成立。无需改动。

## 新发现（登记 BUG-029，待用户决策）

脚注场景独立复现时发现：mammoth 把脚注内容输出为文末 `<ol><li id="footnote-N">`，`_convert_footnotes`（:2022）提取脚注体时**剥离全部标签**，脚注内 `<img>`（assets 引用与 `__SKIPPED_IMAGE_*__` 占位）随之丢失。后果：

- 配额内的脚注图片提取到 `assets/` 但 Markdown 无任何引用（孤儿资源文件）；
- skip 模式下超限的脚注图片无可见跳过说明（与 DEV-006"超限图片留可见说明"的声明语义不符），仅剩悬空 `[^N]` 引用。

该行为在 HEAD 即存在（`_convert_footnotes` 同样剥离标签），属脚注渲染能力缺口而非本轮修复引入；BUG-026 的配额目标本身不受影响。修复需先决定脚注图片的渲染方式（如脚注体保留 `![](...)` 或跳过说明文本），会改变含脚注图片文档的输出，故登记待修复、待决策，未在本轮实施。

## 验证环境与限制

- 本机：135 passed、6 subtests passed（4.61s）；干净快照：127 passed、8 skipped（本机私有 DOCX 夹具）、6 subtests passed（4.08s）。
- compileall、Ruff、`git diff --check` 通过。
- 发现 1/2/3/4/5/6 的复现输入均由本会话按验收报告步骤独立构造（`/tmp/verify_acceptance_fixes.py` 及内联脚本），未依赖 `tests/` helper。
- 未执行 Windows 真机测试；未提交、未发布、未改变 GitHub issue 状态。
