# convert_docx 标题与表格还原修复计划

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 修复 DOCX 转 Markdown 中“编号+加粗标题未转为 Markdown 标题”与“多行/合并单元格表格失真”问题，并通过回归样例验证。

**Architecture:** 采用“先测后改”方式。先补充失败测试锁定当前问题，再在 `convert_docx.py` 中引入更稳健的表格解析与 Markdown 后处理，最后对 `tests` 中目标文档做回归转换和质量校验。

**Tech Stack:** Python 3、mammoth、openpyxl、pytest、标准库 `html.parser`。

---

### Task 1: 测试补齐（RED）

**Files:**
- Create: `tests/test_docx_markdown_quality.py`
- Modify: `tests/test_docx_excel_tables.py`（如需要共享加载工具）
- Test: `pytest -q tests/test_docx_markdown_quality.py`

1. 新增失败测试：`编号+加粗标题` 应转为 `##/###...` 标题。
2. 新增失败测试：带 `rowspan`、多段文本（多 `<p>`）的 HTML 表格应输出列数一致，单元格换行用 `<br>`。
3. 新增失败测试：回归样例（分布式唤醒、设备预约）转换后不应存在管道表断行导致的列数不一致。
4. 运行目标测试并确认失败（记录失败点）。

### Task 2: 实现修复（GREEN）

**Files:**
- Modify: `skills/docx-to-markdown/scripts/convert_docx.py`
- Test: `pytest -q tests/test_docx_markdown_quality.py`

1. 在 `excel_to_markdown` 中增加单元格文本规范化（转义 `|`，换行转 `<br>`）。
2. 增加 HTML 表格专用解析逻辑（`html.parser`）：
   - 支持 `rowspan/colspan` 展开；
   - 合并多段文本为单元格内容；
   - 输出稳定的 Markdown 管道表。
3. 增加 Markdown 后处理：将“编号+加粗标题”提升为 Markdown 标题层级。
4. 运行新增测试，直至全部通过。

### Task 3: 回归验证（VERIFY）

**Files:**
- Output: `tests/output/*/*.md`
- Test: `pytest -q`

1. 重新执行批量转换：`python3 skills/docx-to-markdown/scripts/batch_convert.py tests tests/output`
2. 检查 3 份输出文档：
   - 标题层级是否明显改善；
   - 表格是否无断行错列；
   - 图片/表格占位替换是否正常。
3. 运行相关测试集并确认通过。

### Task 4: 结果汇报

**Files:**
- Reference: `tests/output/*/*.md`

1. 汇总修复点与验证证据（命令与结果）。
2. 标注剩余风险（若有）。
3. 给出可选下一步优化项。
