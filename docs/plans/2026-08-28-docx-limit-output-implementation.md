# DOCX Limit Policy and Output Naming Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 安全完成 `on_limit`、`output_name` 及其 CLI/批处理集成，并修复审查发现的干净环境、命名后缀和 sentinel 策略复用问题。

**Architecture:** ZIP 级恶意特征与可降级资源限制分层；转换主入口向提取器和 Mammoth 回调传递统一策略。输出命名统一规范化，批处理 sentinel 把影响产物的策略纳入缓存键。

**Tech Stack:** Python 3、zipfile、mammoth、openpyxl、unittest/pytest、ruff。

---

### Task 1: 让测试在干净 clone 中自包含

**Files:**
- Modify: `tests/test_docx_output_name.py`
- Modify: `tests/test_docx_security_and_batch.py`
- Modify: `tests/test_docx_excel_tables.py`
- Modify: `tests/test_docx_markdown_quality.py`

**Steps:**
1. 用干净归档复现私有 DOCX 缺失失败。
2. 将核心契约测试改为动态生成最小 DOCX。
3. 为只适用于真实内部文档的回归类增加夹具存在性跳过条件。
4. 在干净归档中重跑全量测试，确认不再因私有文件失败。

### Task 2: 修正 output_name 后缀语义

**Files:**
- Modify: `tests/test_docx_output_name.py`
- Modify: `skills/docx-to-markdown/scripts/convert_docx.py`

**Steps:**
1. 新增失败测试：`原始名称.docx` 输出为 `原始名称.md`。
2. 新增保护测试：`需求V2.4` 保留 `.4`。
3. 仅移除大小写不敏感的末尾 `.docx`，再调用 `sanitize_stem()`。
4. 运行专项测试确认通过。

### Task 3: 将 on_limit 纳入批处理完成判定

**Files:**
- Modify: `tests/test_docx_on_limit.py`
- Modify: `tests/test_docx_security_and_batch.py`
- Modify: `skills/docx-to-markdown/scripts/convert_docx.py`
- Modify: `skills/docx-to-markdown/scripts/batch_convert.py`

**Steps:**
1. 新增失败测试：先以 `skip` 转换，再以 `reject` 批处理不得命中缓存。
2. sentinel 写入 `on_limit`；旧 sentinel 缺字段按 `reject` 读取。
3. `_is_output_complete` 比较请求策略，批处理透传该值。
4. 验证旧 sentinel、同策略跳过和跨策略重转三类行为。

### Task 4: 收紧公开参数校验

**Files:**
- Modify: `tests/test_docx_on_limit.py`
- Modify: `skills/docx-to-markdown/scripts/convert_docx.py`
- Modify: `skills/docx-to-markdown/scripts/batch_convert.py`

**Steps:**
1. 新增失败测试覆盖 validator、extractor 和 batch API 的非法策略。
2. 抽取轻量校验函数并在各公开入口最前执行。
3. 运行专项测试和全量回归。

### Task 5: 文档、台账与最终验证

**Files:**
- Modify: `README.md`
- Modify: `skills/docx-to-markdown/SKILL.md`
- Modify: `skills/docx-to-markdown/references/usage-guide.md`
- Modify: `task-list.md`

**Steps:**
1. 同步 `.docx` 后缀、sentinel 策略字段及旧格式兼容说明。
2. 修正台账中未发生的 commit/发布表述，新增本轮审查修复记录。
3. 运行 `python3 -m pytest -q`、`ruff check .`、`git diff --check`。
4. 在干净归档中应用最终补丁并重跑 pytest，确认仓库外夹具不再是隐式依赖。
