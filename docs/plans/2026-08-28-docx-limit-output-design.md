# DOCX 资源降级与输出命名设计

## 目标

在不削弱 ZIP bomb 防线的前提下，完成 issue #2 的可选资源降级策略和 issue #3 的自定义输出命名，并保证 Python API、单文件 CLI、批处理 CLI 的行为一致。

## 接口

- `convert_docx_to_markdown(..., output_name=None, on_limit="reject")`
- 单文件 CLI：`--output-name`、`--on-limit reject|skip`
- 批处理 API/CLI：`on_limit`、`--on-limit reject|skip`
- 默认 `reject`，保持 V0.1.6 行为。

`output_name` 接受名称或以 `.docx` 结尾的原始上传文件名；只移除末尾 `.docx`，避免把 `需求V2.4` 中的 `.4` 误当扩展名。结果统一经过 `sanitize_stem()`，用于输出目录、Markdown 文件和 sentinel。

## 安全分层

无条件拒绝：总解压量、单 entry 解压量、单 entry 压缩比、总压缩比。任何 `on_limit` 模式都抛 `DocxSecurityError`。

可降级限制：图片数量、单图文件大小、单图像素、嵌入 Excel 大小。默认抛 `ResourceLimitExceeded`；`skip` 模式只跳过具体资源，并输出告警和可见占位。

所有实际读取继续有字节上限。Mammoth 回调复用图片大小、像素和数量配额，不能把提取阶段跳过的资源重新写回磁盘。

## 完成标记

`.converted` 除源 SHA-256 和目录名外记录 `on_limit`。V0.1.6 旧 sentinel 缺少该字段时按 `reject` 解释。批处理只有在源哈希、目录名和请求策略都一致时才跳过，避免 `skip` 产物被后续 `reject` 错误复用。

## 测试与发布

测试必须能在干净 clone 中运行。真实内部 DOCX 仅作为可选本地回归夹具；缺失时跳过对应测试。核心安全、命名、CLI 和批处理契约全部使用测试内动态生成的最小 DOCX。

验收包括：专项测试、全量 pytest、ruff、`git diff --check`，以及在 `git archive HEAD + 暂存补丁` 构造的干净目录中再次跑全量测试。
