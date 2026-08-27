# 任务跟踪列表

记录本项目所有任务：代码 bug、bug 转需求、新增需求、需求调整、功能开发、代码审查、测试数据、文档维护、配置运维等。

> 说明：本文件是当前项目的任务清单。所有新增事项、状态变更和完成记录都应同步写入本文件。
> 字段说明：动作字段只允许以下 8 个固定枚举：修复、开发、优化、调整、规划、检查、文档、运维。
> 时间说明：发现时间和完成时间分开记录，格式为 YYYY-MM-DD HH:MM，使用机器本地时区的 24 小时制时间；未完成事项的完成时间填 -。
> 状态说明：Bug 未完成用待修复，通用未完成用待办（或待开发），进行中/已完成/已修复/已关闭/已解决按语义选用；条目互引用 [[BUG-001]] 语法。
> 归并规则：审计、复核、核查、审查、验证、评估统一记为“检查”；重构、清理统一记为“优化”；方案、梳理统一记为“规划”；记录类文档事项统一记为“文档”。

## 代码 Bug

| ID | 动作 | 问题描述 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| BUG-001 | 修复 | 未知格式图片被误写成 `.png`，导致文件损坏（WMF/EMF/TIFF 等） | 2026-08-07 09:30 | 2026-08-07 10:05 | 已修复 | `detect_image_format()` 新增 TIFF/WMF/EMF 魔数；无法识别返回 `None`，zip 提取保留原扩展名；mammoth 回调按 `content_type` 推断，仍未知用 `.bin`。文件：`scripts/convert_docx.py`。另加表格占位符残留 `__TABLE_PLACEHOLDER_` 告警自检。 |
| BUG-002 | 修复 | Python PDF 引擎遇特定文本（如 `A<B`）直接崩溃 | 2026-08-07 09:30 | 2026-08-07 10:05 | 已修复 | reportlab `Paragraph` 将文本当 XML 解析；传入前统一 `escape()`。文件：`scripts/md_to_pdf.py`。含 `A<B` 的 md 已实测可出 PDF。 |
| BUG-003 | 修复 | pandoc 路径下中文 PDF 必失败（默认 pdflatex） | 2026-08-07 09:30 | 2026-08-07 10:05 | 已修复 | 失败时若有 xelatex，依次用 PingFang SC / Noto Sans CJK SC / Microsoft YaHei / SimSun 重试；`--engine auto` 下 pandoc 彻底失败回退 Python 引擎。文件：`scripts/md_to_pdf.py`。真机 pandoc 路径见 [[TST-002]]。 |
| BUG-004 | 修复 | 新增测试依赖被 `.gitignore` 排除的本机私有 DOCX，干净 clone 全量测试失败 | 2026-08-28 00:05 | 2026-08-28 00:20 | 已修复 | 以 `git archive HEAD` 应用工作区补丁复现 24 项失败；核心命名/安全/批处理测试改为动态构造最小 DOCX，真实业务文档回归仅在对应夹具存在时逐项运行，避免整类误跳过。 |
| BUG-005 | 修复 | `output_name="原始文件名.docx"` 产出 `.docx.md`，与 issue #3 的 Web 上传接口不兼容 | 2026-08-28 00:10 | 2026-08-28 00:18 | 已修复 | 仅对末尾 `.docx` 做大小写不敏感剥离；`需求V2.4` 等非 DOCX 点号后缀完整保留；新增两项回归测试。 |
| BUG-006 | 修复 | 批处理 sentinel 未记录 `on_limit`，先用 skip 转换后再请求 reject 会错误命中缓存 | 2026-08-28 00:10 | 2026-08-28 00:18 | 已修复 | `.converted` 增加 `on_limit`；批处理跳过同时比较源哈希与策略。V0.1.6 缺策略 JSON 按 `reject` 兼容读取，未知策略视为无效。 |
| BUG-007 | 修复 | `extract_content_from_docx` 返回值由 3 项变 4 项破坏下游解包契约；底层/批处理 API 未统一拒绝非法 `on_limit` | 2026-08-28 00:20 | 2026-08-28 00:24 | 已修复 | 保持历史三项返回值，以可选 `skip_state` dict 输出内部状态；新增公共 `validate_on_limit` 并在校验器、提取器、主入口与批处理入口最前调用。 |
| BUG-008 | 修复 | skip 模式图片数量超限后为每个剩余条目累计一条 skipped 记录，不符合“停止提取”语义且可造成额外 O(N) 内存占用 | 2026-08-28 00:25 | 2026-08-28 00:40 | 已修复 | 达到图片配额后只写一条汇总记录并立即退出 ZIP 媒体提取循环；Mammoth 回调继续按引用位置输出可见说明，但不再按 hash 重复增长审计清单。 |
| BUG-009 | 修复 | 单文件 API/CLI 复用同名输出目录时，本轮跳过的超限图片会遗留上次转换的 assets 文件 | 2026-08-28 00:30 | 2026-08-28 00:38 | 已修复 | Markdown 成功写入后、sentinel 写入前，仅清理本轮未引用的 assets 文件/符号链接；转换失败前不预删旧产物。 |
| BUG-010 | 修复 | 图片数量配额未约束 Mammoth 回调，且因大小/像素超限而跳过的图片不占配额，日志/内存/读取仍可 O(N) | 2026-08-28 00:30 | 2026-08-28 01:01 | 已修复 | ZIP 提取预先按物理媒体条目划定唯一白名单，超大/像素图也占名额；Mammoth 复用该白名单，配额外在打开流前短路，同一图片重复引用不重复占配额，避免两阶段各自分配导致 2x 落盘。 |
| BUG-011 | 修复 | 嵌入 XLSX 仅受 DOCX 外层 entry 大小限制，openpyxl 解析前未校验 XLSX 内层 ZIP bomb | 2026-08-28 00:30 | 2026-08-28 00:40 | 已修复 | `excel_to_markdown` 在 openpyxl 前对 XLSX 内层 ZIP 执行无条件恶意特征校验，`DocxSecurityError` 不再被普通 Excel 转换失败捕获并降级。 |
| BUG-012 | 修复 | 旧 assets 清理会误删 `create_subfolder=False` 平铺输出中其他文档资源，且预置 assets 符号链接可将清理导向外部目录 | 2026-08-28 00:48 | 2026-08-28 00:50 | 已修复 | 仅对文档独占子目录执行 stale assets 清理；平铺共享模式保留其他资源；转换前拒绝 assets 符号链接或非目录对象。 |
| BUG-013 | 修复 | ZIP 允许重复 entry 名，同名 `word/media` 可绕过基于路径 set 的图片数量配额 | 2026-08-28 00:55 | 2026-08-28 00:58 | 已修复 | 安全校验对所有 ZIP 条目名去重，发现同名条目无条件抛 `DocxSecurityError`，同时消除 XML/关系文件的不唯一解压语义。 |

## 调整事项

| ID | 动作 | 事项 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |

## 检查事项

| ID | 动作 | 事项 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| CHK-001 | 检查 | 复核 [[BUG-001]]～[[BUG-003]] 修复是否落地且正确 | 2026-08-07 10:18 | 2026-08-07 10:20 | 已完成 | 源码核对 + 针对性实验：魔数/扩展名、`A<B` PDF、pandoc mock（xelatex 重试与 auto 回退）均通过；27 个回归测试全绿。本机无 pandoc/xelatex，真机中文 PDF 未测。 |
| CHK-002 | 检查 | 系统审查 issue #2/#3 实现的安全边界、兼容性、缓存语义与干净环境可复现性 | 2026-08-28 00:05 | 2026-08-28 01:01 | 已完成 | 定位并修复 [[BUG-004]]～[[BUG-013]]；确认 ZIP bomb/重复 entry 不受 skip 影响、DOCX/XLSX 内外层读取有界、提取与 Mammoth 共享物理媒体配额且无旁路、重转无旧资源残留，共享 assets 不误删。设计与实施计划见 `docs/plans/2026-08-28-docx-limit-output-*.md`。 |

## 测试数据

| ID | 动作 | 事项 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| TST-001 | 检查 | 为 [[BUG-001]]～[[BUG-003]] 补回归单测（魔数识别、`A<B` PDF、pandoc mock） | 2026-08-07 10:49 | - | 待办 | 现有 27 测未覆盖这 3 个修复点；建议写入 `tests/`。 |
| TST-002 | 检查 | 本机安装 pandoc + xelatex 后实测中文 PDF 路径 | 2026-08-07 10:49 | - | 待办 | 验证 [[BUG-003]] 的 xelatex 字体重试与 `auto` 回退在真机可用。 |
| TST-003 | 检查 | 为 [[DEV-001]]～[[DEV-004]] 补专项单测并全量回归 | 2026-08-27 17:40 | 2026-08-27 17:58 | 已完成 | 新增 `tests/test_docx_security_and_batch.py` 共 37 测（zip 各上限/压缩比门槛、像素炸弹、sentinel 读写与旧格式、批处理跳过/源变更重转/失败清理/超时/handler 恢复、`sanitize_stem` hash 策略）；全套 64 测全绿（2.6s）。另做真机端到端：批处理→跳过→源变更自动重转、CLI 拒绝 500MB 炸弹（exit 2 且无输出残留）。 |
| TST-004 | 检查 | 为 [[DEV-005]]/[[DEV-006]] 与 [[BUG-004]]～[[BUG-013]] 补专项测试并全量回归 | 2026-08-27 19:10 | 2026-08-28 00:58 | 已完成 | 覆盖 output_name、on_limit 两层防线/API/CLI、策略 sentinel、历史契约、assets 清理/共享/符号链接、超限图片配额前置、同图多引用、重复 ZIP entry、XLSX 内层 ZIP bomb 及干净 clone。最终测试数与干净归档验证见本轮最终验收。 |

## 文档维护

| ID | 动作 | 事项 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| DOC-001 | 文档 | 同步 [[BUG-001]]～[[BUG-003]] 相关行为说明到 skill 文档 | 2026-08-07 09:30 | 2026-08-07 10:05 | 已完成 | 已更新 `SKILL.md`（TIFF/WMF/EMF、未知格式保留原扩展名）与 `references/usage-guide.md`（`detect_image_format` 说明）。；BUG-001 图片格式文档已同步；BUG-002/003 文档见 [[DOC-002]] |
| DOC-002 | 文档 | 补齐 BUG-002（escape）与 BUG-003（pandoc CJK 回退）的 SKILL.md / usage-guide.md 文档说明 | 2026-08-07 11:00 | 2026-08-07 11:00 | 已完成 | 已更新 SKILL.md 引擎策略说明、usage-guide.md 的 run_conversion/中文字体支持/文本安全/FAQ 章节；同步修正 requirements.txt 核心/可选区分与 README.md 依赖表 |
| DOC-003 | 文档 | README / SKILL.md 增加 V0.1.5 版本徽章 | 2026-08-07 11:13 | 2026-08-07 11:14 | 已完成 | 标题下仅增加 `version-0.1.5` shields.io 徽章（不含 build 号）。 |
| DOC-004 | 文档 | 同步 [[DEV-001]]～[[DEV-004]] 行为说明到 skill 文档 | 2026-08-27 17:40 | 2026-08-27 18:00 | 已完成 | SKILL.md 新增 Security/Batch Reliability 章节、批处理跳过语义与 defusedxml 依赖说明；usage-guide.md 更新 convert_docx 异常与新函数、batch_convert 参数/特性、FAQ（自动重转/恶意 DOCX/超时/defusedxml）；README 功能与注意事项同步；requirements.txt 增加 defusedxml 可选依赖；版本徽章升至 0.1.6。 |
| DOC-005 | 文档 | 同步 [[DEV-005]]/[[DEV-006]] 行为说明并准备 V0.1.7 | 2026-08-27 19:10 | 2026-08-28 00:24 | 已完成 | README / SKILL.md / usage-guide.md 已同步两层防线、on_limit、output_name、`.docx` 后缀、策略 sentinel 与旧版兼容语义；版本徽章升至 0.1.7。当前改动尚未提交或发布。 |

## 功能开发

| ID | 动作 | 事项 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| DEV-001 | 开发 | 恶意 DOCX 资源耗尽防线（issue #1，下游生产验证回传） | 2026-08-27 17:30 | 2026-08-27 17:55 | 已完成 | `convert_docx.py` 新增 `DocxSecurityError(ValueError)`（安全拒绝不可降级重试）+ `validate_docx_zip_security`（解压前按 ZIP 元数据校验：总解压 500MB/单 entry 100MB/压缩比 100x×2/图片数 500/单图 20MB/嵌入 Excel 50MB，阈值集中在 `DOCX_SECURITY_LIMITS`）+ `image_pixel_count` 头部像素检测（PNG/JPEG/GIF/BMP/WEBP/TIFF，5000 万上限）+ `read_zip_entry_bounded` 真实解压兜底；XML 解析统一走 defusedxml（未装回退标准库）。 |
| DEV-002 | 开发 | SHA-256 完成标记（`.converted` sentinel，issue #1） | 2026-08-27 17:30 | 2026-08-27 17:55 | 已完成 | 转换成功后原子写（tmp+rename）JSON sentinel `{"folder_name", "source_sha256"}`；批处理跳过需目录+md+有效 sentinel 且哈希与当前源一致；旧格式纯文本 sentinel 视为无效重转；源变更自动重转。配套 `sanitize_stem` 在字符替换/引号删除/超长截断时附加源名短 hash（NFKC 归一化不加 hash，中文场景过普遍，其碰撞由 sentinel 兜底——对 issue 描述做的工程化取舍）。 |
| DEV-003 | 开发 | 批处理单文档超时 `--timeout`（issue #1） | 2026-08-27 17:30 | 2026-08-27 17:55 | 已完成 | `batch_convert.py` 新增 `--timeout`（默认 300s，<=0 不限制），POSIX `signal.alarm` 实现；Windows 无 SIGALRM 自动跳过、非主线程降级为无超时；finally 恢复原 handler。 |
| DEV-004 | 开发 | 半成品目录清理（issue #1） | 2026-08-27 17:30 | 2026-08-27 17:55 | 已完成 | 转换失败（含超时/安全拒绝）清理输出目录；非 force 路径既有输出不可信（sentinel 缺失/无效/哈希不匹配）同样清理重转。安全拒绝单独记日志，计失败不重试。 |
| DEV-005 | 开发 | `output_name` 自定义输出命名（issue #3） | 2026-08-27 19:10 | 2026-08-28 00:18 | 已完成 | `convert_docx_to_markdown` 新增 `output_name=None`，统一用于输出目录、`.md` 与 sentinel；CLI 增加 `--output-name`；末尾 `.docx` 自动去除、版本号点号保留、空白值拒绝，默认 None 行为不变。 |
| DEV-006 | 开发 | 资源超限 `on_limit` 降级跳过策略（issue #2） | 2026-08-27 19:10 | 2026-08-28 00:24 | 已完成 | 防线两层：恶意特征无条件拒绝，可降级资源默认 reject、显式 skip；所有读取有界且 Mammoth 无旁路；超限图片留可见说明与审计日志。Python API / 单文件 CLI / 批处理 CLI 默认一致；策略写入 sentinel，变化时自动重转。 |

## 配置运维

| ID | 动作 | 事项 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |

## 规划事项

| ID | 动作 | 事项 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| PLN-001 | 规划 | 旧版 `.doc` / 嵌入 `.xls`（非 xlsx）转换策略 | 2026-08-07 10:49 | - | 待办 | openpyxl 读不了；可选 LibreOffice `soffice --convert-to` 先转格式，或 `olefile` 枚举 OLE 流。遇到再实现。 |
| PLN-002 | 规划 | 嵌入对象无预览图（`w:object` 无 imagedata）时的替换策略 | 2026-08-07 10:49 | - | 待办 | 当前 hash 匹配会漏；可按文档顺序在 `<w:object>` 出现位置直接插队替换。 |
| PLN-003 | 规划 | OMML 公式完整 LaTeX 输出 | 2026-08-07 10:49 | - | 待办 | 现仅提取文本节点并以 `$$ ... $$` 包裹；真要 LaTeX 可交叉用 pandoc 的 OMML→TeX。 |

## 优化事项

| ID | 动作 | 事项 | 发现时间 | 完成时间 | 状态 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| OPT-001 | 优化 | 重要文档与 pandoc/markitdown 交叉 diff 校验流程 | 2026-08-07 10:49 | - | 待办 | 对重要文档顺手跑 `pandoc input.docx -o ref.md --extract-media=...` 或 markitdown，diff 两份输出以暴露转换器盲区；成本低。 |
