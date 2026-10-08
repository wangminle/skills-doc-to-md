# 修复轮再复核（BUG-039/040～048、TST-013～015、DOC-016）

2026-10-08 20:10（Asia/Shanghai）。用户要求：上一轮登记的问题「都修复好了，再检查一遍」。本轮为**独立再验证**——对 19:17 / 19:29 两轮修复逐项重新构造输入复现，并检查修复本身是否引入新缺陷。**本轮只做检查与台账同步，未修改实现、测试与文档**（新增本报告与 `task-list.md` 条目），未提交发布，未操作 GitHub 状态。

探针脚本全部位于 `/tmp`（`/tmp/rev/v3/`、`/tmp/rev/e2e/`、`/tmp/sec/`），未写入仓库。

---

## 一、结论速览

| 项目 | 结论 |
| --- | --- |
| 全量测试 | `python3 -m unittest discover -s tests` → **Ran 176 tests, OK**（退出码 0） |
| 台账 13 项修复 | **全部复现通过**（BUG-039、BUG-040～048、TST-013～015、DOC-016） |
| 真实文档端到端 | 6 份真实 DOCX，工作区与 HEAD 各 25 个产物；仅 `设备预约` 的 md 不同，**可见文本 0 处丢失、0 处改写、92 处找回**，assets 逐字节一致 |
| 超时纪律回归 | 全文 40 处 `except` 逐条核对，无新增吞 `TimeoutError` 路径 |
| 台账完整性 | 113 行、ID 无重复、动作/状态取值合法、56 个 `[[ID]]` 引用全部可解析、无「已完成但完成时间为空」 |
| 本轮新发现 | **1 个 bug（[[BUG-049]]，P3，fail-closed 类）** + 1 处语义差异（登记 [[OPT-002]]） |

---

## 二、逐项再验证（均由本人独立复现）

| 台账条目 | 验证方式 | 结果 |
| --- | --- | --- |
| [[BUG-039]] | `_safe_realpath` vs `os.path.realpath` 23 场景 + mock `_path_status`/`os.readlink` 抛 `TimeoutError` | 19/23 场景逐字节一致；超时两处均正确上抛（旧 `realpath` 同场景吞掉返回字面路径）；4 处差异全为「不存在的组件后接 `..`」，见 [[OPT-002]] |
| [[BUG-040]] | 标题含字面 `<table>` / `<TABLE>` / 后接真实表格 + 两段正文 | 标题、正文、表格全部保留（修复前正文全丢） |
| [[BUG-041]] | 288KB 输入内 120MB 目录名条目 | 拒绝（`ZIP 目录名条目声明非零内容`）；常驻内存增量 0.0MB（修复前 +316.9MB） |
| [[BUG-042]] | 正文 part 重定向 + 501 张图（真实上限 500） | `reject_redirected` → `ResourceLimitExceeded: 实际 501 张`；`skip_redirected` → 恰好写出 500 个 assets 文件（修复前两种模式均落盘 501） |
| [[BUG-043]] | 目录名图片条目被字面 Target 引用 | 扫描侧计入 3 条（修复前 0 条） |
| [[BUG-044]] | 标题含 `<a href=x>y</a>` / `<time>` / `<div>x</div>` | 不再改写成 `[y](x)`、不再删除，全部以实体保留且渲染可见 |
| [[BUG-045]] | 段落/单元格/文本框中的 `<!-- -->`、`<? ?>`、`<![CDATA[]]>` | 全部实体化，渲染可见文本完整（`正文 <!-- 隐藏 --> 结束。`） |
| [[BUG-046]] | 文本框字面 `<time>`/`<table>` | `> 正文 &lt;time&gt; 结束。`，渲染可见 |
| [[BUG-047]] | 打桩 `_allocate_asset_path`：在分配与独占写入之间用第三方内容占用候选 | md 引用改到 `image1_<hash>_2.png`，为普通文件且字节与输入一致；占用者未被改写。重试循环有界（98 个候选后上抛既有 `OSError`），两处调用点（提取循环、mammoth 回调）均已改造 |
| [[BUG-048]] | 首字符类放宽后的边界（15 组）+ 真实文档 | `<采暖>` → `&lt;采暖&gt;` 且渲染可见；「a < b 且 b > c」保持裸形式；跨度跨段/跨单元格/含 `<br>`/含 `&amp;` 均无内容损失或结构破坏。**台账对原误判的更正属实**：Python-Markdown 3.10.2 对裸 `<采暖>` 自动转义（渲染可见），ASCII 裸标签 `<time>` 才不可见 |
| [[TST-013]] | 包装 `T.load_module` 打桩被测模块的 `_allocate_asset_path` 抛真实 `OSError` | `run=1 failures=0 errors=1 skipped=0`，错误即该 `OSError`（修复前 `skipped=1`） |
| [[TST-014]] | 阅读 `tests/test_docx_markdown_quality.py:442-444`、`:565-569` | 两处 `ImportError` 分支均为 `self.fail(...)`；`markdown` 另有三处裸 `import`（缺库即报错），无静默跳过 |
| [[TST-015]] | `grep` 权限断言用例 | 4 个断言 POSIX 权限位的用例均带 `@unittest.skipUnless(os.name == "posix", ...)` |
| [[DOC-016]] | 阅读 `convert_docx.py:1720-1724` | 注释已改为「各 Mammoth 正文 part（动态解析，见 `_document_image_part_names`）」，与实现一致 |

### 真实文档端到端明细

`tests/` 下 6 份 DOCX，工作区与 HEAD 各 25 个产物，`diff -rq` 仅 `设备预约` 的 md 不同（16 行变更）。按渲染后可见文本做序列比对：

```
可见文本长度 head=7986 cur=8532
插入片段 92 个，删除片段 0 个，替换 0 个
插入片段全部为字面标签样文本（<采暖>、<time>、<room> 等）
```

即修复方向只增不减：找回 HEAD 丢失的字面槽位文本，无任何内容被删改。`assets/` 逐字节一致。

### 超时纪律回归

全文 40 处 `except` 逐条核对：所有 `except OSError` / `except Exception` / `except BaseException` 前均有 `except TimeoutError: raise`（`:136/138`、`:514/516`、`:581/583`、`:623/625`、`:705/707`、`:883/885`、`:1418/1420`、`:1474/1476`、`:1510/1512`、`:1656/1660`、`:1875/1877`、`:1916/1918`）；本轮新增的 `except FileExistsError`（`:536`、`:674`、`:1830`、`:2136`）为窄类型，不吞 `TimeoutError`；`_atomic_write_text` 的 `except BaseException`（`:562`）仅清理临时文件后原样 `raise`。`read_conversion_sentinel` 的 `except (OSError, ValueError)`（`:717`）仍缺守卫，但其唯一调用方 `batch_convert._is_output_complete`（`batch_convert.py:211`）在 `_run_with_timeout`（`:225`）之外、无待决 alarm，不可达。

---

## 三、本轮新发现

### [[BUG-049]]（P3）｜「纯目录占位排除」守卫恒真，目录占位条目被计入图片配额并留下 0 字节产物

`_document_image_part_names`（`convert_docx.py:224-226`）注释称「纯目录占位名（basename 为空，如 `word/media/`）不是图片数据」，但代码是 `posixpath.basename(name.rstrip("/"))`——`rstrip("/")` 后 `word/media/` → `word/media`，`basename` 为 `"media"`（真值），守卫从不生效。

```
$ python3 /tmp/rev/v3/probe_dirplaceholder.py
扫描到的图片条目: ['word/media/', 'word/media/image1.png']
含目录占位 'word/media/': True
'word/media/' 的 file_size: 0  is_dir(): True

$ 端到端（image 关系 Target="media/"）
转换成功: .../dirplaceholder.md
产出: ['image1.png', 'media.png']        ← 多出 0 字节 media.png，md 未引用
md 内容: '正文'
```

影响（均为 fail-closed 方向，不构成防线绕过）：
1. 每个指向目录的关系目标让 `image_count` 多计 1，边界文档（恰好等于 500 张）会被误拒；skip 模式多占一个配额槽位，可能少放一张真实图片；
2. 输出 `assets/` 留下 0 字节的 `media.png`——`prune_stale_assets` 按「已提取图片集合」（`convert_docx.py:2218` 传入 `image_by_hash.values()`）清理，不按 md 引用，故该文件不会被清掉。

触发前提是畸形文档（正常 Word 不产生指向目录的 image 关系）。修复建议：守卫改为判断「去掉尾 `/` 后是否仍含路径分隔符」或直接 `if name.rstrip("/") != "word/media" and posixpath.basename(name.rstrip("/"))`——即显式排除 `word/media/` 这一前缀条目本身。

### [[OPT-002]]｜`_safe_realpath` 与 `os.path.realpath` 在「不存在的组件后接 `..`」时输出不同

```
输入 '/tmp/rev/v3/real/nope/../other'
  got : '/private/tmp/rev/v3/real/nope/../other'   ← 字面拼接，未归一
  want: '/private/tmp/rev/v3/real/other'
```

`_safe_realpath`（`convert_docx.py:616-619`）在遇到不存在组件时按字面拼接剩余部分返回，不对剩余部分做 `..` 归一；`os.path.realpath` 非 strict 模式仍会归一。

**当前不构成防线绕过**：唯一两处调用点（`convert_docx.py:1994-1995`）用同一函数计算 `output_root_real` 与 `final_real`，且 `final_output_dir` 是 `output_dir` 的字面拼接扩展，`os.path.commonpath` 的字符串前缀关系在任何一侧都成立；文档字符串也已声明「组件不存在时剩余部分按字面拼接返回」。属语义一致性与文档措辞问题，非安全问题。

---

## 四、台账同步

新增：[[BUG-049]]（1 项，待修复）、[[OPT-002]]（1 项，待办）、[[CHK-023]]（本轮检查，已完成）。
状态变更：[[BUG-039]]、[[BUG-040]]～[[BUG-048]]、[[TST-013]]～[[TST-015]]、[[DOC-016]] 经本轮独立复现确认为真实修复，状态保持已完成/已修复。
