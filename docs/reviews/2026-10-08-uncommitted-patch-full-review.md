# 未提交补丁全面复审（V0.1.8 → V0.1.9）

2026-10-08 18:50（Asia/Shanghai）。用户要求：检查全部未提交代码是否还有 bug、issue 提出的问题是否都已修复、`task-list.md` 是否还有未修复的 bug。**本轮只做检查，未修改转换器/测试/文档，未提交发布，未操作 GitHub 状态**（仅本报告与 `task-list.md` 为新增/更新）。

审查范围：`skills/docx-to-markdown/scripts/convert_docx.py`（未提交补丁 742 行改动）、`tests/*.py`、`SKILL.md`/`usage-guide.md`/`README.md`、`task-list.md`、GitHub issue #1～#7。

方法：安全边界 / 内容保真 / 测试与文档三个只读子审查并行进行，**其报告的每一条发现都由我本人重新构造输入独立复现**后才写入本报告；探针脚本全部位于 `/tmp`（`/tmp/rev/v2/`、`/tmp/sec/`、`/tmp/fid/`、`/tmp/qa-probe/`）。

---

## 一、结论速览

| 项目 | 结论 |
| --- | --- |
| 全量测试 | 工作区 `python3 -m unittest discover -s tests` → **154 passed**；干净快照（`git archive HEAD` + 本次测试文件）→ 44/48 新增测试失败，回归信号真实 |
| GitHub issue | #1～#3 已关闭且实现落地；**#4/#5/#6 工作区已修复（本轮重新复现确认）**；#7 本地代码守卫在位，无 Windows 真机可验 |
| 台账未修复 bug | 仅 **[[BUG-039]]**（P3，realpath 吞超时）；其余 6 项待办均为非 bug 事项（[[TST-001]]、[[TST-002]]、[[PLN-001]]～[[PLN-003]]、[[OPT-001]]） |
| 本轮新发现 | **8 个 bug**（P1×2、P2×2、P3×4）+ 3 个测试缺陷 + 1 处陈旧注释，见第二节 |

**优先级建议**：[[BUG-040]]（标题含字面 `<table>` 静默截断整篇）与 [[BUG-041]]（目录名条目使 ZIP 体积防线全失效）应在本轮提交前一并修掉——前者是本次补丁引入的内容丢失回归，后者是"无条件防线"被完全绕过。

---

## 二、确认的问题（均已独立复现）

### BUG-040（P1，本次补丁引入的回归）｜标题含字面 `<table>` 会静默吞掉该处到文末的全部内容

`replace_html_tables` 深度计数分支在找不到配对 `</table>` 时，`end` 仍保持初值 `len(html)`（`convert_docx.py:1098`、`1101-1102`、`1113`），于是把「从这个 `<table>` 到文档结尾」整段当作一张表格交给 `table_html_to_markdown`，产出为空。触发源是 `_replace_anchored_heading` 在 `:2177` 提前 `unescape`，把标题里的字面 `&lt;table&gt;` 还原成裸标签，早于 `:2229` 的表格转换。

```
$ python3 /tmp/rev/v2/p1_heading.py
A: 标题「2.1 槽位 <table> 配置」+ 两段正文
  CUR  = '# 2.1 槽位'                                    ← 后续正文全丢
  HEAD = '# 2.1 槽位  配置\n\n后续正文A。\n\n后续正文B。'
B: 标题同上 + 后接真实表格
  CUR  = '# 2.1 槽位 | 列A | 列B |\n| --- | --- |\n| 1 | 2 |'   ← 表格前后正文一并丢失
  HEAD = '# 2.1 槽位 | 列A | 列B |\n| --- | --- |\n| 1 | 2 |\n\n表格后的正文。'
C: 标题含 <TABLE>（大写）→ CUR 同样只留 '# 2.1 槽位'
E: 对照 普通段落含 <table> → CUR '2.1 槽位 &lt;table&gt; 配置\n\n后续正文E。'（正常）
```

触发面：标题（`heading_N` 书签/`w:pStyle` 数字样式）文本含 `<table`（大小写不敏感，带属性亦可）。脚注、表格单元格、文本框路径不受影响（脚注实测见第三节）。修复建议：找不到配对结束标签时不消费该标签（`end` 保持 `None`，按字面放行并继续扫描），并在 `:2177` 之后把 `<`/`>` 重新实体化（可一并修掉 [[BUG-044]]）。

### BUG-041（P1）｜ZIP 目录名条目跳过尺寸/压缩比/总量统计，解压前防线全部失效

`validate_docx_zip_security` 在 `convert_docx.py:193-194` 对 `info.is_dir()`（名字以 `/` 结尾）直接 `continue`，位置在 `total_compressed/total_uncompressed`、`entry_uncompressed`、`entry_ratio` 三道检查之前；而 `zipfile`/mammoth 仍能按字面名打开并读取该条目。

```
$ python3 /tmp/sec/t6_zipbomb.py
[dir]   entry='word/real.xml/' declared=125830191 compressed=294277 ratio=428x file=288.7KB
[dir]   validate_docx_zip_security: PASSED (no exception)
[plain] entry='word/real.xml'  validate_docx_zip_security: REJECTED -> 单文件上限 100.0MB
conversion result: SUCCESS -> /tmp/sec/work6/out/dir/dir.md
peak RSS before=278790144 after=611057664 (delta 316.9MB)      ← 实测内存放大
```

289 KB 输入即可让进程解压解析 120 MB XML；多个此类条目可线性放大（总量上限同样不计数）。修复建议：目录名条目照常计入总量并执行单文件/压缩比检查；更稳妥是直接拒绝「名字以 `/` 结尾但 `file_size > 0`」的条目。

### BUG-042（P2）｜正文 part 被 `_rels/.rels` 重定向后图片数量配额可完全绕过，skip 模式仍经回调兜底落盘

`_document_image_part_names` 只读硬编码的 `_MAMMOTH_CONTENT_PARTS`（`convert_docx.py:105-109`），而 mammoth 按 `_rels/.rels` 的 `officeDocument` 关系动态解析正文 part。把正文 part 指向 `word/real.xml`（保留 `word/document.xml` 作为诱饵）并把图片放 `word/custom/`，扫描侧看到 0 张图片 → reject 模式数量防线不生效；skip 模式 `media_quota_exceeded` 恒为 False，回调兜底分支（`:1950-1963`）无配额检查地写盘。

```
$ python3 /tmp/sec/t2_redirect.py      # image_count=2，实际 3 张
RESULT: conversion SUCCEEDED (expected ResourceLimitExceeded)
assets: ['image_3fbc3cc75f05cd08.png', 'image_d2cd15d70c596c5b.png', 'image_e7355d3cfb652cc2.png']
$ python3 /tmp/sec/t3_real_limit.py    # 真实上限 500，实际 501 张
reject_plain        -> ResourceLimitExceeded: 图片数量超过上限 500: 实际 501 张
reject_redirected   -> SUCCESS, 501 asset files written
skip_plain          -> SUCCESS, 500 asset files written
skip_redirected     -> SUCCESS, 501 asset files written
```

修复建议：按 mammoth 同口径动态解析正文 part（读 `_rels/.rels` → 正文 part → 其 rels），并让回调兜底分支 fail-closed（无法证明配额内时返回可见跳过占位）。

### BUG-043（P2）｜目录名图片条目不计入数量防线，reject/skip 两种模式都写出超配额图片

同一处 `is_dir()` 过滤（`convert_docx.py:130-136`）使被字面 Target（`media/img1.png/`）引用的图片条目永不出现在 `image_part_names` 中；mammoth 回调仍能读到并落盘。

```
$ python3 /tmp/sec/t5_dirname.py       # image_count=2，实际 3 张
scan sees image parts: []
RESULT reject(limit=2): SUCCESS
  assets: 3 个 image_<hash>.png
RESULT skip(limit=2): SUCCESS
  assets: 3 个 image_<hash>.png
```

单图大小/像素防线仍在回调内生效，故影响限于图片数量与解压总量（与 [[BUG-041]] 叠加后总量防线亦失效）。

### BUG-044（P3，HEAD 亦存在）｜标题中的字面标签被删除或被改写成真实 Markdown 链接

同 `:2177` 提前 unescape 的根因。

```
$ python3 /tmp/rev/v2/misc.py
标题含 <a href=x>y</a>  -> '# 2.1 槽位 [y](x) 配置'      ← 字面标签被改写成真实链接
标题含 <time>           -> '# 2.1 槽位  配置'             ← 字面标签消失
标题含 <div>x</div>     -> '# 2.1 槽位 x 配置'
```

### BUG-045（P3）｜实体保护只覆盖字母开头标签，`<!-- -->` / `<? ?>` / `<![CDATA[]]>` 渲染后不可见

最终保护正则 `&lt;(/?[A-Za-z][^\x00\x01]*?)&gt;`（`convert_docx.py:2249`）与 `_escape_plain_cell_text` 的 `</?[A-Za-z][^<>]*>`（`:662`）都要求字母或 `/字母` 开头。

```
$ python3 /tmp/rev/v2/entities.py
段落  "正文 <!-- 隐藏 --> 结束。"   CUR md='正文 <!-- 隐藏 --> 结束。'  渲染可见='正文  结束。'   ← 文字消失
段落  "正文 <?php echo 1?> 结束。"  CUR md='正文 <?php echo 1?> 结束。'  渲染可见='正文  结束。'
段落  "正文 <![CDATA[abc]]> 结束。" CUR md='正文 <![CDATA[abc]]> 结束。' 渲染可见='正文  结束。'
表格单元格 / 文本框 同样不可见
```

HEAD 在这几例中直接丢字（`'正文  结束。'`），本补丁已把「丢字」改善为「源文件有、渲染不可见」，属同类缺陷的残余部分。修复建议：字符类放宽到 `[/!?A-Za-z]`。

### BUG-046（P3，HEAD 亦存在）｜文本框内容未经转义直接拼入 Markdown

`extract_textbox_content` 的结果以 `> {block}` 原样追加（`convert_docx.py:2007-2016`），绕过整条实体保护管线；文本框里的字面 `<time>`、`<table>` 以裸标签进入 Markdown，渲染时不可见（`/tmp/rev/v2/entities.py` 文本框段：CUR md 为 `> 正文 <time> 结束。`，渲染可见 `正文  结束。`）。文本框追加发生在 `html_to_markdown` 之后，故不会触发 [[BUG-040]] 的截断。

### BUG-047（P3）｜`_write_asset_file_exclusive` 遇 `FileExistsError` 静默返回，与文档承诺不一致

`convert_docx.py:538-549`：独占创建失败即 `return`，调用方（`:1679`、`:1963`）不校验落盘结果。分配与写入之间的窄竞争窗口内被第三方占用时，Markdown 引用会指向符号链接/外来文件，而 `SKILL.md:214`、`usage-guide.md:460-461`、`README.md:114` 承诺「每个引用都指向内容正确的普通文件」。（同线程内无竞争，触发条件窄。）

### 测试与文档类

- **TST-013（P2）**：`tests/test_docx_security_and_batch.py:777-783` 的 `except (OSError, NotImplementedError)` 包住了整段 `_occupied_candidate_test("link")`，把 `_allocate_asset_path` 抛出的真实 `OSError`（"无法为图片分配可用的输出文件名"）报成「环境不支持符号链接」跳过。独立复现（把被测模块的 `_allocate_asset_path` 打桩抛错后运行该用例）：`run=1 failures=0 errors=0 skipped=1`，跳过原因为该 OSError。同文件 `:1007-1012` 的 `_symlink_or_skip` 是正确写法。
- **TST-014（P3）**：唯一在渲染层验证 [[BUG-020]] 的 5 个用例依赖可选包 `markdown`，缺库时 `skipTest`（`tests/test_docx_markdown_quality.py:435-440`），套件仍报 OK，声称的渲染回归覆盖静默蒸发。
- **TST-015（P3，未在 Windows 验证）**：`test_new_files_respect_caller_umask` / `test_replacement_preserves_existing_file_mode`（`tests/test_docx_security_and_batch.py:799-822`）直接断言 POSIX 权限位，无平台守卫；本机 macOS 通过，Windows 预期失败。
- **DOC-016（P3）**：`convert_docx.py:1576-1578` 注释仍写「物理图片条目 = word/media/* ∪ document.xml.rels 的 image 关系目标」，与 [[BUG-026]] 后覆盖 document/footnotes/endnotes 三 part 的实现不符（`usage-guide.md:104` 已正确）。

---

## 三、已复核且未发现问题的区域

- **issue #4/#5/#6 修复有效**（本轮用工作区代码重新复现）：#4 输出子目录为符号链接 → `ValueError: 输出子目录已存在且不是普通目录（不允许符号链接）`，外部 `keep.bin`（`KEEP-DO-NOT-DELETE`）与 `evil.md`（`ORIGINAL`）均未被改写；#5 嵌套表格外层 2 列 2 行，外层 A1/B1/B2 与内层 X/Y 及后续段落渲染全部可见；#6 单行 `[![](assets/image1.png)](https://example.com/target)`，链接文字变体亦正常。#7 代码守卫在位（无 Windows 真机）。
- **脚注路径不受 [[BUG-040]] 影响**：脚注含字面 `<table>`/`<time>` 时（`/tmp/rev/v2/footnote_table.py`）正文与脚注定义均完整保留，产出 `[^1]: 脚注里的 &lt;table&gt; 标签说明。`。
- **真实文档端到端回归**（本轮之前完成，代码未再变动）：`tests/` 下 6 份真实 DOCX，工作区与 HEAD 各 25 个产物，`diff -rq` 仅 `设备预约` 的 md 不同（7113→7611 字符，0 处丢失、92 处找回的字面标签），assets 逐字节一致，表格 46 行全部保留。
- **TimeoutError 传播**：安全子审查枚举全部 30 处 `except` 并做 400 次真实 `setitimer` 扫描（`{'TimeoutError': 216, 'returned-normally': 184}`），未出现静默降级；`read_conversion_sentinel` 的 `except (OSError, ValueError)` 缺守卫但其调用方在计时窗口外，不可达。[[BUG-039]] 仍按原结论（`:1840-1841` 的 `os.path.realpath` 吞超时）。
- **内容保真其余边界**（保真子审查，均为其实测且与实现一致）：Excel/Word 单元格 1/2/3/4 条反斜线加管道、行首行尾管道、反引号/星号/下划线、嵌套表格三层、脚注 18 组边界（CRLF、单引号属性、属性含 `>`、未闭合标签、多段落、脚注内图片与超限跳过）、图片与链接组合（链接内两张图、`<br>` 文本、无引号 src）均无丢字错位。
- **测试与文档一致性**：测试方法名无重复；44/48 新增测试在 HEAD 上失败（回归信号真实）；`task-list.md` 97 行 ID 无重复、44 个 `[[ID]]` 引用全部可解析、状态与动作取值均在词表内；测试项数声明与实测 154 一致。

---

## 四、台账同步

新增：[[BUG-040]]～[[BUG-047]]（8 项，均待修复）、[[TST-013]]～[[TST-015]]（3 项，待办）、[[DOC-016]]（1 项，待办）、[[CHK-020]]（本轮检查，已完成）。
未变更：[[BUG-039]] 仍为待修复（P3，等用户决定是否本轮修）。
