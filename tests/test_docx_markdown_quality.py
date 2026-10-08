import importlib.util
import os
import re
import tempfile
import unittest
from pathlib import Path


WORKSPACE_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = WORKSPACE_ROOT / "skills" / "docx-to-markdown" / "scripts" / "convert_docx.py"
DOCX_DIST_WAKE = WORKSPACE_ROOT / "tests" / "分布式唤醒V1.9.1—【唤醒体验v03】播控指令与唤醒暂停的策略梳理.docx"
DOCX_HEAT_APPOINT = WORKSPACE_ROOT / "tests" / "设备预约V2.7.2—采暖炉新增采暖开启_关闭的预约.docx"
DOCX_VAD = WORKSPACE_ROOT / "tests" / "自研语义VAD（云端VAD-4.0）型号接入需求文档V2.docx"


def load_convert_module():
    spec = importlib.util.spec_from_file_location("convert_docx_module", SCRIPT_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def count_malformed_table_blocks(content: str) -> int:
    bad_blocks = 0
    in_block = False
    expected_pipes = 0

    def unescaped_pipe_count(line: str) -> int:
        return len([m for m in re.finditer(r"(?<!\\)\|", line)])

    for line in content.splitlines() + [""]:
        is_table_line = line.startswith("|") and line.rstrip().endswith("|")
        if is_table_line:
            pipe_count = unescaped_pipe_count(line)
            is_separator = set(line.replace("|", "").strip()) <= {"-", ":", " "}
            if not in_block:
                in_block = True
                expected_pipes = pipe_count
            elif pipe_count != expected_pipes and not is_separator:
                bad_blocks += 1
                in_block = False
                expected_pipes = 0
        else:
            in_block = False
            expected_pipes = 0

    return bad_blocks


class TestMarkdownQualityRegression(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_convert_module()

    def test_numbered_bold_paragraphs_are_promoted_to_headings(self):
        html = (
            "<p>1. <strong>需求背景</strong></p>"
            "<p>3.1 <strong>终端应用</strong></p>"
            "<p>普通内容</p>"
        )

        markdown = self.convert.html_to_markdown(html)

        self.assertIn("# 1. 需求背景", markdown)
        self.assertIn("## 3.1 终端应用", markdown)
        self.assertIn("普通内容", markdown)

    def test_html_table_with_rowspan_and_multiline_cell_is_stable(self):
        html = (
            "<table>"
            "<tr><td rowspan='2'><p>采暖炉</p></td><td><p>time</p></td><td><p>time词典</p><p>intervalTime词典</p></td></tr>"
            "<tr><td><p>deviceName</p></td><td><p>平台统一</p></td></tr>"
            "</table>"
        )

        markdown = self.convert.html_to_markdown(html)
        table_lines = [line for line in markdown.splitlines() if line.startswith("|") and line.endswith("|")]

        self.assertGreaterEqual(len(table_lines), 3)
        self.assertEqual(0, count_malformed_table_blocks(markdown))
        self.assertIn("time词典<br>intervalTime词典", markdown)

    def test_list_item_with_paragraph_does_not_introduce_loose_list_gap(self):
        html = "<ul><li><p>text</p></li><li><p>text2</p></li></ul>"
        markdown = self.convert.html_to_markdown(html)

        self.assertIn("- text", markdown)
        self.assertIn("- text2", markdown)
        self.assertNotIn("- text\n\n- text2", markdown)

    def test_nested_unordered_list_keeps_hierarchy(self):
        html = "<ul><li>item1<ul><li>nested</li></ul></li></ul>"
        markdown = self.convert.html_to_markdown(html)

        self.assertIn("- item1", markdown)
        self.assertIn("  - nested", markdown)

    def test_table_after_ordered_list_has_blank_line_separator(self):
        html = (
            "<ol><li>第一条</li><li>第二条</li></ol>"
            "<table><tr><td>背景</td><td>执行结果</td></tr>"
            "<tr><td>场景A</td><td>成功</td></tr></table>"
        )
        markdown = self.convert.html_to_markdown(html)

        self.assertIn("2. 第二条\n\n| 背景 | 执行结果 |", markdown)

    @unittest.skipUnless(DOCX_VAD.is_file(), "可选真实 DOCX 回归夹具未提供")
    def test_vad_doc_keeps_original_numbered_subsections(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            vad_md = self.convert.convert_docx_to_markdown(str(DOCX_VAD), tmpdir)
            content = Path(vad_md).read_text(encoding="utf-8")

        self.assertIn("# 4. 第三步：立项和需求", content)
        self.assertIn("# 5. 第四步：上线前的验证", content)
        self.assertIn("# 6. 第五步：质量部发布上线测试报告", content)
        self.assertIn("### 1. 偏差投诉与反馈", content)
        self.assertIn("### 1. 推理阶段的隐私保护", content)
        self.assertIn("### 1. 模型自身的安全性", content)

    @unittest.skipUnless(DOCX_DIST_WAKE.is_file(), "可选真实 DOCX 回归夹具未提供")
    def test_dist_wake_doc_keeps_level3_numbered_subheadings(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            md_path = self.convert.convert_docx_to_markdown(str(DOCX_DIST_WAKE), tmpdir)
            content = Path(md_path).read_text(encoding="utf-8")

        self.assertIn("## 3.1 终端应用", content)
        self.assertIn("### 1. 终端播控状态改变的三个途径", content)
        self.assertIn("### 2. 播控暂停比唤醒暂停和恢复优先级更高", content)
        self.assertIn("### 3. 终端设备应该遵循的播控业务逻辑", content)

    @unittest.skipUnless(
        DOCX_DIST_WAKE.is_file() and DOCX_HEAT_APPOINT.is_file(),
        "可选真实 DOCX 回归夹具未提供",
    )
    def test_regression_docs_have_no_malformed_markdown_table_blocks(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            dist_md = self.convert.convert_docx_to_markdown(str(DOCX_DIST_WAKE), tmpdir)
            heat_md = self.convert.convert_docx_to_markdown(str(DOCX_HEAT_APPOINT), tmpdir)

            dist_content = Path(dist_md).read_text(encoding="utf-8")
            heat_content = Path(heat_md).read_text(encoding="utf-8")

        self.assertEqual(0, count_malformed_table_blocks(dist_content))
        self.assertEqual(0, count_malformed_table_blocks(heat_content))

    def test_same_named_parent_sections_do_not_cross_increment_subheading_counter(self):
        markdown = (
            "## Section\n\n"
            "### 1. 子项A\n\n"
            "## Section\n\n"
            "### 1. 子项B\n"
        )

        result = self.convert.promote_numbered_bold_headings(markdown)
        self.assertIn("### 1. 子项A", result)
        self.assertIn("### 1. 子项B", result)
        self.assertNotIn("### 2. 子项B", result)

    def test_multilevel_numbered_headings_keep_original_numbers(self):
        markdown = (
            "## Parent\n\n"
            "### 1.1 子项A\n\n"
            "### 1.1 子项B\n"
        )

        result = self.convert.promote_numbered_bold_headings(markdown)
        self.assertIn("### 1.1 子项A", result)
        self.assertIn("### 1.1 子项B", result)

    def test_html_image_src_supports_single_and_unquoted_attr(self):
        markdown_single = self.convert.html_to_markdown("<p>txt</p><img src='a.png' alt='x'>")
        markdown_unquoted = self.convert.html_to_markdown("<p>txt</p><img src=a.png alt=x>")

        self.assertIn("![](a.png)", markdown_single)
        self.assertIn("![](a.png)", markdown_unquoted)

    def test_invalid_docx_raises_clear_value_error(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            bad_docx = Path(tmpdir) / "bad.docx"
            bad_docx.write_text("not a zip", encoding="utf-8")

            with self.assertRaises(ValueError):
                self.convert.convert_docx_to_markdown(str(bad_docx), tmpdir)

    def test_anchored_heading_keeps_literal_markdown_stars(self):
        html = "<p><a id='heading_1'></a>术语 **KEEP**</p>"
        markdown = self.convert.html_to_markdown(html, {"heading_1": 2})

        self.assertIn("## 术语 **KEEP**", markdown)

    def test_bilingual_parallel_numbered_headings_do_not_cross_increment(self):
        markdown = "## 1. DEFINITIONS\n\n## 1. 定义\n"
        result = self.convert.promote_numbered_bold_headings(markdown)

        self.assertIn("## 1. DEFINITIONS", result)
        self.assertIn("## 1. 定义", result)
        self.assertNotIn("## 2. 定义", result)

    def test_leading_full_bold_line_is_promoted_to_h1_when_numbered_sections_exist(self):
        html = (
            "<p><strong>文档标题</strong></p>"
            "<p><a id='heading_0'></a>1. <strong>第一章</strong></p>"
            "<p>正文</p>"
        )
        markdown = self.convert.html_to_markdown(html, {"heading_0": 1})

        self.assertIn("# 文档标题", markdown)
        self.assertIn("# 1. 第一章", markdown)

    def test_leading_full_bold_line_without_numbered_sections_is_not_promoted(self):
        html = "<p><strong>仅强调文本</strong></p><p>普通正文</p>"
        markdown = self.convert.html_to_markdown(html)

        self.assertIn("**仅强调文本**", markdown)
        self.assertNotIn("# 仅强调文本", markdown)

    # --- E2: 残留预览文本清理 ---
    def test_table_placeholder_removes_residual_preview_text(self):
        md = "| A |\n| --- |\n| 1 |\n\n**点击图片可查看完整电子表格**\n\n后续段落"
        cleaned = re.sub(
            r"\n+\*{0,2}点击图片可查看完整电子表格\*{0,2}\s*\n",
            "\n",
            md,
        )
        self.assertNotIn("点击图片", cleaned)
        self.assertIn("后续段落", cleaned)

    # --- E3: 脚注转换 ---
    def test_footnote_html_converts_to_markdown_footnote_syntax(self):
        html = (
            '<p>Clause<sup><a href="#footnote-1" id="footnote-ref-1">[1]</a></sup> text</p>'
            '<p>Another<sup><a href="#footnote-2" id="footnote-ref-2">[2]</a></sup></p>'
            '<ol><li id="footnote-1"><p>First note <a href="#footnote-ref-1">↑</a></p></li>'
            '<li id="footnote-2"><p>Second note <a href="#footnote-ref-2">↑</a></p></li></ol>'
        )
        markdown = self.convert.html_to_markdown(html)
        self.assertIn("[^1]", markdown)
        self.assertIn("[^2]", markdown)
        self.assertIn("[^1]: First note", markdown)
        self.assertIn("[^2]: Second note", markdown)

    def test_footnote_body_keeps_image_and_skip_placeholder(self):
        """脚注体内的图片与超限占位不得在纯文本剥离中丢失（BUG-029）。"""
        html = (
            '<p>正文<sup><a href="#footnote-1" id="footnote-ref-1">[1]</a></sup></p>'
            '<ol>'
            '<li id="footnote-1"><p>说明 <img src="assets/image1.png" /> '
            '<a href="#footnote-ref-1">↑</a></p></li>'
            '<li id="footnote-2"><p><img src="__SKIPPED_IMAGE_count__" /> '
            '<a href="#footnote-ref-2">↑</a></p></li>'
            '</ol>'
        )
        markdown = self.convert.html_to_markdown(html)
        self.assertIn("[^1]", markdown)
        self.assertRegex(markdown, r"\[\^1\]: 说明 !\[\]\(assets/image1\.png\)")
        self.assertRegex(markdown, r"\[\^2\]: !\[\]\(__SKIPPED_IMAGE_count__\)")
        self.assertNotIn("↑", markdown)

    def test_footnote_literal_tags_survive_conversion_and_rendering(self):
        """BUG-031：脚注中的字面标签必须作为可见文本保留。"""
        import markdown as markdown_lib

        html = (
            '<p>正文<sup><a href="#footnote-1">[1]</a></sup></p>'
            '<ol><li id="footnote-1"><p>字面 &lt;time&gt; &lt;table&gt; '
            '&amp; END</p><a href="#footnote-ref-1">↑</a></li></ol>'
        )
        md = self.convert.html_to_markdown(html)
        self.assertIn("[^1]: 字面 &lt;time&gt; &lt;table&gt; & END", md)
        rendered = markdown_lib.markdown(md, extensions=["footnotes"])
        self.assertIn("字面 &lt;time&gt; &lt;table&gt; &amp; END", rendered)
        self.assertNotIn("<time>", rendered)
        self.assertNotIn("<table>", rendered)

    def test_footnote_blank_lines_keep_images_on_definition_line(self):
        """BUG-032：源文本和字符实体中的空行不得截断定义。"""
        import markdown as markdown_lib

        html = (
            '<p>正文<sup><a href="#footnote-1">[1]</a></sup></p>'
            '<ol><li id="footnote-1"><p>Alpha\n\nBeta&#10;&#10;Gamma'
            '&#9;&nbsp;Delta<img src="assets/image1.png"/>'
            '<br/><img src="__SKIPPED_IMAGE_count__"/></p></li></ol>'
        )
        md = self.convert.html_to_markdown(html)
        definition = next(line for line in md.splitlines() if line.startswith("[^1]:"))
        self.assertEqual(
            definition,
            "[^1]: Alpha Beta Gamma Delta![](assets/image1.png) "
            "![](__SKIPPED_IMAGE_count__)",
        )
        rendered = markdown_lib.markdown(md, extensions=["footnotes"])
        main, footnotes = rendered.split('<div class="footnote">', 1)
        self.assertNotIn("<img", main)
        self.assertIn('src="assets/image1.png"', footnotes)
        self.assertIn('src="__SKIPPED_IMAGE_count__"', footnotes)

    def test_footnote_nested_lists_keep_tail_and_images_in_definition(self):
        """BUG-033：多层列表、后续段落与相邻脚注必须完整提取。"""
        import markdown as markdown_lib

        html = (
            '<p>正文<sup><a href="#footnote-1">[1]</a></sup>'
            '<sup><a href="#footnote-2">[2]</a></sup></p>'
            '<ol><li id="footnote-1"><p>说明</p><ul><li>第一项'
            '<ol><li>内层</li></ol></li><li>第二项'
            '<img src="assets/image2.png"/>'
            '<img src="__SKIPPED_IMAGE_count__"/></li></ul><p>末尾</p>'
            '<a href="#footnote-ref-1">↑</a></li>'
            '<li id="footnote-2"><p>相邻脚注</p></li></ol>'
            '<p>正文尾部</p>'
        )
        md = self.convert.html_to_markdown(html)
        definition = next(line for line in md.splitlines() if line.startswith("[^1]:"))
        self.assertEqual(
            definition,
            "[^1]: 说明 第一项 内层 第二项![](assets/image2.png)"
            "![](__SKIPPED_IMAGE_count__) 末尾",
        )
        self.assertIn("[^2]: 相邻脚注", md)
        self.assertIn("正文尾部", md)
        self.assertNotIn("↑", md)
        rendered = markdown_lib.markdown(md, extensions=["footnotes"])
        main, footnotes = rendered.split('<div class="footnote">', 1)
        self.assertNotIn("第二项", main)
        self.assertNotIn("<img", main)
        self.assertIn("第二项", footnotes)
        self.assertIn('src="assets/image2.png"', footnotes)

    # --- E6: --force 模式（批量转换） ---
    @unittest.skipUnless(DOCX_DIST_WAKE.is_file(), "可选真实 DOCX 回归夹具未提供")
    def test_batch_convert_force_reconverts_existing(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            md_path = self.convert.convert_docx_to_markdown(str(DOCX_DIST_WAKE), tmpdir)
            content_first = Path(md_path).read_text(encoding="utf-8")

            import importlib.util as _ilu
            batch_spec = _ilu.spec_from_file_location(
                "batch_convert_module",
                WORKSPACE_ROOT / "skills" / "docx-to-markdown" / "scripts" / "batch_convert.py",
            )
            batch_mod = _ilu.module_from_spec(batch_spec)
            batch_spec.loader.exec_module(batch_mod)

            src_dir = str(DOCX_DIST_WAKE.parent)
            batch_mod.batch_convert(src_dir, tmpdir, force=True)

            content_second = Path(md_path).read_text(encoding="utf-8")
            self.assertEqual(content_first, content_second)

    @unittest.skipUnless(DOCX_DIST_WAKE.is_file(), "可选真实 DOCX 回归夹具未提供")
    def test_force_handles_same_name_file_without_aborting(self):
        """--force 遇到同名普通文件（非目录）时不应中断批处理。"""
        with tempfile.TemporaryDirectory() as tmpdir:
            out = Path(tmpdir) / "out"
            out.mkdir()
            stem = self.convert.sanitize_stem(DOCX_DIST_WAKE.stem)
            (out / stem).write_text("placeholder", encoding="utf-8")
            self.assertTrue((out / stem).is_file())

            import importlib.util as _ilu
            batch_spec = _ilu.spec_from_file_location(
                "batch_convert_module2",
                WORKSPACE_ROOT / "skills" / "docx-to-markdown" / "scripts" / "batch_convert.py",
            )
            batch_mod = _ilu.module_from_spec(batch_spec)
            batch_spec.loader.exec_module(batch_mod)

            batch_mod.batch_convert(str(DOCX_DIST_WAKE.parent), str(out), force=True)
            self.assertTrue((out / stem).is_dir())

    def test_multi_paragraph_footnote_preserves_separation(self):
        """多段脚注正文不应被拼接成一个词。"""
        html = (
            '<p>Text<sup><a href="#footnote-1" id="footnote-ref-1">[1]</a></sup></p>'
            '<ol><li id="footnote-1">'
            "<p>Alpha.</p><p>Beta.</p>"
            '<a href="#footnote-ref-1">↑</a></li></ol>'
        )
        markdown = self.convert.html_to_markdown(html)
        self.assertIn("[^1]: Alpha. Beta.", markdown)
        self.assertNotIn("Alpha.Beta.", markdown)

    # --- Issue #5: 嵌套表格 ---
    def test_nested_table_keeps_outer_cell_after_inner_table(self):
        """内层表格之后的外层单元格不得丢失或被挤成正文段落。"""
        html = (
            "<p>嵌套表格测试</p>"
            "<table>"
            "<tr><td><p>外层A1</p></td><td><p>外层B1</p></td></tr>"
            "<tr><td>"
            "<table><tr><td><p>内层X</p></td><td><p>内层Y</p></td></tr></table>"
            "</td><td><p>外层B2</p></td></tr>"
            "</table>"
            "<p>表格之后的段落</p>"
        )
        markdown = self.convert.html_to_markdown(html)
        table_lines = [line for line in markdown.splitlines() if line.startswith("|") and line.endswith("|")]

        # 内层之后的外层单元格内容保留在表格行内，而不是脱离表格
        self.assertTrue(any("外层B2" in line for line in table_lines))
        self.assertIn("外层A1", markdown)
        self.assertIn("外层B1", markdown)
        # 内层表格内容不丢失（单元格内被转义为文本）
        self.assertIn("内层X", markdown)
        self.assertIn("内层Y", markdown)
        self.assertEqual(0, count_malformed_table_blocks(markdown))

    # --- Issue #6: 带超链接的图片 ---
    def test_hyperlink_wrapped_image_stays_single_line(self):
        """链接内图片应产出单行 [![](src)](url)，不得被换行拆断。"""
        html = '<p><a href="https://example.com/target"><img src="assets/image1.png" /></a></p>'
        markdown = self.convert.html_to_markdown(html)
        self.assertIn("[![](assets/image1.png)](https://example.com/target)", markdown)
        self.assertNotIn("[![](assets/image1.png)\n", markdown)
        # 链接内非图片文本与链接外的普通图片不受影响
        plain = self.convert.html_to_markdown(
            '<p><a href="https://x.com">click</a> <img src="a.png" /></p>'
        )
        self.assertIn("[click](https://x.com)", plain)
        self.assertIn("![](a.png)", plain)

    # --- BUG-021: 图片属性等号后的空白 ---
    def test_image_src_tolerates_whitespace_after_equals(self):
        """src= 与引号/值之间的空白是合法 HTML，不得让图片消失。"""
        self.assertIn("![](a.png)", self.convert.html_to_markdown('<p><img src= "a.png" /></p>'))
        self.assertIn("![](b.png)", self.convert.html_to_markdown("<p><img src=\t'b.png' /></p>"))
        self.assertIn("![](c.png)", self.convert.html_to_markdown("<p><img src=\nc.png /></p>"))
        linked = self.convert.html_to_markdown(
            '<p><a href="https://example.com/t"><img src=  "d.png"/></a></p>')
        self.assertIn("[![](d.png)](https://example.com/t)", linked)

    # --- BUG-020: 嵌套表格的字面标签/实体/转义层次 ---
    def _render_with_project_markdown(self, markdown):
        # markdown 是项目 PDF 路径（md_to_pdf.py）的运行时依赖，因此是
        # 测试必需依赖：缺库时显性失败而非静默跳过——否则唯一在渲染层
        # 验证 BUG-020/045/048 的用例会在该环境静默蒸发，套件仍报 OK
        # （见 TST-014）。安装：pip install markdown
        try:
            import markdown as markdown_lib
        except ImportError:
            self.fail("渲染回归需要 markdown 库（项目 PDF 路径的运行时依赖），"
                      "当前环境未安装：pip install markdown")
        return markdown_lib.markdown(markdown, extensions=["tables", "fenced_code"])

    @staticmethod
    def _visible_text(rendered_html):
        from html import unescape
        return unescape(re.sub(r"<[^>]+>", "", rendered_html))

    def test_nested_table_preserves_literal_tag_text(self):
        """内层单元格中的字面 <table> 文本不得在递归解析或标签清理中丢失，
        且渲染后作为可见文本出现（不被当作原始 HTML 吞掉）。"""
        html = (
            "<table><tr><td><p>外层A1</p></td></tr>"
            "<tr><td>"
            "<table><tr><td><p>保留 &lt;table&gt; 标签</p></td><td><p>内层Y</p></td></tr></table>"
            "</td><td><p>外层B2</p></td></tr></table>"
        )
        markdown = self.convert.html_to_markdown(html)

        # 源文件保留实体形式（裸 <table> 会被渲染器当作原始 HTML）
        self.assertIn("保留 &lt;table&gt; 标签", markdown)
        self.assertIn("内层Y", markdown)
        self.assertIn("外层B2", markdown)
        self.assertEqual(0, count_malformed_table_blocks(markdown))

        rendered = self._render_with_project_markdown(markdown)
        # 仅外层一个表格元素：字面标签不得形成嵌套 HTML 元素
        self.assertEqual(rendered.count("<table>"), 1, rendered)
        visible = self._visible_text(rendered)
        self.assertIn("保留 <table> 标签", visible)
        self.assertIn("内层Y", visible)
        self.assertIn("外层B2", visible)

    def test_literal_time_slot_renders_as_visible_text(self):
        """正文字面 <time> 槽位文本渲染后必须可见（PDF 路径验收补充）。"""
        html = (
            "<p>开头</p>"
            "<table><tr><td>"
            "<table><tr><td><p>槽位 &lt;time&gt; 文本</p></td></tr></table>"
            "</td><td><p>外层B2</p></td></tr></table>"
            "<p>结尾</p>"
        )
        rendered = self._render_with_project_markdown(self.convert.html_to_markdown(html))

        self.assertNotIn("<time>", rendered)
        visible = self._visible_text(rendered)
        self.assertIn("槽位 <time> 文本", visible)
        self.assertIn("外层B2", visible)
        self.assertIn("结尾", visible)

    def test_nested_table_entity_text_round_trips(self):
        """内层单元格的实体文本（& < >）只解码一次，不提前解码或二次解码。"""
        html = (
            "<table><tr><td>"
            "<table><tr><td><p>5 &amp; 6 &lt; 7 &gt; 4</p></td></tr></table>"
            "</td><td><p>外层B2</p></td></tr></table>"
        )
        markdown = self.convert.html_to_markdown(html)

        self.assertIn("5 & 6 < 7 > 4", markdown)

    def test_excel_backslash_pipe_combinations_render_single_cell(self):
        """原文 1/2/3 条反斜线相邻管道时，渲染必须保持单格且尾列不丢。

        仅检查管道前一个字符会把原文连续反斜线误判为已有转义，
        导致 Python Markdown 按裸管道拆列（见 BUG-020 验收补充）。
        """
        import io
        import openpyxl

        cases = [
            (r"A|B", "A|B"),
            (r"A\|B", "A\\|B"),
            (r"A\\|B", "A\\\\|B"),
            (r"A\\\|B", "A\\\\\\|B"),
        ]
        for raw, _ in cases:
            with self.subTest(raw=raw):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(["H1", "H2"])
                ws.append([raw, "TAIL"])
                buf = io.BytesIO()
                wb.save(buf)
                wb.close()

                table_md = self.convert.excel_to_markdown(buf.getvalue())
                rendered = self._render_with_project_markdown(table_md)
                cells = re.findall(r"<t[dh]>(.*?)</t[dh]>", rendered)

                self.assertIn(raw, cells, rendered)  # 原文完整保留在单格
                self.assertIn("TAIL", cells, rendered)  # 尾列不丢

    def test_nested_word_table_backslash_pipe_renders_single_cell(self):
        """嵌套 Word 表格内的反斜线+管道文本同样不拆列、不丢尾格。"""
        html = (
            "<table>"
            "<tr><td><p>外层A1</p></td><td><p>外层B1</p></td></tr>"
            "<tr><td>"
            "<table><tr><td><p>A\\\\|B</p></td><td><p>内层Y</p></td></tr></table>"
            "</td><td><p>外层B2</p></td></tr>"
            "</table>"
            "<p>表格之后的尾部段落</p>"
        )
        markdown = self.convert.html_to_markdown(html)
        rendered = self._render_with_project_markdown(markdown)

        cells = re.findall(r"<t[dh]>(.*?)</t[dh]>", rendered)
        self.assertTrue(any("A\\\\|B" in c for c in cells), rendered)
        self.assertTrue(any("外层B2" in c and "A" not in c for c in cells), rendered)
        self.assertIn("<p>表格之后的尾部段落</p>", rendered)

    def test_nested_table_pipe_and_tail_render_with_project_markdown(self):
        """内层含管道的嵌套表格经项目 Python Markdown 渲染不丢列、不丢尾部。

        项目 PDF 路径（md_to_pdf.py）使用 markdown.markdown(extensions=
        ['tables', 'fenced_code'])；二次转义产生的连续反斜线会让外层
        第二行的单元格错位、尾部内容脱离表格（见 BUG-020）。
        """
        try:
            import markdown as markdown_lib
        except ImportError:
            # markdown 为项目 PDF 路径的运行时依赖，缺库属环境不完整，
            # 显性失败而非静默跳过（见 TST-014）
            self.fail("渲染回归需要 markdown 库，当前环境未安装: pip install markdown")
        html = (
            "<table>"
            "<tr><td><p>外层A1</p></td><td><p>外层B1</p></td></tr>"
            "<tr><td>"
            "<table><tr><td><p>A|B</p></td><td><p>内层Y</p></td></tr></table>"
            "</td><td><p>外层B2</p></td></tr>"
            "</table>"
            "<p>表格之后的尾部段落</p>"
        )
        markdown = self.convert.html_to_markdown(html)
        rendered = markdown_lib.markdown(markdown, extensions=["tables", "fenced_code"])

        cells = re.findall(r"<td>(.*?)</td>", rendered, flags=re.DOTALL)
        self.assertTrue(any("A|B" in c for c in cells), rendered)
        self.assertTrue(any("外层B2" in c and "A|B" not in c for c in cells), rendered)
        self.assertIn("<p>表格之后的尾部段落</p>", rendered)
        self.assertNotIn("\\|", rendered)  # 不残留未消解的转义序列

    def test_three_level_nested_table_keeps_all_text(self):
        html = (
            "<table><tr><td><p>外层A1</p></td></tr><tr><td>"
            "<table><tr><td><p>中层M</p></td><td>"
            "<table><tr><td><p>内层最深</p></td></tr></table>"
            "</td></tr></table>"
            "</td><td><p>外层B2</p></td></tr></table><p>尾部段落</p>"
        )
        markdown = self.convert.html_to_markdown(html)

        for text in ("外层A1", "中层M", "内层最深", "外层B2", "尾部段落"):
            self.assertIn(text, markdown)
        self.assertEqual(0, count_malformed_table_blocks(markdown))


class TestLiteralTagTextFidelity(unittest.TestCase):
    """BUG-040/044/045/046 回归：字面标签样文本在标题/正文/单元格/文本框中保持可见。

    文档里的字面 &lt;table&gt;/&lt;time&gt;/&lt;a&gt;/&lt;!-- --&gt;/&lt;? ?&gt;/
    &lt;![CDATA[]]&gt; 等文本不是真实 HTML 标签：不得截断正文（BUG-040）、
    不得被删改写（BUG-044）、不得因实体保护正则缺口在渲染后消失
    （BUG-045）、文本框拼接同样要走实体化（BUG-046）。
    """

    @classmethod
    def setUpClass(cls):
        cls.convert = load_convert_module()

    def test_anchored_heading_literal_table_keeps_rest_of_document(self):
        # BUG-040 红证据：修复前标题中的字面 <table> 让该标题之后到文末
        # 的全部内容被当作一张未闭合表格吞掉
        html = (
            "<p><a id='heading_1'></a>2.1 槽位 &lt;table&gt; 配置</p>"
            "<p>后续正文A。</p>"
            "<p>后续正文B。</p>"
        )
        markdown = self.convert.html_to_markdown(html, {"heading_1": 2})
        self.assertIn("# 2.1 槽位 &lt;table&gt; 配置", markdown)
        self.assertIn("后续正文A。", markdown)
        self.assertIn("后续正文B。", markdown)

    def test_anchored_heading_literal_table_keeps_real_table_around(self):
        # 字面 <table> 标题 + 后接真实表格：表格前后的正文与表格本身都保留
        html = (
            "<p><a id='heading_1'></a>2.1 槽位 &lt;table&gt; 配置</p>"
            "<p>表格前的正文。</p>"
            "<table><tr><td>列A</td><td>列B</td></tr><tr><td>1</td><td>2</td></tr></table>"
            "<p>表格后的正文。</p>"
        )
        markdown = self.convert.html_to_markdown(html, {"heading_1": 2})
        self.assertIn("表格前的正文。", markdown)
        self.assertIn("| 列A | 列B |", markdown)
        self.assertIn("| 1 | 2 |", markdown)
        self.assertIn("表格后的正文。", markdown)

    def test_anchored_heading_literal_tags_not_rewritten_or_deleted(self):
        # BUG-044 红证据：修复前 <a href=x>y</a> 被改写成真实链接 [y](x)，
        # <time>/<div>x</div> 被当作真实标签删除
        for literal, expected in (
            ("&lt;a href=x&gt;y&lt;/a&gt;", "&lt;a href=x&gt;y&lt;/a&gt;"),
            ("&lt;time&gt;", "&lt;time&gt;"),
            ("&lt;div&gt;x&lt;/div&gt;", "&lt;div&gt;x&lt;/div&gt;"),
            ("&lt;TABLE&gt;", "&lt;TABLE&gt;"),
        ):
            with self.subTest(literal=literal):
                html = f"<p><a id='heading_1'></a>2.1 槽位 {literal} 配置</p>"
                markdown = self.convert.html_to_markdown(html, {"heading_1": 2})
                self.assertIn(f"# 2.1 槽位 {expected} 配置", markdown)
                # 不得出现被改写后的真实链接语法
                self.assertNotIn("](x)", markdown)

    def test_replace_html_tables_unclosed_tag_does_not_swallow_rest(self):
        # replace_html_tables 的独立判别：未配对的 <table> 标签按字面放行，
        # 不消费其后内容（BUG-040 的防御深度层）
        html = "前文<table>后续未闭合内容"
        self.assertEqual(self.convert.replace_html_tables(html), html)
        html2 = "前文<table>中间<tail>尾部"
        self.assertEqual(self.convert.replace_html_tables(html2), html2)

    def test_comment_pi_cdata_kept_as_entities_in_paragraphs_and_cells(self):
        # BUG-045 红证据：修复前 <!-- -->、<? ?>、<![CDATA[]]> 不在实体
        # 保护正则内，最终 Markdown 里以裸标签存在，渲染后文字消失。
        # mammoth 对字面文本输出实体形式（&lt;!-- --&gt;），管线输入即如此
        for literal in ("&lt;!-- 隐藏 --&gt;", "&lt;?php echo 1?&gt;", "&lt;![CDATA[abc]]&gt;"):
            html = f"<p>正文 {literal} 结束。</p>"
            markdown = self.convert.html_to_markdown(html)
            self.assertIn(literal, markdown)

        cell_html = (
            "<table><tr><td>标题</td><td>正文 &lt;!-- 隐藏 --&gt; 结束。</td></tr></table>"
        )
        markdown = self.convert.html_to_markdown(cell_html)
        self.assertIn("&lt;!-- 隐藏 --&gt;", markdown)

    def test_non_ascii_tag_names_escaped_consistently(self):
        # BUG-048：非 ASCII 标签首字符（中文槽位标签 <采暖>）与 ASCII
        # 标签同样实体化——修复前以裸标签进入 Markdown（Python-Markdown
        # 恰好会自动转义故可见，但 md 源风格不一致，且依赖渲染器智能）
        html = "<p>正文 &lt;采暖&gt; 结束。</p>"
        markdown = self.convert.html_to_markdown(html)
        self.assertIn("正文 &lt;采暖&gt; 结束。", markdown)

        cell_html = ("<table><tr><td>槽位</td>"
                     "<td>（预约丨设置）&lt;time&gt;[的]&lt;采暖&gt;</td></tr></table>")
        markdown = self.convert.html_to_markdown(cell_html)
        self.assertIn("（预约丨设置）&lt;time&gt;[的]&lt;采暖&gt;", markdown)

        # Excel 单元格路径（纯文本序列化层）同样实体化
        excel_cell = self.convert._normalize_markdown_cell_text("正文 <采暖> 结束。")
        self.assertIn("正文 &lt;采暖&gt; 结束。", excel_cell)

    def test_plain_comparison_text_with_spaces_not_escaped(self):
        # 空白首字符的 "< b" 不是标签样文本：保持裸形式维持 md 源可读性
        # （转义本身渲染无损，但会降低源文件可读性）
        html = "<p>a &lt; b 且 b &gt; c</p>"
        markdown = self.convert.html_to_markdown(html)
        self.assertIn("a < b 且 b > c", markdown)
        excel_cell = self.convert._normalize_markdown_cell_text("a < b 且 b > c")
        self.assertIn("a < b 且 b > c", excel_cell)

    def test_textbox_literal_tags_escaped_in_markdown(self):
        # BUG-046 红证据：修复前文本框内容绕过实体保护管线，字面 <time>
        # 以裸标签进入 Markdown，渲染时不可见
        ns = (
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
        )
        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document.main+xml"/></Types>'
        )
        document = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {ns}><w:body>'
            '<w:p><w:r><w:t xml:space="preserve">正文一。</w:t></w:r></w:p>'
            '<w:p><w:r><w:drawing><wps:wsp><wps:txbx><w:txbxContent>'
            '<w:p><w:r><w:t xml:space="preserve">正文 &lt;time&gt; 结束。</w:t></w:r></w:p>'
            '</w:txbxContent></wps:txbx></wps:wsp></w:drawing></w:r></w:p>'
            '</w:body></w:document>'
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            import zipfile
            path = os.path.join(tmpdir, "textbox.docx")
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("[Content_Types].xml", content_types)
                zf.writestr(
                    "_rels/.rels",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/'
                    'package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/officeDocument" '
                    'Target="word/document.xml"/></Relationships>')
                zf.writestr(
                    "word/_rels/document.xml.rels",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/'
                    'package/2006/relationships"/>')
                zf.writestr("word/document.xml", document)

            md_path = self.convert.convert_docx_to_markdown(path, tmpdir)
            content = Path(md_path).read_text(encoding="utf-8")

        self.assertIn("正文 &lt;time&gt; 结束。", content)


if __name__ == "__main__":
    unittest.main()
