import importlib.util
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


if __name__ == "__main__":
    unittest.main()
