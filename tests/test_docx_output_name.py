"""Issue #3 专项测试：output_name 自定义输出命名。

验证 output_name 统一经 sanitize_stem 清洗后用于输出目录、Markdown 文件名
与 .converted sentinel 的 folder_name，三处保持一致；默认 None 时行为不变。
"""

import importlib.util
import json
import os
import tempfile
import unittest
import zipfile
from pathlib import Path


WORKSPACE_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = WORKSPACE_ROOT / "skills" / "docx-to-markdown" / "scripts"
def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def build_minimal_docx(path):
    """构造无需私有夹具、可被 mammoth 转换的最小 DOCX。"""
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body><w:p><w:r><w:t>命名测试正文</w:t></w:r></w:p></w:body></w:document>'
    )
    content_types = (
        '<?xml version="1.0"?><Types '
        'xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("word/document.xml", document)


class TestOutputName(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_module("convert_docx_module", SCRIPTS_DIR / "convert_docx.py")
        cls.fixture_dir = tempfile.TemporaryDirectory()
        cls.docx_path = Path(cls.fixture_dir.name, "stored-upload-id.docx")
        build_minimal_docx(cls.docx_path)

    @classmethod
    def tearDownClass(cls):
        cls.fixture_dir.cleanup()

    def test_output_name_drives_folder_md_and_sentinel(self):
        # 目录、.md 文件名、sentinel 的 folder_name 三处使用同一命名
        with tempfile.TemporaryDirectory() as tmp:
            md_path = self.convert.convert_docx_to_markdown(
                str(self.docx_path), tmp, output_name="自定义输出名")
            out_dir = os.path.dirname(md_path)

            self.assertEqual(os.path.basename(out_dir), "自定义输出名")
            self.assertEqual(os.path.basename(md_path), "自定义输出名.md")

            sentinel = json.loads(Path(out_dir, ".converted").read_text(encoding="utf-8"))
            self.assertEqual(sentinel["folder_name"], "自定义输出名")
            # sentinel 仍记录真实源文件哈希，与命名来源解耦
            self.assertEqual(sentinel["source_sha256"], self.convert.sha256_file(str(self.docx_path)))

    def test_output_name_sanitized_same_as_source_name(self):
        # 含非法字符的 output_name 走与源文件名相同的 sanitize_stem 清洗
        # （a:b -> a_b_<hash>，不与其他名称共享输出目录）
        with tempfile.TemporaryDirectory() as tmp:
            md_path = self.convert.convert_docx_to_markdown(
                str(self.docx_path), tmp, output_name="a:b")
            expected = self.convert.sanitize_stem("a:b")
            out_dir = os.path.dirname(md_path)

            self.assertEqual(os.path.basename(out_dir), expected)
            self.assertTrue(os.path.isfile(os.path.join(out_dir, f"{expected}.md")))

    def test_default_output_name_matches_source_stem(self):
        # 默认 None：与既有行为一致，用源文件名命名
        with tempfile.TemporaryDirectory() as tmp:
            md_path = self.convert.convert_docx_to_markdown(str(self.docx_path), tmp)
            expected = self.convert.sanitize_stem(self.docx_path.stem)
            self.assertEqual(os.path.basename(os.path.dirname(md_path)), expected)
            self.assertEqual(os.path.basename(md_path), f"{expected}.md")

    def test_output_name_without_subfolder_names_md_only(self):
        # create_subfolder=False 时仅 .md 文件名使用 output_name
        with tempfile.TemporaryDirectory() as tmp:
            md_path = self.convert.convert_docx_to_markdown(
                str(self.docx_path), tmp, create_subfolder=False, output_name="flat_name")
            self.assertEqual(md_path, os.path.join(tmp, "flat_name.md"))

    def test_blank_output_name_rejected(self):
        # 空白 output_name 属于调用方错误，直接拒绝而非静默回退
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(ValueError):
                self.convert.convert_docx_to_markdown(str(self.docx_path), tmp, output_name="   ")

    def test_output_name_docx_suffix_is_removed(self):
        with tempfile.TemporaryDirectory() as tmp:
            md_path = self.convert.convert_docx_to_markdown(
                str(self.docx_path), tmp, output_name="用户原始名称.DOCX")
            self.assertEqual(os.path.basename(os.path.dirname(md_path)), "用户原始名称")
            self.assertEqual(os.path.basename(md_path), "用户原始名称.md")

    def test_output_name_version_suffix_is_preserved(self):
        with tempfile.TemporaryDirectory() as tmp:
            md_path = self.convert.convert_docx_to_markdown(
                str(self.docx_path), tmp, output_name="需求V2.4")
            self.assertEqual(os.path.basename(md_path), "需求V2.4.md")


if __name__ == "__main__":
    unittest.main()
