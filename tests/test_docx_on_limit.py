"""Issue #2 专项测试：on_limit 资源超限降级跳过策略。

核心验收点：
- on_limit="reject"（默认）完全保持既有行为：四类可降级资源超限抛 DocxSecurityError；
- on_limit="skip" 只跳过具体资源，正文与其余资源保留，所有读取仍有实际字节上限；
- ZIP bomb 等恶意特征在 skip 模式下依旧整篇拒绝；
- mammoth 回调执行同一套防线与图片数量配额，超限图片不会经回调兜底写盘复活。
"""

import importlib.util
import io
import os
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock


WORKSPACE_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = WORKSPACE_ROOT / "skills" / "docx-to-markdown" / "scripts"
CONVERT_SCRIPT = SCRIPTS_DIR / "convert_docx.py"


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def tiny_png(width=2, height=2):
    """头部声明 width x height 的最小 PNG 片段（无需为合法完整图片）。"""
    ihdr = struct.pack(">II", width, height)
    return b"\x89PNG\r\n\x1a\n" + b"\x00\x00\x00\rIHDR" + ihdr + b"\x08\x06\x00\x00\x00"


def pixel_bomb_png():
    """20000x15000（3 亿像素）PNG 头，触发默认 5000 万像素上限。"""
    return tiny_png(20000, 15000)


def build_minimal_docx(path, media_entries=None, extra_entries=None):
    """构造最小可用 DOCX（含 word/document.xml），可注入 media/其他条目。"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        zf.writestr("word/document.xml", "<w:document/>")
        for name, data in (media_entries or {}).items():
            zf.writestr(f"word/media/{name}", data)
        for name, data in (extra_entries or {}).items():
            zf.writestr(name, data)


_DOC_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)


def build_docx_with_images(path, image_bytes_list, text="正文保留测试", reference_ids=None):
    """构造正文引用了 media 图片的最小 DOCX，使 mammoth 触发 convert_image 回调。"""
    if reference_ids is None:
        reference_ids = list(range(1, len(image_bytes_list) + 1))
    paragraphs = []
    for drawing_id, relationship_id in enumerate(reference_ids, 1):
        paragraphs.append(
            f"<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"100\"/>"
            f"<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
            f"<pic:pic><pic:nvPicPr><pic:cNvPr id=\"{drawing_id}\" name=\"img{drawing_id}\"/>"
            f"<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>"
            f"<a:blip r:embed=\"rId{relationship_id}\"/></pic:blipFill><pic:spPr/></pic:pic>"
            f"</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
        )
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<w:document {_DOC_NS}><w:body>"
        + "".join(paragraphs)
        + f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"
        "</w:body></w:document>"
    )
    rels = ['<?xml version="1.0"?><Relationships '
            'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i, _ in enumerate(image_bytes_list, 1):
        rels.append(
            f'<Relationship Id="rId{i}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/image{i}.png"/>'
        )
    rels.append("</Relationships>")
    content_types = (
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType='
        '"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/_rels/document.xml.rels", "".join(rels))
        for i, data in enumerate(image_bytes_list, 1):
            zf.writestr(f"word/media/image{i}.png", data)


class TestOnLimitPolicy(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_module("convert_docx_module", CONVERT_SCRIPT)

    def test_resource_limit_error_is_security_error_subclass(self):
        # reject 模式下两类异常处置一致：ResourceLimitExceeded ⊂ DocxSecurityError
        self.assertTrue(issubclass(self.convert.ResourceLimitExceeded,
                                   self.convert.DocxSecurityError))
        self.assertTrue(issubclass(self.convert.ResourceLimitExceeded, ValueError))

    def test_reject_default_and_explicit_raise(self):
        # 像素炸弹：默认与显式 reject 均整篇拒绝
        for kwargs in ({}, {"on_limit": "reject"}):
            with self.subTest(kwargs=kwargs):
                with tempfile.TemporaryDirectory() as tmp:
                    path = os.path.join(tmp, "bomb.docx")
                    build_minimal_docx(path, media_entries={"image1.png": pixel_bomb_png()})
                    with self.assertRaises(self.convert.DocxSecurityError):
                        self.convert.convert_docx_to_markdown(
                            path, os.path.join(tmp, "out"), **kwargs)

    def test_invalid_on_limit_value_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_docx_with_images(path, [])
            with self.assertRaises(ValueError):
                self.convert.convert_docx_to_markdown(path, tmp, on_limit="ignore")

    def test_public_helpers_reject_invalid_on_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_docx_with_images(path, [])
            with zipfile.ZipFile(path, "r") as zf:
                with self.assertRaises(ValueError):
                    self.convert.validate_docx_zip_security(zf, on_limit="ignore")
            with self.assertRaises(ValueError):
                self.convert.extract_content_from_docx(
                    path, os.path.join(tmp, "assets"), on_limit="ignore")

    def test_extract_content_keeps_three_item_return_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            assets_dir = os.path.join(tmp, "assets")
            os.makedirs(assets_dir)
            build_docx_with_images(path, [])

            result = self.convert.extract_content_from_docx(path, assets_dir)

            self.assertEqual(len(result), 3)

    def test_validate_skip_mode_tolerates_degradable_metadata_only(self):
        # validate 的 on_limit 分层：skip 模式容忍可降级项（图片数量）的元数据超限，
        # 但 ZIP 级恶意特征（压缩比）在 skip 模式下仍拒绝
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_minimal_docx(path, media_entries={f"image{i}.png": tiny_png() for i in range(3)})
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"image_count": 2}):
                with zipfile.ZipFile(path, "r") as zf:
                    with self.assertRaises(self.convert.DocxSecurityError):
                        self.convert.validate_docx_zip_security(zf)  # 默认 reject
                    self.convert.validate_docx_zip_security(zf, on_limit="skip")  # 数量超限不抛

            # 压缩比超限（zip bomb 特征）：skip 模式同样拒绝
            bomb = os.path.join(tmp, "bomb.docx")
            build_minimal_docx(bomb, extra_entries={"sponge.bin": b"A" * (10 * 1024 * 1024)})
            with zipfile.ZipFile(bomb, "r") as zf:
                with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"entry_ratio": 10}):
                    with self.assertRaises(self.convert.DocxSecurityError):
                        self.convert.validate_docx_zip_security(zf, on_limit="skip")

    def test_zip_bomb_rejected_even_in_skip_mode(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bomb.docx")
            build_minimal_docx(path, extra_entries={"sponge.bin": b"A" * (10 * 1024 * 1024)})
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"entry_ratio": 10}):
                with self.assertRaises(self.convert.DocxSecurityError):
                    self.convert.convert_docx_to_markdown(
                        path, os.path.join(tmp, "out"), on_limit="skip")
            # 恶意特征拒绝发生在任何输出创建之前
            self.assertFalse(os.path.exists(os.path.join(tmp, "out")))

    def test_duplicate_zip_entry_name_is_always_rejected(self):
        buffer = io.BytesIO()
        with self.assertWarns(UserWarning):
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("word/media/same.png", tiny_png())
                zf.writestr("word/media/same.png", tiny_png(3, 3))

        with zipfile.ZipFile(io.BytesIO(buffer.getvalue()), "r") as zf:
            with self.assertRaisesRegex(self.convert.DocxSecurityError, "重复条目"):
                self.convert.validate_docx_zip_security(zf, on_limit="skip")

    def test_skip_mode_pixel_bomb_body_kept_and_image_not_written(self):
        # 旁路封堵核心用例：正文引用像素炸弹图片，skip 模式转出正文，
        # 原位置显示跳过说明，超限图片绝不落盘（含 mammoth 回调兜底路径）
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bomb.docx")
            build_docx_with_images(path, [pixel_bomb_png()])
            with self.assertLogs("convert_docx_module", level="WARNING") as logs:
                md_path = self.convert.convert_docx_to_markdown(
                    path, os.path.join(tmp, "out"), on_limit="skip")

            markdown = Path(md_path).read_text(encoding="utf-8")
            self.assertIn("正文保留测试", markdown)
            self.assertIn("【图片已跳过：单图像素超过上限】", markdown)
            assets_dir = os.path.join(os.path.dirname(md_path), "assets")
            self.assertEqual(os.listdir(assets_dir), [])  # 无任何超限图片落盘
            self.assertTrue(any("跳过" in line for line in logs.output))

    def test_skip_mode_over_size_image_skipped(self):
        # 单图超过大小上限（patch 至 1MB，图片用不可压缩随机数据避免误触压缩比防线）
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            big_image = tiny_png() + os.urandom(2 * 1024 * 1024)
            build_docx_with_images(path, [big_image])
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS,
                                 {"image_file_size": 1024 * 1024}):
                with self.assertRaises(self.convert.DocxSecurityError):
                    self.convert.convert_docx_to_markdown(path, os.path.join(tmp, "out-reject"))

                md_path = self.convert.convert_docx_to_markdown(
                    path, os.path.join(tmp, "out"), on_limit="skip")

            markdown = Path(md_path).read_text(encoding="utf-8")
            self.assertIn("正文保留测试", markdown)
            self.assertIn("【图片已跳过：单图超过大小上限】", markdown)
            self.assertEqual(os.listdir(os.path.join(os.path.dirname(md_path), "assets")), [])

    def test_skip_reconversion_removes_stale_previously_extracted_image(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "same.docx")
            out_dir = os.path.join(tmp, "out")
            build_docx_with_images(path, [tiny_png(2, 2)])
            first_md = self.convert.convert_docx_to_markdown(path, out_dir)
            assets_dir = os.path.join(os.path.dirname(first_md), "assets")
            self.assertEqual(os.listdir(assets_dir), ["image1.png"])

            build_docx_with_images(path, [pixel_bomb_png()])
            second_md = self.convert.convert_docx_to_markdown(
                path, out_dir, on_limit="skip"
            )

            self.assertIn(
                "【图片已跳过：单图像素超过上限】",
                Path(second_md).read_text(encoding="utf-8"),
            )
            self.assertEqual(os.listdir(assets_dir), [])

    def test_flat_output_preserves_assets_not_owned_by_current_document(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            output_dir = os.path.join(tmp, "out")
            assets_dir = os.path.join(output_dir, "assets")
            os.makedirs(assets_dir)
            unrelated = Path(assets_dir) / "another-document.png"
            unrelated.write_bytes(tiny_png(7, 7))
            build_docx_with_images(path, [tiny_png(2, 2)])

            self.convert.convert_docx_to_markdown(
                path, output_dir, create_subfolder=False, on_limit="skip"
            )

            self.assertTrue(unrelated.exists())

    def test_assets_symlink_is_rejected_before_conversion(self):
        if not hasattr(os, "symlink"):
            self.skipTest("当前平台不支持符号链接")
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            output_dir = Path(tmp) / "out" / "input"
            external_dir = Path(tmp) / "external"
            output_dir.mkdir(parents=True)
            external_dir.mkdir()
            (external_dir / "keep.txt").write_text("keep", encoding="utf-8")
            os.symlink(external_dir, output_dir / "assets")
            build_docx_with_images(path, [tiny_png(2, 2)])

            with self.assertRaisesRegex(ValueError, "assets 目录"):
                self.convert.convert_docx_to_markdown(
                    path, os.path.join(tmp, "out"), on_limit="skip"
                )
            self.assertEqual((external_dir / "keep.txt").read_text(encoding="utf-8"), "keep")

    def test_skip_mode_embedded_excel_over_size_skipped(self):
        # 嵌入 Excel 超限（patch 至 1MB）：reject 拒绝；skip 跳过表格、正文保留
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_docx_with_images(path, [])
            with zipfile.ZipFile(path, "a", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("word/embeddings/sheet1.xlsx", os.urandom(2 * 1024 * 1024))
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS,
                                 {"embedded_excel_size": 1024 * 1024}):
                with self.assertRaises(self.convert.DocxSecurityError):
                    self.convert.convert_docx_to_markdown(path, os.path.join(tmp, "out-reject"))

                md_path = self.convert.convert_docx_to_markdown(
                    path, os.path.join(tmp, "out"), on_limit="skip")

            markdown = Path(md_path).read_text(encoding="utf-8")
            self.assertIn("正文保留测试", markdown)

    def test_skip_mode_image_count_quota_blocks_callback_fallback(self):
        # 图片数量配额（patch 至 1 张）：第二张图片在提取与回调兜底两条路径
        # 都被同一配额拦截，只有第一张落盘
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_docx_with_images(path, [tiny_png(2, 2), tiny_png(3, 3)])
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"image_count": 1}):
                md_path = self.convert.convert_docx_to_markdown(
                    path, os.path.join(tmp, "out"), on_limit="skip")

            markdown = Path(md_path).read_text(encoding="utf-8")
            self.assertIn("正文保留测试", markdown)
            self.assertEqual(markdown.count("![]("), 1)  # 仅第一张以图片输出
            self.assertIn("【图片已跳过：图片数量超过上限】", markdown)
            assets_dir = os.path.join(os.path.dirname(md_path), "assets")
            self.assertEqual(os.listdir(assets_dir), ["image1.png"])

    def test_skip_mode_image_count_records_one_bounded_summary(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            assets_dir = os.path.join(tmp, "assets")
            os.makedirs(assets_dir)
            build_docx_with_images(
                path, [tiny_png(2, 2), tiny_png(3, 3), tiny_png(4, 4), tiny_png(5, 5)]
            )
            skip_state = {}

            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"image_count": 1}):
                self.convert.extract_content_from_docx(
                    path, assets_dir, on_limit="skip", skip_state=skip_state
                )

            count_records = [
                item for item in skip_state["skipped"] if "图片数量超过上限" in item[1]
            ]
            self.assertEqual(len(count_records), 1)

    def test_full_conversion_image_count_keeps_one_summary(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_docx_with_images(
                path, [tiny_png(2, 2), tiny_png(3, 3), tiny_png(4, 4), tiny_png(5, 5)]
            )
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"image_count": 1}):
                with self.assertLogs("convert_docx_module", level="WARNING") as logs:
                    md_path = self.convert.convert_docx_to_markdown(
                        path, os.path.join(tmp, "out"), on_limit="skip"
                    )

            summary_lines = [line for line in logs.output if "共跳过" in line]
            self.assertEqual(len(summary_lines), 1)
            self.assertIn("共跳过 1 项", summary_lines[0])
            markdown = Path(md_path).read_text(encoding="utf-8")
            self.assertEqual(markdown.count("【图片已跳过：图片数量超过上限】"), 3)

    def test_repeated_references_to_one_physical_image_do_not_consume_quota(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_docx_with_images(
                path, [tiny_png(2, 2)], reference_ids=[1, 1, 1, 1]
            )
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"image_count": 1}):
                md_path = self.convert.convert_docx_to_markdown(
                    path, os.path.join(tmp, "out"), on_limit="skip"
                )

            markdown = Path(md_path).read_text(encoding="utf-8")
            self.assertEqual(markdown.count("![](assets/image1.png)"), 4)
            self.assertNotIn("图片数量超过上限", markdown)

    def test_callback_reuses_zip_selected_media_quota(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            # ZIP 顺序中 image1 占用唯一配额，正文却只引用 image2。
            # 回调不得重新分配一份配额把 image2 兜底落盘。
            build_docx_with_images(
                path, [tiny_png(2, 2), tiny_png(3, 3)], reference_ids=[2]
            )
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"image_count": 1}):
                md_path = self.convert.convert_docx_to_markdown(
                    path, os.path.join(tmp, "out"), on_limit="skip"
                )

            markdown = Path(md_path).read_text(encoding="utf-8")
            self.assertIn("【图片已跳过：图片数量超过上限】", markdown)
            # 只保留 ZIP 阶段选中的 image1；回调不得再落盘 image2。
            self.assertEqual(os.listdir(Path(md_path).parent / "assets"), ["image1.png"])

    def test_image_count_counts_oversized_images_before_reading_remaining_callbacks(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_docx_with_images(path, [pixel_bomb_png()] * 4)

            original_pixel_count = self.convert.image_pixel_count
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"image_count": 1}):
                with mock.patch.object(
                    self.convert, "image_pixel_count", wraps=original_pixel_count
                ) as pixel_counter:
                    with self.assertLogs("convert_docx_module", level="WARNING") as logs:
                        md_path = self.convert.convert_docx_to_markdown(
                            path, os.path.join(tmp, "out"), on_limit="skip"
                        )

            # 配额内第一张在 ZIP 提取与 Mammoth 回调各检查一次；
            # 后三张在两条路径中都不应再打开/解析。
            self.assertEqual(pixel_counter.call_count, 2)
            count_records = [line for line in logs.output if "已跳过" in line and "图片数量超过上限" in line]
            self.assertEqual(len(count_records), 1)
            markdown = Path(md_path).read_text(encoding="utf-8")
            self.assertEqual(markdown.count("【图片已跳过：图片数量超过上限】"), 3)

    def test_nested_xlsx_zip_bomb_rejected_in_skip_mode(self):
        nested = io.BytesIO()
        with zipfile.ZipFile(nested, "w", zipfile.ZIP_DEFLATED) as xlsx_zip:
            xlsx_zip.writestr("xl/worksheets/sheet1.xml", b"A" * (2 * 1024 * 1024))

        # 直接验证 XLSX 内层 ZIP，避免 DOCX 外层压缩比检查
        # 先拦截 word/embeddings/*.xlsx 而造成假阳性。
        with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"entry_ratio": 10}):
            with self.assertRaises(self.convert.DocxSecurityError):
                self.convert.excel_to_markdown(nested.getvalue())

        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_docx_with_images(path, [])
            with zipfile.ZipFile(path, "a", zipfile.ZIP_DEFLATED) as docx_zip:
                docx_zip.writestr("word/embeddings/bomb.xlsx", nested.getvalue())

            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"entry_ratio": 10}):
                with self.assertRaises(self.convert.DocxSecurityError):
                    self.convert.convert_docx_to_markdown(
                        path, os.path.join(tmp, "out"), on_limit="skip"
                    )


class TestOnLimitCLI(unittest.TestCase):
    """CLI 入口：--on-limit 默认 reject（exit 2），skip 正常转出（exit 0）。"""

    def test_cli_on_limit_flag(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bomb.docx")
            build_docx_with_images(path, [pixel_bomb_png()])

            reject_out = os.path.join(tmp, "out-reject")
            result = subprocess.run(
                [sys.executable, str(CONVERT_SCRIPT), path, reject_out],
                capture_output=True, text=True)
            self.assertEqual(result.returncode, 2)
            # 单文件 reject 中途失败可能留下空输出目录（清理职责在批处理），
            # 但不得产出任何 Markdown
            self.assertTrue(os.listdir(reject_out) == [] or
                            not any(f.endswith(".md") for f in os.listdir(reject_out)))

            skip_out = os.path.join(tmp, "out")
            result = subprocess.run(
                [sys.executable, str(CONVERT_SCRIPT), path, skip_out, "--on-limit", "skip"],
                capture_output=True, text=True)
            self.assertEqual(result.returncode, 0)
            md_path = Path(skip_out, "bomb", "bomb.md")
            self.assertTrue(md_path.is_file())
            self.assertIn("【图片已跳过：单图像素超过上限】", md_path.read_text(encoding="utf-8"))


class TestBatchOnLimit(unittest.TestCase):
    def setUp(self):
        self.convert = load_module("convert_docx_module", CONVERT_SCRIPT)
        self.batch = load_module("batch_convert_module", SCRIPTS_DIR / "batch_convert.py")
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.src = os.path.join(self.tmp.name, "src")
        self.out = os.path.join(self.tmp.name, "out")
        os.makedirs(self.src)
        self.docx_path = os.path.join(self.src, "bomb.docx")
        build_docx_with_images(self.docx_path, [pixel_bomb_png()])
        self.target_dir = os.path.join(self.out, "bomb")

    def _run_batch(self, **kwargs):
        with self.assertLogs("batch_convert_module", level="INFO") as cm:
            self.batch.batch_convert(self.src, self.out, **kwargs)
        return cm.output

    def test_batch_default_rejects_over_limit_doc(self):
        logs = self._run_batch()
        self.assertTrue(any("安全拒绝" in line for line in logs))
        self.assertTrue(any("失败 1 个" in line for line in logs))
        self.assertFalse(os.path.exists(self.target_dir))  # 半成品被清理

    def test_batch_skip_mode_converts_over_limit_doc(self):
        logs = self._run_batch(on_limit="skip")
        self.assertTrue(any("成功 1 个" in line for line in logs))
        md_path = Path(self.target_dir, "bomb.md")
        self.assertTrue(md_path.is_file())
        self.assertIn("【图片已跳过：单图像素超过上限】", md_path.read_text(encoding="utf-8"))

    def test_batch_policy_change_does_not_reuse_skip_sentinel(self):
        self._run_batch(on_limit="skip")

        logs = self._run_batch(on_limit="reject")

        self.assertFalse(any("源文件未变更，跳过" in line for line in logs))
        self.assertTrue(any("失败 1 个" in line for line in logs))
        self.assertFalse(os.path.exists(self.target_dir))

    def test_batch_api_rejects_invalid_on_limit(self):
        with self.assertRaises(ValueError):
            self.batch.batch_convert(self.src, self.out, on_limit="ignore")


if __name__ == "__main__":
    unittest.main()
