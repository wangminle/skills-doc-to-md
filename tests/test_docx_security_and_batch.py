"""Issue #1 专项测试：恶意 DOCX 资源耗尽防线、SHA-256 完成标记、
批处理单文档超时与半成品清理。"""

import importlib.util
import os
import signal
import struct
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from unittest import mock


WORKSPACE_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = WORKSPACE_ROOT / "skills" / "docx-to-markdown" / "scripts"
def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def build_minimal_docx(path, media_entries=None, extra_entries=None):
    """构造最小可用 DOCX（含 word/document.xml），可注入 media/其他条目。"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
        zf.writestr(
            "word/document.xml",
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p><w:r><w:t>测试正文</w:t></w:r></w:p></w:body></w:document>',
        )
        for name, data in (media_entries or {}).items():
            zf.writestr(f"word/media/{name}", data)
        for name, data in (extra_entries or {}).items():
            zf.writestr(name, data)


def rewrite_docx(path, comment=b"changed"):
    """重写 ZIP（仅改 comment）使文件字节变化但内容仍为有效 DOCX。"""
    with zipfile.ZipFile(path, "r") as zf:
        entries = [(info.filename, zf.read(info.filename)) for info in zf.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.comment = comment
        for name, data in entries:
            zf.writestr(name, data)


class TestZipSecurityValidation(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_module("convert_docx_module", SCRIPTS_DIR / "convert_docx.py")

    def _make_docx(self, tmpdir, **kwargs):
        path = os.path.join(tmpdir, "input.docx")
        build_minimal_docx(path, **kwargs)
        return path

    def test_security_error_is_value_error_subclass(self):
        # 继承 ValueError 兼容既有处理，同时可被调用方精确区分
        self.assertTrue(issubclass(self.convert.DocxSecurityError, ValueError))

    def test_real_docx_passes_validation(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_docx(tmp)
            with zipfile.ZipFile(path, "r") as zf:
                self.convert.validate_docx_zip_security(zf)  # 不抛异常即通过

    def _assert_rejected(self, tmpdir, patched_limits, **kwargs):
        path = self._make_docx(tmpdir, **kwargs)
        with zipfile.ZipFile(path, "r") as zf:
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, patched_limits):
                with self.assertRaises(self.convert.DocxSecurityError):
                    self.convert.validate_docx_zip_security(zf)

    def test_total_uncompressed_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            self._assert_rejected(
                tmp,
                {"total_uncompressed": 1024 * 1024},
                extra_entries={"big.bin": b"\0" * (2 * 1024 * 1024)},
            )

    def test_entry_uncompressed_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            self._assert_rejected(
                tmp,
                {"entry_uncompressed": 512 * 1024},
                extra_entries={"big.bin": b"\0" * (1024 * 1024)},
            )

    def test_entry_compression_ratio_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            # 100KB 高度可压缩数据，压缩后远小于 1/10
            self._assert_rejected(
                tmp,
                {"entry_ratio": 10},
                extra_entries={"sponge.bin": b"A" * (100 * 1024)},
            )

    def test_total_compression_ratio_limit_with_size_floor(self):
        with tempfile.TemporaryDirectory() as tmp:
            # 压缩后总大小低于门槛时不判定（避免小文件舍入误伤）……
            path = self._make_docx(tmp, extra_entries={"sponge.bin": b"A" * (1024 * 1024)})
            with zipfile.ZipFile(path, "r") as zf:
                with mock.patch.dict(
                    self.convert.DOCX_SECURITY_LIMITS,
                    {"entry_ratio": 10 ** 6, "total_ratio": 10,
                     "total_ratio_min_compressed": 10 * 1024 * 1024},
                ):
                    self.convert.validate_docx_zip_security(zf)  # 通过
            # ……超过门槛且总压缩比超限时拒绝
            with zipfile.ZipFile(path, "r") as zf:
                with mock.patch.dict(
                    self.convert.DOCX_SECURITY_LIMITS,
                    {"entry_ratio": 10 ** 6, "total_ratio": 10,
                     "total_ratio_min_compressed": 512},
                ):
                    with self.assertRaises(self.convert.DocxSecurityError):
                        self.convert.validate_docx_zip_security(zf)

    def test_image_count_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            self._assert_rejected(
                tmp,
                {"image_count": 2},
                media_entries={"image1.png": b"\x89PNG\r\n\x1a\n",
                               "image2.png": b"\x89PNG\r\n\x1a\n",
                               "image3.png": b"\x89PNG\r\n\x1a\n"},
            )

    def test_image_file_size_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            self._assert_rejected(
                tmp,
                {"image_file_size": 1024 * 1024},
                media_entries={"image1.png": b"\x89PNG\r\n\x1a\n" + b"\0" * (2 * 1024 * 1024)},
            )

    def test_embedded_excel_size_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            self._assert_rejected(
                tmp,
                {"embedded_excel_size": 1024 * 1024},
                extra_entries={"word/embeddings/sheet1.xlsx": b"\0" * (2 * 1024 * 1024)},
            )

    def test_conversion_rejects_zip_bomb_before_any_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._make_docx(tmp, extra_entries={"sponge.bin": b"A" * (10 * 1024 * 1024)})
            with mock.patch.dict(self.convert.DOCX_SECURITY_LIMITS, {"entry_ratio": 10}):
                with self.assertRaises(self.convert.DocxSecurityError):
                    self.convert.convert_docx_to_markdown(path, os.path.join(tmp, "out"))
            # 安全拒绝发生在任何输出创建之前
            self.assertFalse(os.path.exists(os.path.join(tmp, "out")))


class TestImagePixelCount(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_module("convert_docx_module", SCRIPTS_DIR / "convert_docx.py")

    def test_png(self):
        ihdr = struct.pack(">II", 20000, 15000)
        data = b"\x89PNG\r\n\x1a\n" + b"\x00\x00\x00\rIHDR" + ihdr + b"\x08\x06\x00\x00\x00"
        self.assertEqual(self.convert.image_pixel_count(data), 20000 * 15000)

    def test_jpeg(self):
        data = (b"\xff\xd8" + b"\xff\xc0" + b"\x00\x11" + b"\x08"
                + struct.pack(">HH", 3000, 4000)  # height, width
                + b"\x03" + b"\x01\x11\x00" * 3 + b"\xff\xd9")
        self.assertEqual(self.convert.image_pixel_count(data), 3000 * 4000)

    def test_gif(self):
        data = b"GIF89a" + struct.pack("<HH", 1000, 2000)
        self.assertEqual(self.convert.image_pixel_count(data), 1000 * 2000)

    def test_bmp(self):
        data = b"BM" + b"\x00" * 16 + struct.pack("<ii", 1000, 2000) + b"\x00" * 8
        self.assertEqual(self.convert.image_pixel_count(data), 1000 * 2000)

    def test_webp_vp8x(self):
        buf = bytearray(34)
        buf[0:4] = b"RIFF"
        buf[8:12] = b"WEBP"
        buf[12:16] = b"VP8X"
        buf[24:27] = (2000 - 1).to_bytes(3, "little")
        buf[27:30] = (3000 - 1).to_bytes(3, "little")
        self.assertEqual(self.convert.image_pixel_count(bytes(buf)), 2000 * 3000)

    def test_tiff(self):
        entries = struct.pack("<H", 2)
        entries += struct.pack("<HHII", 256, 4, 1, 2000)  # ImageWidth
        entries += struct.pack("<HHII", 257, 4, 1, 1500)  # ImageLength
        entries += struct.pack("<I", 0)
        data = b"II*\x00" + struct.pack("<I", 8) + entries
        self.assertEqual(self.convert.image_pixel_count(data), 2000 * 1500)

    def test_unknown_format_returns_none(self):
        self.assertIsNone(self.convert.image_pixel_count(b"\xd7\xcd\xc6\x9a" + b"\x00" * 20))

    def test_pixel_bomb_rejected_during_extraction(self):
        # 头部声明 20000x15000（3 亿像素）的 PNG，默认 5000 万上限即拒绝
        ihdr = struct.pack(">II", 20000, 15000)
        png = b"\x89PNG\r\n\x1a\n" + b"\x00\x00\x00\rIHDR" + ihdr + b"\x08\x06\x00\x00\x00"
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "bomb.docx")
            build_minimal_docx(path, media_entries={"image1.png": png})
            with self.assertRaises(self.convert.DocxSecurityError):
                self.convert.convert_docx_to_markdown(path, os.path.join(tmp, "out"))


class TestSanitizeStem(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_module("convert_docx_module", SCRIPTS_DIR / "convert_docx.py")

    def test_plain_name_unchanged(self):
        self.assertEqual(self.convert.sanitize_stem("设备预约V2.7.2"), "设备预约V2.7.2")

    def test_nfkc_only_normalizes_without_hash_suffix(self):
        # 全角→半角归一化过于普遍（尤其中文文档），不附加 hash；
        # 由此产生的罕见碰撞由 sentinel 源哈希校验兜底
        self.assertEqual(self.convert.sanitize_stem("自研语义VAD（测试）"), "自研语义VAD(测试)")

    def test_forbidden_char_replacement_appends_distinct_hash(self):
        # a:b 与 a_b 均清洗为 a_b，但 hash 后缀不同，不共享输出目录
        self.assertNotEqual(self.convert.sanitize_stem("a:b"), self.convert.sanitize_stem("a_b"))
        self.assertTrue(self.convert.sanitize_stem("a:b").startswith("a_b_"))

    def test_quote_removal_appends_hash(self):
        stemmed = self.convert.sanitize_stem('文档"引用"')
        self.assertTrue(stemmed.startswith("文档引用_"))

    def test_long_name_truncated_with_hash(self):
        long_name = "很长的文档名" * 30
        stemmed = self.convert.sanitize_stem(long_name)
        self.assertLessEqual(len(stemmed), 120)
        self.assertIn("_", stemmed)


class TestConversionSentinel(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_module("convert_docx_module", SCRIPTS_DIR / "convert_docx.py")

    def test_conversion_writes_json_sentinel_with_source_hash(self):
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = os.path.join(tmp, "input.docx")
            build_minimal_docx(docx_path)
            md_path = self.convert.convert_docx_to_markdown(docx_path, tmp)
            out_dir = os.path.dirname(md_path)

            sentinel = self.convert.read_conversion_sentinel(out_dir)
            self.assertIsNotNone(sentinel)
            self.assertEqual(sentinel["source_sha256"], self.convert.sha256_file(docx_path))
            self.assertEqual(sentinel["folder_name"], os.path.basename(out_dir))
            self.assertEqual(sentinel["on_limit"], "reject")

    def test_v016_json_sentinel_defaults_to_reject_policy(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, ".converted").write_text(
                '{"folder_name": "x", "source_sha256": "abc"}', encoding="utf-8"
            )
            sentinel = self.convert.read_conversion_sentinel(tmp)
            self.assertEqual(sentinel["on_limit"], "reject")

    def test_unknown_sentinel_policy_is_invalid(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, ".converted").write_text(
                '{"folder_name": "x", "source_sha256": "abc", "on_limit": "ignore"}',
                encoding="utf-8",
            )
            self.assertIsNone(self.convert.read_conversion_sentinel(tmp))

    def test_legacy_plain_text_sentinel_is_invalid(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, ".converted").write_text("done", encoding="utf-8")
            self.assertIsNone(self.convert.read_conversion_sentinel(tmp))

    def test_corrupt_or_incomplete_sentinel_is_invalid(self):
        for payload in ('{"folder_name": "x"}',  # 缺 source_sha256
                        "not json at all",
                        '["list", "not", "dict"]'):
            with tempfile.TemporaryDirectory() as tmp:
                Path(tmp, ".converted").write_text(payload, encoding="utf-8")
                self.assertIsNone(self.convert.read_conversion_sentinel(tmp))

    def test_missing_sentinel_is_invalid(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertIsNone(self.convert.read_conversion_sentinel(tmp))


class TestBatchConvert(unittest.TestCase):
    def setUp(self):
        self.convert = load_module("convert_docx_module", SCRIPTS_DIR / "convert_docx.py")
        self.batch = load_module("batch_convert_module", SCRIPTS_DIR / "batch_convert.py")
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.src = os.path.join(self.tmp.name, "src")
        self.out = os.path.join(self.tmp.name, "out")
        os.makedirs(self.src)
        self.docx_path = os.path.join(self.src, "input.docx")
        build_minimal_docx(self.docx_path)
        base_name = os.path.splitext(os.path.basename(self.docx_path))[0]
        self.folder_name = self.convert.sanitize_stem(base_name)
        self.target_dir = os.path.join(self.out, self.folder_name)

    def _run_batch(self, **kwargs):
        with self.assertLogs("batch_convert_module", level="INFO") as cm:
            self.batch.batch_convert(self.src, self.out, **kwargs)
        return [line for line in cm.output]

    def test_skip_when_complete_and_source_unchanged(self):
        self._run_batch()
        self.assertTrue(os.path.isfile(os.path.join(self.target_dir, f"{self.folder_name}.md")))

        logs = self._run_batch()
        self.assertTrue(any("源文件未变更，跳过" in line for line in logs))
        self.assertTrue(any("成功 0 个, 跳过 1 个, 失败 0 个" in line for line in logs))

    def test_source_change_triggers_reconversion_without_force(self):
        self._run_batch()
        rewrite_docx(self.docx_path)  # 字节变化，内容仍为有效 DOCX

        logs = self._run_batch()
        self.assertTrue(any("重新转换" in line for line in logs))
        sentinel = self.convert.read_conversion_sentinel(self.target_dir)
        self.assertEqual(sentinel["source_sha256"], self.convert.sha256_file(self.docx_path))

    def test_legacy_sentinel_output_is_reconverted(self):
        self._run_batch()
        Path(self.target_dir, ".converted").write_text("done", encoding="utf-8")  # 旧格式

        logs = self._run_batch()
        self.assertFalse(any("源文件未变更，跳过" in line for line in logs))
        self.assertIsNotNone(self.convert.read_conversion_sentinel(self.target_dir))

    def test_tampered_sentinel_hash_is_reconverted(self):
        self._run_batch()
        sentinel_path = Path(self.target_dir, ".converted")
        sentinel_path.write_text(
            '{"folder_name": "%s", "source_sha256": "%s"}' % (self.folder_name, "0" * 64),
            encoding="utf-8",
        )

        logs = self._run_batch()
        self.assertFalse(any("源文件未变更，跳过" in line for line in logs))

    def test_failed_conversion_cleans_half_finished_output(self):
        os.makedirs(self.target_dir)
        Path(self.target_dir, "half.md").write_text("半成品", encoding="utf-8")

        with mock.patch.object(
            self.batch, "convert_docx_to_markdown", side_effect=RuntimeError("boom")
        ):
            logs = self._run_batch()

        self.assertTrue(any("失败: boom" in line for line in logs))
        self.assertFalse(os.path.exists(self.target_dir))  # 半成品被清理

    def test_security_rejection_counts_as_failure_and_cleans_output(self):
        os.makedirs(self.target_dir)
        # 注意用 batch 命名空间中的异常类：batch 模块经 sys.path 导入的
        # convert_docx 与本测试 spec 加载的是不同实例，isinstance 不互通
        with mock.patch.object(
            self.batch,
            "convert_docx_to_markdown",
            side_effect=self.batch.DocxSecurityError("zip bomb"),
        ):
            logs = self._run_batch()

        self.assertTrue(any("安全拒绝" in line for line in logs))
        self.assertTrue(any("失败 1 个" in line for line in logs))
        self.assertFalse(os.path.exists(self.target_dir))

    def test_timeout_kills_hanging_conversion_and_cleans_output(self):
        if not hasattr(signal, "SIGALRM"):
            self.skipTest("平台无 SIGALRM（Windows），超时路径自动跳过")

        def slow_convert(*args, **kwargs):
            time.sleep(30)
            return "never"

        os.makedirs(self.target_dir)
        with mock.patch.object(self.batch, "convert_docx_to_markdown", side_effect=slow_convert):
            logs = self._run_batch(timeout=1)

        self.assertTrue(any("超时" in line for line in logs))
        self.assertFalse(os.path.exists(self.target_dir))

    def test_run_with_timeout_restores_previous_handler(self):
        if not hasattr(signal, "SIGALRM"):
            self.skipTest("平台无 SIGALRM")

        old_handler = signal.getsignal(signal.SIGALRM)
        try:
            self.assertEqual(self.batch._run_with_timeout(lambda: "ok", 30), "ok")
            self.assertEqual(signal.getsignal(signal.SIGALRM), old_handler)
            self.assertEqual(signal.alarm(0), 0)  # 无未决 alarm
        finally:
            signal.signal(signal.SIGALRM, old_handler)

    def test_run_with_timeout_disabled_for_non_positive(self):
        self.assertEqual(self.batch._run_with_timeout(lambda: "ok", 0), "ok")
        self.assertEqual(self.batch._run_with_timeout(lambda: "ok", -1), "ok")


class TestBoundedRead(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_module("convert_docx_module", SCRIPTS_DIR / "convert_docx.py")

    def test_bounded_read_rejects_actual_decompression_over_limit(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.docx")
            build_minimal_docx(path, extra_entries={"sponge.bin": b"A" * (2 * 1024 * 1024)})
            with zipfile.ZipFile(path, "r") as zf:
                with self.assertRaises(self.convert.DocxSecurityError):
                    self.convert.read_zip_entry_bounded(zf, "sponge.bin", 1024 * 1024)
                # 上限内正常读取
                data = self.convert.read_zip_entry_bounded(zf, "sponge.bin", 4 * 1024 * 1024)
                self.assertEqual(len(data), 2 * 1024 * 1024)


if __name__ == "__main__":
    unittest.main()
