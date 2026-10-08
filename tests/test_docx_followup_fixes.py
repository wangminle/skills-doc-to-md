"""BUG-049 / OPT-002：目录占位资源与缺失路径组件的回归验证。"""

import importlib.util
import os
import struct
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock


SCRIPT = Path(__file__).resolve().parents[1] / "skills/docx-to-markdown/scripts/convert_docx.py"


def load_converter():
    spec = importlib.util.spec_from_file_location("convert_followup", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class TestRelatedDirectoryPlaceholder(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_converter()
        cls.png = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + struct.pack(">II", 2, 2)

    def build_docx(self, path, directory="word/media/", reference_directory=False, absolute=True):
        target = "/" + directory if absolute else directory.removeprefix("word/")
        ns = (
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
        )
        ids = ["image"] + (["directory"] if reference_directory else [])
        drawings = "".join(
            '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="picture"/><pic:cNvPicPr/>'
            f'</pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/></pic:blipFill>'
            '<pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline>'
            '</w:drawing></w:r></w:p>' for rid in ids
        )
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(
                "[Content_Types].xml",
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="png" ContentType="image/png"/>'
                f'<Override PartName="/{directory}" ContentType="image/png"/></Types>',
            )
            zf.writestr(
                "word/document.xml", f'<w:document {ns}><w:body>{drawings}'
                '<w:p><w:r><w:t>正文保留</w:t></w:r></w:p></w:body></w:document>',
            )
            zf.writestr(
                "word/_rels/document.xml.rels",
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="directory" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                f'Target="{target}"/>'
                '<Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                'Target="media/image1.png"/></Relationships>',
            )
            # 目录排在真图前，验证不会占用唯一图片配额。
            zf.writestr(directory, b"")
            zf.writestr("word/media/image1.png", self.png)

    def test_related_empty_directories_not_counted(self):
        with tempfile.TemporaryDirectory() as tmp:
            cases = [(directory, True) for directory in (
                "word/media/", "word/media/nested/", "custom/", "word/media/empty.png/"
            )] + [("word/media/", False), ("word/media/nested/", False)]
            for directory, absolute in cases:
                with self.subTest(directory=directory, absolute=absolute):
                    path = Path(tmp, "a.docx")
                    self.build_docx(path, directory, absolute=absolute)
                    with zipfile.ZipFile(path) as zf:
                        self.assertEqual(self.convert._document_image_part_names(zf), {"word/media/image1.png"})

    def test_directory_relation_does_not_write_empty_asset_or_consume_quota(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "a.docx")
            self.build_docx(path)
            for policy in ("reject", "skip"):
                with self.subTest(policy=policy), mock.patch.dict(
                    self.convert.DOCX_SECURITY_LIMITS, {"image_count": 1}
                ):
                    md = Path(self.convert.convert_docx_to_markdown(str(path), str(Path(tmp, policy)), on_limit=policy))
                    assets = list((md.parent / "assets").iterdir())
                    self.assertEqual([p.name for p in assets], ["image1.png"])
                    self.assertEqual(assets[0].read_bytes(), self.png)
                    content = md.read_text()
                    self.assertIn("assets/image1.png", content)
                    self.assertIn("正文保留", content)
                    self.assertNotIn("图片数量超过上限", content)

    def test_referenced_directory_does_not_revive_via_callback(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp, "a.docx")
            self.build_docx(path, reference_directory=True)
            for policy in ("reject", "skip"):
                with self.subTest(policy=policy), mock.patch.dict(
                    self.convert.DOCX_SECURITY_LIMITS, {"image_count": 1}
                ):
                    md = Path(self.convert.convert_docx_to_markdown(str(path), str(Path(tmp, policy)), on_limit=policy))
                    self.assertEqual([p.name for p in (md.parent / "assets").iterdir()], ["image1.png"])
                    content = md.read_text()
                    self.assertIn("assets/image1.png", content)
                    self.assertIn("图片关系指向空目录", content)
                    self.assertNotIn("图片数量超过上限", content)


@unittest.skipUnless(os.name == "posix", "POSIX 路径解析与符号链接验证")
class TestMissingComponentRealpath(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.convert = load_converter()

    def test_missing_component_parent_normalization_and_later_symlink(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "real/sub").mkdir(parents=True)
            os.symlink("real/sub", root / "link")
            os.symlink("absent/../link", root / "dangling")
            for suffix in (
                "absent/../other", "absent/deeper/../../link",
                "absent/../link/..", "dangling/tail", "absent/../../../../other",
            ):
                path = str(root / suffix)
                with self.subTest(path=path):
                    self.assertEqual(self.convert._safe_realpath(path), os.path.realpath(path))

    def test_timeout_after_missing_component_is_not_swallowed(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = str(Path(tmp) / "absent/../later")
            original = self.convert._path_status

            def timeout_on_later(current):
                if current.endswith("/later"):
                    raise TimeoutError("单文档转换超时")
                return original(current)

            with mock.patch.object(self.convert, "_path_status", side_effect=timeout_on_later):
                with self.assertRaises(TimeoutError):
                    self.convert._safe_realpath(path)

    def test_symlink_cycles_allow_resolving_remaining_components(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "real/sub").mkdir(parents=True)
            os.symlink("real/sub", root / "link")
            os.symlink("loop", root / "loop")
            os.symlink("pair2", root / "pair1")
            os.symlink("pair1", root / "pair2")
            for suffix in ("loop/../link", "pair1/../link", "link/../../pair1"):
                path = str(root / suffix)
                with self.subTest(path=path):
                    self.assertEqual(self.convert._safe_realpath(path), os.path.realpath(path))

    def test_acyclic_long_symlink_chain_matches_realpath(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "real/sub").mkdir(parents=True)
            for i in range(60):
                os.symlink(f"link{i + 1}" if i < 59 else "real/sub", root / f"link{i}")
            path = str(root / "link0/tail")
            self.assertEqual(self.convert._safe_realpath(path), os.path.realpath(path))
