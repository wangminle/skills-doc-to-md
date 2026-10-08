#!/usr/bin/env python3
"""
将docx文档转换为markdown格式，并提取所有图片到assets文件夹
支持将嵌入的Excel表格转换为Markdown表格
"""

import hashlib
import json
import logging
import math
import os
import stat
import sys
import zipfile
import re
import io
import unicodedata
from html import escape as _escape_html, unescape
from html.parser import HTMLParser
from collections import defaultdict
import posixpath
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# XML 解析统一入口：安装了 defusedxml 时防御实体膨胀/外部实体等 XML 攻击，
# 未安装自动回退标准库 xml.etree（功能等价，仅防护降级）。
try:
    from defusedxml.ElementTree import fromstring as _safe_xml_fromstring
except ImportError:
    from xml.etree.ElementTree import fromstring as _safe_xml_fromstring


_FORBIDDEN_FILENAME_CHARS_RE = re.compile(r'[\\/:*?"<>|]')
_WHITESPACE_RE = re.compile(r"\s+")
_QUOTE_CHARS = '"“”‘’‚‛„‟«»‹›'


class DocxSecurityError(ValueError):
    """输入 DOCX 触发资源耗尽防线（zip bomb / 超大资源）。

    继承 ValueError 以兼容既有 except ValueError 处理；调用方可精确区分
    “安全拒绝”与普通格式/转换错误——安全拒绝表示输入恶意或异常，
    不可降级重试（若降级重试会绕过防线）。
    """


class ResourceLimitExceeded(DocxSecurityError):
    """可降级资源限制超限（图片数量/单图大小/单图像素/嵌入 Excel 大小）。

    继承 DocxSecurityError：默认（on_limit="reject"）处置与安全拒绝完全
    一致——整篇拒绝、不降级不重试；仅当调用方显式选择 on_limit="skip"
    时，才用于精确捕获并降级为跳过该资源。ZIP bomb 等恶意特征（总量/
    单 entry 解压量/压缩比）始终抛基类 DocxSecurityError，任何模式下
    都不降级。
    """


# 资源耗尽防线阈值。集中为 dict 便于测试注入与按需收紧。
DOCX_SECURITY_LIMITS = {
    "total_uncompressed": 500 * 1024 * 1024,    # ZIP 总解压上限
    "entry_uncompressed": 100 * 1024 * 1024,    # 单 entry 解压上限
    "entry_ratio": 100,                         # 单 entry 压缩比上限
    "total_ratio": 100,                         # 总压缩比上限
    "total_ratio_min_compressed": 1024 * 1024,  # 总压缩比仅对压缩后 >1MB 的包判定
    "image_count": 500,                         # 物理图片条目数量上限（word/media 及关系目标）
    "image_file_size": 20 * 1024 * 1024,        # 单图文件大小上限
    "image_pixels": 50_000_000,                 # 单图像素上限（解压炸弹检测）
    "embedded_excel_size": 50 * 1024 * 1024,    # 嵌入 Excel 大小上限
}

# 批处理跳过判定用的完成标记文件名（JSON，记录源文件 SHA-256）
SENTINEL_FILENAME = ".converted"

# on_limit="skip" 时超限资源，以及两种策略下空目录伪图片的可见说明文案。
# mammoth 回调用 __SKIPPED_IMAGE_<reason>__ 作为 src 占位，转换后统一
# 替换为对应说明，保证跳过在输出中可见且不引用不存在的资源文件。
SKIPPED_IMAGE_NOTE = {
    "size": "单图超过大小上限",
    "pixels": "单图像素超过上限",
    "count": "图片数量超过上限",
    "directory": "图片关系指向空目录",
}


def validate_on_limit(on_limit: str) -> None:
    """校验公开 API 共用的资源超限处置参数。"""
    if on_limit not in ("reject", "skip"):
        raise ValueError(f"on_limit 仅支持 'reject' 或 'skip': {on_limit!r}")


def _fmt_bytes(size: int) -> str:
    if size >= 1024 * 1024 * 1024:
        return f"{size / (1024 * 1024 * 1024):.1f}GB"
    if size >= 1024 * 1024:
        return f"{size / (1024 * 1024):.1f}MB"
    return f"{size / 1024:.1f}KB"


# relationship 文件的真实读取上限：正常 rels 仅数 KB，超过即视为恶意构造。
_RELS_MAX_BYTES = 16 * 1024 * 1024

# Mammoth 会转换为 HTML 的正文相关 part 及其关系文件按包关系动态解析
# （见 _mammoth_content_parts）：正文 part 可经 _rels/.rels 重定向到诱饵
# document.xml 之外的任意条目，图片关系也可以出现在脚注/尾注/批注的
# rels 中（mammoth 同样触发图片回调）；只统计硬编码主文档关系会被改写
# Target 的文档绕过数量配额（见 BUG-026/BUG-042）。

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _rels_path_for(part_name: str) -> str:
    """part 的关系文件名，与 mammoth 同口径：同目录 _rels/<basename>.rels。"""
    dirname, _, basename = part_name.rpartition("/")
    return f"{dirname}/_rels/{basename}.rels" if dirname else f"_rels/{basename}.rels"


def _mammoth_join_path(base: str, target: str) -> str:
    """两段路径拼接，与 mammoth zips.join_path 同口径：以 / 开头的 target 重置。"""
    if target.startswith("/"):
        return target
    return f"{base}/{target}" if base else target


def _rels_internal_targets(
    zip_ref: zipfile.ZipFile, rels_name: str, type_suffix: str
) -> List[str]:
    """读取关系文件，返回 Type 以 /<type_suffix> 结尾的内部关系 Target 原文。

    关系文件缺失/损坏/不可读时返回空列表（对应 mammoth 的
    _try_read_entry_or_default 回退语义）；External 关系排除；
    TimeoutError 显式重抛，不得被降级吞掉。
    """
    try:
        rels_content = read_zip_entry_bounded(zip_ref, rels_name, _RELS_MAX_BYTES)
        rels_root = _safe_xml_fromstring(rels_content)
    except TimeoutError:
        raise
    except (KeyError, OSError, zipfile.BadZipFile, ValueError, SyntaxError):
        return []
    targets = []
    for rel in rels_root.findall(f".//{{{_REL_NS}}}Relationship"):
        if (rel.get("Type") or "").endswith(f"/{type_suffix}") and \
                (rel.get("TargetMode") or "Internal") != "External":
            target = rel.get("Target") or ""
            if target:
                targets.append(target)
    return targets


def _mammoth_content_parts(zip_ref: zipfile.ZipFile) -> List[Tuple[str, str]]:
    """按 mammoth 同口径动态解析会被读取的正文相关 part 及其关系文件。

    mammoth（docx/_find_part_paths）从 _rels/.rels 的 officeDocument 关系
    定位正文 part（不硬编码 word/document.xml），footnotes/endnotes/
    comments 又按正文 part 的关系文件定位；每个被读取 part 的关系文件
    都可能承载图片关系。只扫硬编码清单会被重定向 Target 的文档绕过
    图片数量配额（见 BUG-042）。

    解析口径与 mammoth 一致：关系目标按 join_path(base, target) 解析、
    lstrip("/") 后逐个检查存在性、取第一个命中的；无命中时回退默认
    路径（正文为 word/document.xml，notes 为 word/<名>.xml）。
    """
    entry_names = {info.filename for info in zip_ref.infolist()}
    main = None
    for target in _rels_internal_targets(zip_ref, "_rels/.rels", "officeDocument"):
        candidate = _mammoth_join_path("", target).lstrip("/")
        if candidate in entry_names:
            main = candidate
            break
    if main is None:
        main = "word/document.xml"

    parts = [(main, _rels_path_for(main))]
    base_dir = posixpath.dirname(main)
    for name in ("footnotes", "endnotes", "comments"):
        part = None
        for target in _rels_internal_targets(zip_ref, _rels_path_for(main), name):
            candidate = _mammoth_join_path(base_dir, target).lstrip("/")
            if candidate in entry_names:
                part = candidate
                break
        if part is None:
            part = f"word/{name}.xml"
        parts.append((part, _rels_path_for(part)))
    return parts


def _document_image_part_names(zip_ref: zipfile.ZipFile) -> set:
    """收集包内全部“物理图片资源”条目名。

    覆盖两类位置，缺一不可（见 BUG-026/BUG-042）：
      1. word/media/ 下的条目（无论是否被引用）——历史统计语义；
      2. Mammoth 实际读取的各正文 part（动态解析，见 _mammoth_content_parts）
         关系文件中 Type 以 /image 结尾的关系目标。DOCX 图片关系可指向包内
         任意位置（word/custom/、包根 custom/ 等），仅按 word/media/ 前缀
         统计会被改写 Target 的文档绕过资源防线。
    关系目标同时按两种口径解析（见 BUG-038）：Mammoth 读图用
    uri_to_zip_entry_name 的字面拼接（/ 开头取 uri[1:]，否则
    "word/" + uri，不做 normpath），含 . / .. 段的字面条目
    （word/custom/../imageN.png）只有字面口径能命中；规范化口径
    覆盖按规范名存储的正常文档。两种名字只要真实存在于 ZIP 即计入。
    非零内容的目录名条目（名字以 / 结尾）仍计入扫描：字面 Target
    指向的条目能被 zipfile 按字面名读取（见 BUG-043），安全校验随后
    无条件拒绝此类条目。零字节目录占位在前缀及关系两种口径均排除
    （见 BUG-049），不占图片配额。
    仅返回 ZIP 中真实存在的条目；External 链接自动排除。
    """
    entry_names = set()
    image_parts = set()
    for info in zip_ref.infolist():
        # 仅排除零字节目录占位；带内容的目录名条目仍扫描并由 ZIP
        # 安全校验拒绝，不能恢复 BUG-043 的 is_dir() 无条件排除。
        if info.is_dir() and info.file_size == 0:
            continue
        entry_names.add(info.filename)
        if info.filename.startswith("word/media/") and \
                len(info.filename) > len("word/media/"):
            image_parts.add(info.filename)

    for source_part, rels_name in _mammoth_content_parts(zip_ref):
        for target in _rels_internal_targets(zip_ref, rels_name, "image"):
            # 字面口径与 Mammoth 一致：正文 part 均在 word/ 下，
            # base 固定为 "word"，且不做 strip/反斜杠替换/normpath。
            literal = target[1:] if target.startswith("/") else f"word/{target}"
            for name in (literal, resolve_part_path(target, source_part)):
                # entry_names 已排除零字节目录，关系目标不能将其重新加入。
                if name in entry_names:
                    image_parts.add(name)
    return image_parts


def validate_docx_zip_security(zip_ref: zipfile.ZipFile, on_limit: str = "reject") -> None:
    """解压前依据 ZIP 中央目录元数据做资源耗尽防线校验。

    只读 compress_size/file_size 等声明值，不解压条目内容；超限抛
    DocxSecurityError。实际读取条目时另由 read_zip_entry_bounded /
    _read_media_image 在真实解压路径上兜底，防元数据谎报。

    防线分两层，on_limit 只影响第二层：
      1. ZIP bomb 等恶意特征（总解压量/单 entry 解压量/压缩比）无条件抛
         DocxSecurityError，任何模式下都不降级；
      2. 可降级资源（图片数量/单图大小/嵌入 Excel 大小的声明值检查）抛
         ResourceLimitExceeded；on_limit="skip" 时不在此抛出，改由提取
         阶段按真实读取逐项跳过（见 extract_content_from_docx）。
    """
    validate_on_limit(on_limit)
    limits = DOCX_SECURITY_LIMITS
    skip_mode = on_limit == "skip"
    total_compressed = 0
    total_uncompressed = 0
    image_count = 0
    seen_names = set()
    # 图片资源按“实际关系目标”识别：word/media 前缀之外的图片同样计数
    image_part_names = _document_image_part_names(zip_ref)

    for info in zip_ref.infolist():
        if info.filename in seen_names:
            raise DocxSecurityError(f"ZIP 包含重复条目名，解压语义不唯一: {info.filename}")
        seen_names.add(info.filename)
        if info.is_dir():
            # 名字以 / 结尾的“目录”条目仍能被 zipfile 按字面名读取：
            # 声明非零内容的（正常目录条目 file_size 恒为 0）直接拒绝，
            # 不得因 is_dir 在体积/压缩比/总量三道检查之前跳过（见 BUG-041）。
            if info.file_size > 0:
                raise DocxSecurityError(
                    f"ZIP 目录名条目声明非零内容（按字面名可读取）: "
                    f"{info.filename}（{_fmt_bytes(info.file_size)}）"
                )
            total_compressed += info.compress_size
            continue
        name = info.filename
        compressed = info.compress_size
        uncompressed = info.file_size
        total_compressed += compressed
        total_uncompressed += uncompressed

        if uncompressed > limits["entry_uncompressed"]:
            raise DocxSecurityError(
                f"ZIP 条目解压后超过单文件上限 {_fmt_bytes(limits['entry_uncompressed'])}: "
                f"{name}（{_fmt_bytes(uncompressed)}）"
            )
        if compressed > 0 and uncompressed > compressed * limits["entry_ratio"]:
            raise DocxSecurityError(
                f"ZIP 条目压缩比超过 {limits['entry_ratio']}x: {name}"
            )

        if name in image_part_names:
            image_count += 1
            if not skip_mode and uncompressed > limits["image_file_size"]:
                raise ResourceLimitExceeded(
                    f"图片超过单图大小上限 {_fmt_bytes(limits['image_file_size'])}: {name}"
                )
        if name.startswith("word/embeddings/") and name.lower().endswith(".xlsx"):
            if not skip_mode and uncompressed > limits["embedded_excel_size"]:
                raise ResourceLimitExceeded(
                    f"嵌入 Excel 超过大小上限 {_fmt_bytes(limits['embedded_excel_size'])}: {name}"
                )

    if not skip_mode and image_count > limits["image_count"]:
        raise ResourceLimitExceeded(f"图片数量超过上限 {limits['image_count']}: 实际 {image_count} 张")
    if total_uncompressed > limits["total_uncompressed"]:
        raise DocxSecurityError(
            f"ZIP 总解压大小超过上限 {_fmt_bytes(limits['total_uncompressed'])}: "
            f"实际 {_fmt_bytes(total_uncompressed)}"
        )
    if total_compressed > limits["total_ratio_min_compressed"]:
        if total_uncompressed > total_compressed * limits["total_ratio"]:
            raise DocxSecurityError(
                f"ZIP 总压缩比超过 {limits['total_ratio']}x: "
                f"{_fmt_bytes(total_uncompressed)}/{_fmt_bytes(total_compressed)}"
            )


def read_zip_entry_bounded(
    zip_ref: zipfile.ZipFile, name: str, max_bytes: int,
    error_cls: type = DocxSecurityError,
) -> bytes:
    """带实际上限的条目读取：边解压边计数，超过 max_bytes 立即中止。

    validate_docx_zip_security 依赖 ZIP 声明值，本函数在真实解压路径上
    再兜一道底，防止中央目录元数据与实际数据不一致的恶意构造。
    error_cls 允许可降级资源（图片/嵌入 Excel）抛 ResourceLimitExceeded，
    供 on_limit="skip" 精确捕获降级；默认 DocxSecurityError 保持既有语义。
    """
    chunks = []
    total = 0
    with zip_ref.open(name) as fh:
        while True:
            chunk = fh.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > max_bytes:
                raise error_cls(
                    f"ZIP 条目实际解压超过上限 {_fmt_bytes(max_bytes)}: {name}"
                )
            chunks.append(chunk)
    return b"".join(chunks)


def _png_dimensions(data: bytes) -> Optional[Tuple[int, int]]:
    if len(data) < 24 or data[12:16] != b"IHDR":
        return None
    return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")


_JPEG_SOF_MARKERS = frozenset(range(0xC0, 0xD0)) - {0xC4, 0xC8, 0xCC}


def _jpeg_dimensions(data: bytes) -> Optional[Tuple[int, int]]:
    i = 2
    length = len(data)
    while i + 4 <= length:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker == 0xFF:
            i += 1
            continue
        if marker in (0xD8, 0xD9, 0xDA):  # SOI/EOI/SOS：SOF 应在 SOS 之前出现
            return None
        if 0xD0 <= marker <= 0xD7 or marker == 0x01:
            i += 2
            continue
        seg_len = int.from_bytes(data[i + 2:i + 4], "big")
        if seg_len < 2:
            return None
        if marker in _JPEG_SOF_MARKERS and i + 9 <= length:
            height = int.from_bytes(data[i + 5:i + 7], "big")
            width = int.from_bytes(data[i + 7:i + 9], "big")
            return width, height
        i += 2 + seg_len
    return None


def _gif_dimensions(data: bytes) -> Optional[Tuple[int, int]]:
    if len(data) < 10:
        return None
    return int.from_bytes(data[6:8], "little"), int.from_bytes(data[8:10], "little")


def _bmp_dimensions(data: bytes) -> Optional[Tuple[int, int]]:
    if len(data) < 26:
        return None
    width = int.from_bytes(data[18:22], "little", signed=True)
    height = abs(int.from_bytes(data[22:26], "little", signed=True))
    return width, height


def _webp_dimensions(data: bytes) -> Optional[Tuple[int, int]]:
    if len(data) < 30:
        return None
    chunk = data[12:16]
    if chunk == b"VP8X":
        # 扩展格式：画布宽高各 24bit，存储的是实际值-1
        return (int.from_bytes(data[24:27], "little") + 1,
                int.from_bytes(data[27:30], "little") + 1)
    if chunk == b"VP8 ":
        # 有损格式：keyframe 宽高各 14bit
        return (int.from_bytes(data[26:28], "little") & 0x3FFF,
                int.from_bytes(data[28:30], "little") & 0x3FFF)
    if chunk == b"VP8L" and len(data) >= 25:
        # 无损格式：签名字节后宽高各 14bit，存储的是实际值-1
        bits = int.from_bytes(data[21:25], "little")
        return (bits & 0x3FFF) + 1, ((bits >> 14) & 0x3FFF) + 1
    return None


def _tiff_dimensions(data: bytes) -> Optional[Tuple[int, int]]:
    if len(data) < 8:
        return None
    if data[:2] == b"II":
        endian = "little"
    elif data[:2] == b"MM":
        endian = "big"
    else:
        return None
    ifd_offset = int.from_bytes(data[4:8], endian)
    if ifd_offset + 2 > len(data):
        return None
    entry_count = int.from_bytes(data[ifd_offset:ifd_offset + 2], endian)
    width = height = None
    for idx in range(entry_count):
        base = ifd_offset + 2 + idx * 12
        if base + 12 > len(data):
            break
        tag = int.from_bytes(data[base:base + 2], endian)
        if tag not in (256, 257):  # ImageWidth / ImageLength
            continue
        vtype = int.from_bytes(data[base + 2:base + 4], endian)
        raw = data[base + 8:base + 12]
        if vtype == 3:  # SHORT
            value = int.from_bytes(raw[:2], endian)
        elif vtype == 4:  # LONG
            value = int.from_bytes(raw[:4], endian)
        elif vtype in (16, 17):  # LONG8 / SLONG8
            value = int.from_bytes(raw[:8], endian)
        else:
            continue
        if tag == 256:
            width = value
        else:
            height = value
    if width and height:
        return width, height
    return None


def image_pixel_count(image_data: bytes) -> Optional[int]:
    """从图片头部解析宽高并返回像素数（宽×高）；无法解析返回 None。

    用于解压炸弹（decompression bomb）检测：仅需头部几十字节即可判定，
    不解码像素数据。WMF/EMF 等矢量格式无固定位图像素，返回 None。
    """
    data = image_data or b""
    dims: Optional[Tuple[int, int]] = None
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        dims = _png_dimensions(data)
    elif data[:2] == b'\xff\xd8':
        dims = _jpeg_dimensions(data)
    elif data[:6] in (b'GIF87a', b'GIF89a'):
        dims = _gif_dimensions(data)
    elif data[:2] == b'BM':
        dims = _bmp_dimensions(data)
    elif data[:4] == b'RIFF' and data[8:12] == b'WEBP':
        dims = _webp_dimensions(data)
    elif data[:4] in (b'II*\x00', b'MM\x00*'):
        dims = _tiff_dimensions(data)
    if not dims:
        return None
    width, height = dims
    if width <= 0 or height <= 0:
        return None
    return width * height


def _read_media_image(zip_ref: zipfile.ZipFile, name: str) -> bytes:
    """读取包内媒体图片并执行大小/像素防线（真实解压路径的兜底）。

    大小与像素均属可降级资源限制，抛 ResourceLimitExceeded：reject 模式
    下与 DocxSecurityError 处置一致（整篇拒绝），skip 模式下由调用方
    捕获并跳过该图片。
    """
    limits = DOCX_SECURITY_LIMITS
    image_data = read_zip_entry_bounded(
        zip_ref, name, limits["image_file_size"], error_cls=ResourceLimitExceeded
    )
    pixels = image_pixel_count(image_data)
    if pixels is not None and pixels > limits["image_pixels"]:
        raise ResourceLimitExceeded(
            f"图片像素超过上限 {limits['image_pixels']}: {name}（{pixels} 像素）"
        )
    return image_data


def sha256_file(path: str, chunk_size: int = 1024 * 1024) -> str:
    """流式计算文件 SHA-256，避免大文件一次性读入内存。"""
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _path_status(path: str) -> Optional[os.stat_result]:
    """lstat 查询路径状态；TimeoutError 显式重抛，其余 OSError 按不存在处理。

    os.path.lexists/islink/isfile 内部会捕获 OSError（含其子类
    TimeoutError），路径检查会吞掉批处理超时信号（见 BUG-035）——必须
    直接 lstat 并在降级前显式重抛。
    """
    try:
        return os.lstat(path)
    except TimeoutError:
        raise
    except OSError:
        return None


def _open_exclusive_temp(directory: str, mode: int = 0o666) -> Tuple[int, str]:
    """在 directory 内独占创建随机短名临时文件，返回 (fd, 路径)。

    临时名固定为 ``.`` + 12 位十六进制随机字符 + ``.tmp``，长度与目标
    文件名无关：目标名接近文件系统单名上限（ext4 按 255 字节计，80 个
    汉字的文档名即达到）时，按目标名拼前缀会让临时名越限触发
    ENAMETOOLONG（见 BUG-036）。以 mode 请求创建权限、由内核按当前
    umask 归一；新目标默认 0o666，覆盖普通文件时传入原权限，确保临时
    文件从创建起就不扩大访问权限（见 BUG-037）。无需查询进程 umask：查询需要
    ``os.umask(0)``/恢复两步，超时信号落在两步之间会让进程 umask 永久
    变成 0（见 BUG-034）。
    """
    for _ in range(100):
        path = os.path.join(directory, f".{os.urandom(6).hex()}.tmp")
        try:
            fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, mode)
        except FileExistsError:
            continue
        return fd, path
    raise OSError(f"无法在目录内创建可用的临时文件: {directory}")


def _atomic_write_text(path: str, text: str) -> None:
    """独占创建同目录随机临时文件写入后原子替换到 path。

    临时文件随机且独占创建，os.replace 只替换 path 的目录项本身：目标
    位置即使预置了指向外部的符号链接也不会被跟随，写入不会越出输出
    目录（见 BUG-023）。新文件的权限由独占创建时的 umask 自然决定；
    目标已是普通文件时，先按原权限创建临时文件，写入期间不扩大权限，
    替换前再对齐原权限位（见 BUG-028/037）；全程不查询/修改进程级
    umask（见 BUG-034）。
    """
    st = _path_status(path)
    target_mode = stat.S_IMODE(st.st_mode) if st is not None and stat.S_ISREG(st.st_mode) else None
    fd, tmp_path = _open_exclusive_temp(
        os.path.dirname(path) or ".", mode=target_mode if target_mode is not None else 0o666)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)
        if target_mode is not None:
            os.chmod(tmp_path, target_mode)
        os.replace(tmp_path, path)
    except BaseException:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


def _same_file_content(path: str, data: bytes) -> bool:
    """按“先比大小再读内容”判断 path 是否与 data 一致；读失败按不一致处理。

    TimeoutError 是 OSError 子类且承载批处理超时信号，必须显式重抛，
    不得按“内容不一致”降级吞掉（见 BUG-025 验收补充）。
    """
    try:
        if os.path.getsize(path) != len(data):
            return False
        with open(path, "rb") as f:
            return f.read() == data
    except TimeoutError:
        raise
    except OSError:
        return False


def _safe_realpath(path: str) -> str:
    """os.path.realpath 的超时安全替代：路径组件上的 TimeoutError 继续上抛。

    realpath 非 strict 模式内部 except OSError 会把组件 lstat 抛出的
    TimeoutError（批处理 SIGALRM 的单文档超时信号）当“组件不存在”吞掉
    并返回字面路径，使 alarm 静默丢失（见 BUG-039）。这里按 realpath
    （非 strict、跟随符号链接、不展开 ~）的语义逐段解析：lstat 走
    _path_status 的超时安全封装，readlink 显式重抛 TimeoutError；组件
    不存在时按普通组件继续解析，后续 .. 仍回退并解析后续符号链接；
    readlink 失败按普通组件处理。用待解析栈及链接缓存识别成环链接，
    成环处保留字面组件并继续解析余下路径；长链不使用 Python 递归，
    不会因固定层数上限提前返回。Windows 无 SIGALRM、无吞超时问题，
    直接用 os.path.realpath。
    """
    if os.name == "nt":
        return os.path.realpath(path)
    abs_path = path if os.path.isabs(path) else os.path.join(os.getcwd(), path)
    # None 标记表示某个链接的目标已经完整解析，下一项是缓存键。
    parts = abs_path.split("/")[::-1]
    seen = {}
    resolved = "/"
    while parts:
        part = parts.pop()
        if part is None:
            seen[parts.pop()] = resolved
            continue
        if not part or part == ".":
            continue
        if part == "..":
            resolved = posixpath.dirname(resolved) or "/"
            continue
        current = posixpath.join(resolved, part)
        st = _path_status(current)
        if st is None or not stat.S_ISLNK(st.st_mode):
            # 非 strict 语义：缺失组件仍保留，但后续 .. / 链接继续解析。
            resolved = current
            continue
        if current in seen:
            # None 是仍在解析的链接，说明成环；完整解析的链接则复用缓存。
            resolved = seen[current] or current
            continue
        try:
            target = os.readlink(current)
        except TimeoutError:
            raise
        except OSError:
            resolved = current  # 链接不可读：按普通组件继续
            continue
        seen[current] = None
        parts.extend((current, None))
        parts.extend(target.split("/")[::-1])
        if target.startswith("/"):
            resolved = "/"
    return resolved


def _allocate_asset_path(
    assets_dir: str, base_name: str, actual_ext: str,
    digest: str, image_data: bytes,
) -> Tuple[str, str]:
    """为图片分配输出文件名与路径：自然名 → 短 hash 名 → hash_序号名。

    只有“普通文件且字节一致”才复用既有条目；符号链接（含悬空）、目录
    或内容不同的占用一律换下一候选（见 BUG-027：hash 候选被占用时不得
    盲复用）。返回 (文件名, 路径)；路径可能已存在（复用）或不存在（待写）。
    """
    candidates = [f"{base_name}{actual_ext}", f"{base_name}_{digest[:8]}{actual_ext}"]
    candidates += [
        f"{base_name}_{digest[:8]}_{seq}{actual_ext}" for seq in range(2, 100)
    ]
    for name in candidates:
        path = os.path.join(assets_dir, name)
        st = _path_status(path)
        if st is None:
            return name, path
        # lstat 判定的普通文件才可能是复用对象：符号链接（含悬空）、
        # 目录或内容不同的占用一律换下一候选（见 BUG-027）。
        if stat.S_ISREG(st.st_mode) and _same_file_content(path, image_data):
            return name, path
    raise OSError(f"无法为图片分配可用的输出文件名: {base_name}{actual_ext}")


def _write_asset_file_exclusive(path: str, data: bytes) -> bool:
    """以独占创建写入图片文件；返回是否实际写入。

    O_CREAT|O_EXCL 保证检查与写入之间不会被同名符号链接/文件置换
    （见 BUG-027）。条目已存在时不再静默跳过（见 BUG-047）：普通文件
    且字节一致视为安全复用（返回 False，无需写入）；符号链接/目录/
    内容不同的占用抛 FileExistsError，由调用方换下一候选重试，保证
    Markdown 引用的文件始终是内容正确的普通文件。
    """
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    except FileExistsError:
        st = _path_status(path)
        if st is not None and stat.S_ISREG(st.st_mode) and _same_file_content(path, data):
            return False
        raise
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return True


def write_conversion_sentinel(
    final_output_dir: str, folder_name: str, source_sha256: str,
    on_limit: str = "reject",
) -> None:
    """原子写入转换完成标记（随机独占临时文件 + rename），记录输出目录名与源文件哈希。

    批处理据此判断“输出完整且与当前源一致”；仅转换全部成功后调用。
    标记写失败不影响本次转换结果，仅意味着批处理下次会重转。
    临时文件必须用随机名独占创建：固定名（.converted.tmp）可被预置为
    指向外部的符号链接，open 会跟随其覆盖外部文件（见 BUG-024）。
    写入与权限策略复用 _atomic_write_text（见 BUG-028/034/036）。
    """
    sentinel_path = os.path.join(final_output_dir, SENTINEL_FILENAME)
    validate_on_limit(on_limit)
    payload = json.dumps(
        {"folder_name": folder_name, "source_sha256": source_sha256, "on_limit": on_limit},
        ensure_ascii=False,
        sort_keys=True,
    )
    try:
        _atomic_write_text(sentinel_path, payload)
    except TimeoutError:
        raise
    except OSError:
        logger.warning("写入完成标记失败: %s", sentinel_path, exc_info=True)


def read_conversion_sentinel(directory: str) -> Optional[Dict[str, str]]:
    """读取完成标记；缺失/损坏/旧格式（纯文本等非 JSON 对象）返回 None。"""
    sentinel_path = os.path.join(directory, SENTINEL_FILENAME)
    try:
        with open(sentinel_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict):
        return None
    folder_name = data.get("folder_name")
    source_sha256 = data.get("source_sha256")
    # V0.1.6 sentinel 没有策略字段，当时唯一行为是 reject。
    on_limit = data.get("on_limit", "reject")
    if not isinstance(folder_name, str) or not folder_name:
        return None
    if not isinstance(source_sha256, str) or not source_sha256:
        return None
    if on_limit not in ("reject", "skip"):
        return None
    return {"folder_name": folder_name, "source_sha256": source_sha256, "on_limit": on_limit}


def prune_stale_assets(assets_dir: str, current_image_sources) -> None:
    """成功转换后删除当前结果不再使用的旧资源文件。

    单文件 API/CLI 允许复用同名输出目录。只在 Markdown 已成功生成后执行，
    避免失败转换提前破坏旧结果；目录及其子目录不删除。
    """
    current_names = {
        os.path.basename(source)
        for source in current_image_sources
        if isinstance(source, str) and source
    }
    with os.scandir(assets_dir) as entries:
        for entry in entries:
            if entry.name in current_names:
                continue
            if entry.is_symlink() or entry.is_file(follow_symlinks=False):
                os.remove(entry.path)


def _mammoth_embedded_media_key(image, known_parts: set) -> Optional[str]:
    """尽力从 Mammoth 图片打开函数中取出嵌入媒体路径。

    Mammoth 的公开 Image 对象不暴露 relationship/path，但当前稳定实现
    （docx/body_xml.py 的 open_image 闭包）会捕获按 relationship 目标
    解析出的 zip 条目名（如 word/custom/image1.png）。闭包值就是 Mammoth
    实际读取的字面条目名，可含 . / .. 段（见 BUG-038），而配额白名单
    allowed_media_paths 也按字面条目名存放——先按字面匹配，未命中再退
    回归一化/关系解析两种口径。只接受能在 known_parts（包内物理图片
    条目集合）中命中的字符串，避免把闭包里的 content_type 等无关字符串
    误判为媒体路径；链接图片不归入 ZIP 配额。
    """
    opener = getattr(image, "open", None)
    closure = getattr(opener, "__closure__", None) or ()
    for cell in closure:
        try:
            value = cell.cell_contents
        except ValueError:
            continue
        if isinstance(value, str):
            if value in known_parts:
                return value
            normalized = posixpath.normpath(value.replace("\\", "/").lstrip("/"))
            if normalized in known_parts:
                return normalized
            resolved = resolve_part_path(value)
            if resolved in known_parts:
                return resolved
    return None


def _escape_plain_cell_text(text: str) -> str:
    """原始纯文本单元格片段的最终序列化转义。

    反斜线与管道必须整体转义且顺序固定（先反斜线后管道）：只检查管道
    前一个字符会把原文里的连续反斜线误判为已有转义，Python Markdown
    渲染时管道仍会拆列（如原文 A\\\\|B，见 BUG-020 验收补充）。字面
    标签样文本（如 <table>、<!-- -->、<?php ?>、<![CDATA[]]>、<采暖>）
    转成实体，避免下游渲染器把它当作原始 HTML 处理：标签名首字符覆盖
    非 ASCII（槽位语法的中文标签，见 BUG-048），仅排除空白首字符
    （"a < b" 类普通比较文本不是标签样，保留裸形式维持可读性——
    转义本身渲染无损，但会降低 md 源文件可读性）。
    """
    text = (text or "").replace("\\", "\\\\").replace("|", r"\|")
    return re.sub(
        r"<[^\s<>][^<>]*>",
        lambda m: m.group(0).replace("<", "&lt;", 1).replace(">", "&gt;"),
        text,
    )


def _escape_unescaped_pipes(text: str) -> str:
    """对已处于最终转义层次的文本（嵌套表格产物）仅转义未受保护的管道。

    管道前的连续反斜线为偶数条（含 0）说明这些反斜线都是转义产物、
    管道本身未转义，需要补一条转义；奇数条说明管道已被转义，保持原样。
    """
    def _repl(match):
        run = match.group(1)
        if len(run) % 2:
            return match.group(0)
        return run + r"\|"

    return re.sub(r"(\\*)\|", _repl, text)


def _normalize_markdown_cell_text(value: str) -> str:
    """将原始纯文本单元格内容（Excel 路径）规范化为管道表单元格单行文本。

    在统一最终序列化层执行：反斜线→\\\\、管道→\\|（顺序固定），字面
    ASCII 标签转实体。不做 HTML unescape——Excel 单元格是纯文本。
    """
    text = (value or "").replace("\xa0", " ")
    text = _escape_plain_cell_text(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.split("\n")]
    lines = [line for line in lines if line]
    return "<br>".join(lines) if lines else ""


def _serialize_table_cell(parts) -> str:
    """将 HTML 表格单元格收集到的片段序列化为管道表单元格的单行最终文本。

    纯文本片段（str）处于原始形态，整体应用 _escape_plain_cell_text；
    嵌套表格产物（("nested", md) 元组）的文本已由内层序列化到最终转义
    层次，仅对未受保护的管道补转义（内层结构管道），反斜线序列不再改动。
    两类片段的转义规则互补，任何管道在进入最终 Markdown 时都恰好携带
    一条有效转义，不产生双重转义或裸管道（见 BUG-020）。
    """
    out = []
    for part in parts:
        if isinstance(part, tuple):
            out.append(_escape_unescaped_pipes(part[1]))
        else:
            out.append(_escape_plain_cell_text(part))
    text = "".join(out)
    text = text.replace("\xa0", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.split("\n")]
    lines = [line for line in lines if line]
    return "<br>".join(lines) if lines else ""


class _TableHTMLParser(HTMLParser):
    """解析单个 HTML table，保留 rowspan/colspan 与单元格文本。

    单元格内的嵌套表格（<td> 中再含 <table>）先递归转换为 Markdown，
    以 ("nested", md) 元组追加到单元格片段；纯文本片段为 str。两类片段
    由 _serialize_table_cell 在统一的最终转义层次分别序列化。
    """

    def __init__(self):
        super().__init__()
        self.rows = []
        self._in_tr = False
        self._in_cell = False
        self._current_row = []
        self._cell_parts = []  # str（原始文本/换行标记）或 ("nested", md)
        self._cell_tag = None
        self._cell_rowspan = 1
        self._cell_colspan = 1
        # 嵌套表格原始 HTML 收集状态（0 表示未在收集）
        self._nested_depth = 0
        self._nested_parts = []

    @staticmethod
    def _safe_int(raw, default=1):
        try:
            value = int(raw)
            return value if value > 0 else default
        except TimeoutError:
            raise  # 超时信号不得被降级捕获吞掉（见 BUG-025）
        except Exception:
            return default

    def _append_cell_newline(self):
        """在单元格内补换行标记；末尾片段是嵌套产物时无条件补。"""
        if self._cell_parts and isinstance(self._cell_parts[-1], str) \
                and self._cell_parts[-1].endswith("\n"):
            return
        self._cell_parts.append("\n")

    def handle_starttag(self, tag, attrs):
        attrs_map = dict(attrs)
        tag = tag.lower()

        if self._nested_depth:
            # 收集嵌套表格的原始 HTML，交给递归转换处理。
            if tag == "table":
                self._nested_depth += 1
            self._nested_parts.append(self.get_starttag_text() or f"<{tag}>")
            return

        if tag == "table" and self._in_cell:
            self._nested_depth = 1
            self._nested_parts = [self.get_starttag_text() or "<table>"]
            return

        if tag == "tr":
            self._in_tr = True
            self._current_row = []
            return

        if tag in ("td", "th") and self._in_tr:
            self._in_cell = True
            self._cell_tag = tag
            self._cell_parts = []
            self._cell_rowspan = self._safe_int(attrs_map.get("rowspan"), 1)
            self._cell_colspan = self._safe_int(attrs_map.get("colspan"), 1)
            return

        if self._in_cell and tag in ("br",):
            self._cell_parts.append("\n")
        elif self._in_cell and tag in ("p", "div", "li"):
            self._append_cell_newline()

    def handle_startendtag(self, tag, attrs):
        if self._nested_depth:
            self._nested_parts.append(self.get_starttag_text() or f"<{tag}/>")
            return
        super().handle_startendtag(tag, attrs)

    def handle_endtag(self, tag):
        tag = tag.lower()

        if self._nested_depth:
            self._nested_parts.append(f"</{tag}>")
            if tag == "table":
                self._nested_depth -= 1
                if self._nested_depth == 0:
                    nested_html = "".join(self._nested_parts)
                    self._nested_parts = []
                    nested_md = table_html_to_markdown(nested_html)
                    if nested_md:
                        self._cell_parts.append(("nested", nested_md))
            return

        if tag in ("p", "div", "li") and self._in_cell:
            self._append_cell_newline()
            return

        if tag in ("td", "th") and self._in_cell:
            text = _serialize_table_cell(self._cell_parts)
            self._current_row.append(
                {
                    "text": text,
                    "rowspan": self._cell_rowspan,
                    "colspan": self._cell_colspan,
                    "is_header": self._cell_tag == "th",
                }
            )
            self._in_cell = False
            self._cell_parts = []
            self._cell_tag = None
            self._cell_rowspan = 1
            self._cell_colspan = 1
            return

        if tag == "tr" and self._in_tr:
            if self._current_row:
                self.rows.append(self._current_row)
            self._in_tr = False
            self._current_row = []

    def handle_data(self, data):
        if self._nested_depth:
            # 嵌套收集按原始 HTML 语义重新序列化：HTMLParser 已把实体解码
            # 成文本，必须重新转义，否则字面标签文本（如正文写的 <table>）
            # 会在递归解析时被当作真实标签，导致整块内容丢失（见 BUG-020）。
            self._nested_parts.append(_escape_html(data, quote=False))
            return
        if self._in_cell:
            # 单元格文本保持实体形式，由 html_to_markdown 末尾的统一
            # unescape 还原；提前解码会让字面标签文本被后续的 HTML
            # 标签清理步骤删除。
            self._cell_parts.append(_escape_html(data, quote=False))


def _normalize_list_item_text(value: str) -> str:
    lines = []
    for raw in (value or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not raw.strip():
            continue
        # 嵌套列表行保留原始缩进，避免层级被破坏。
        if re.match(r"^\s+(-|\d+\.)\s+", raw):
            lines.append(raw.rstrip())
        else:
            lines.append(re.sub(r"\s+", " ", raw).strip())
    return "\n".join(lines)


class _ListHTMLTransformer(HTMLParser):
    """将 HTML 列表结构转换为 Markdown 列表，保留嵌套层级。"""

    def __init__(self):
        super().__init__()
        self._out = []
        self._list_stack = []  # [{"type": "ul"/"ol", "items": [str, ...]}]
        self._li_stack = []  # [list[str], ...]

    @staticmethod
    def _attrs_to_str(attrs):
        if not attrs:
            return ""
        pairs = []
        for k, v in attrs:
            if v is None:
                pairs.append(k)
            else:
                escaped = str(v).replace('"', "&quot;")
                pairs.append(f'{k}="{escaped}"')
        return " " + " ".join(pairs)

    @staticmethod
    def _render_list(context, depth):
        indent = "  " * depth
        is_ordered = context["type"] == "ol"
        lines = []
        for idx, item in enumerate(context["items"], 1):
            marker = f"{idx}. " if is_ordered else "- "
            normalized = _normalize_list_item_text(item)
            if not normalized:
                lines.append(f"{indent}{marker}".rstrip())
                continue
            item_lines = normalized.splitlines()
            lines.append(f"{indent}{marker}{item_lines[0].strip()}")
            for extra in item_lines[1:]:
                if extra.startswith("  "):
                    lines.append(f"{indent}{extra}")
                else:
                    lines.append(f"{indent}  {extra.strip()}")
        return "\n".join(lines)

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in ("ul", "ol"):
            self._list_stack.append({"type": tag, "items": []})
            return
        if tag == "li" and self._list_stack:
            self._li_stack.append([])
            return

        if self._li_stack:
            if tag == "br":
                self._li_stack[-1].append("\n")
            elif tag in ("p", "div"):
                if self._li_stack[-1] and not self._li_stack[-1][-1].endswith("\n"):
                    self._li_stack[-1].append("\n")
            return

        self._out.append(f"<{tag}{self._attrs_to_str(attrs)}>")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("ul", "ol") and self._list_stack:
            context = self._list_stack.pop()
            md = self._render_list(context, len(self._list_stack))
            if self._li_stack:
                if self._li_stack[-1] and not self._li_stack[-1][-1].endswith("\n"):
                    self._li_stack[-1].append("\n")
                self._li_stack[-1].append(md)
            else:
                # 顶层列表后补空行，避免后续表格被当作列表延续文本。
                self._out.append("\n" + md + "\n\n")
            return

        if tag == "li" and self._li_stack:
            item_text = "".join(self._li_stack.pop())
            if self._list_stack:
                self._list_stack[-1]["items"].append(item_text)
            return

        if self._li_stack:
            if tag in ("p", "div"):
                if self._li_stack[-1] and not self._li_stack[-1][-1].endswith("\n"):
                    self._li_stack[-1].append("\n")
            return

        self._out.append(f"</{tag}>")

    def handle_startendtag(self, tag, attrs):
        tag = tag.lower()
        if self._li_stack and tag == "br":
            self._li_stack[-1].append("\n")
            return
        if self._li_stack:
            return
        self._out.append(f"<{tag}{self._attrs_to_str(attrs)}/>")

    def handle_data(self, data):
        # HTMLParser 已把实体解码成文本，重新转义保持实体形式，由
        # html_to_markdown 末尾的统一 unescape 还原一次；否则字面标签
        # 文本（如正文写的 <table>）会被后续的 HTML 标签清理步骤删除
        # （见 BUG-020）。
        if self._li_stack:
            self._li_stack[-1].append(_escape_html(data, quote=False))
            return
        if self._list_stack:
            # 列表容器中但不在 li 内的噪声文本通常只有空白，忽略。
            return
        self._out.append(_escape_html(data, quote=False))

    def get_output(self):
        return "".join(self._out)


def transform_html_lists_to_markdown(html: str) -> str:
    parser = _ListHTMLTransformer()
    parser.feed(html)
    parser.close()
    return parser.get_output()


def _expand_table_rows(rows):
    """将包含 rowspan/colspan 的行展开为等宽二维表。"""
    expanded = []
    spans = {}  # col_idx -> {"rows_left": int, "text": str}

    for row in rows:
        out_row = []
        col = 0

        def consume_span_at_current_col():
            nonlocal col
            while col in spans:
                span = spans[col]
                out_row.append(span["text"])
                span["rows_left"] -= 1
                if span["rows_left"] <= 0:
                    spans.pop(col, None)
                col += 1

        consume_span_at_current_col()
        for cell in row:
            consume_span_at_current_col()
            text = cell["text"]
            rowspan = max(1, int(cell["rowspan"]))
            colspan = max(1, int(cell["colspan"]))
            for offset in range(colspan):
                out_row.append(text)
                if rowspan > 1:
                    spans[col + offset] = {"rows_left": rowspan - 1, "text": text}
            col += colspan

        consume_span_at_current_col()
        expanded.append(out_row)

    while spans:
        out_row = []
        col = 0
        max_col = max(spans.keys())
        while col <= max_col:
            if col in spans:
                span = spans[col]
                out_row.append(span["text"])
                span["rows_left"] -= 1
                if span["rows_left"] <= 0:
                    spans.pop(col, None)
            else:
                out_row.append("")
            col += 1
        expanded.append(out_row)

    width = max((len(row) for row in expanded), default=0)
    if width:
        expanded = [row + [""] * (width - len(row)) for row in expanded]
    return expanded


def table_html_to_markdown(table_html: str) -> str:
    parser = _TableHTMLParser()
    parser.feed(table_html)
    parser.close()

    rows = _expand_table_rows(parser.rows)
    if not rows:
        return ""

    lines = []
    lines.append("| " + " | ".join(rows[0]) + " |")
    lines.append("| " + " | ".join(["---"] * len(rows[0])) + " |")
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines) + "\n\n"


_TABLE_TAG_RE = re.compile(r"</?table\b[^>]*>", re.IGNORECASE)

# 匹配 <img src=...>，支持双引号、单引号、无引号三种写法；
# 等号两侧的空白也要容忍（src= "a.png" 是合法 HTML，见 BUG-021）
_IMG_TAG_RE = re.compile(
    r"<img\b[^>]*\bsrc\s*=\s*"
    r"(?:\"(?P<src1>[^\"]*)\"|'(?P<src2>[^']*)'|(?P<src3>[^\s\"'=<>`]+))"
    r"[^>]*/?>",
    flags=re.IGNORECASE,
)


def replace_html_tables(html: str) -> str:
    """按配对标签转换所有 HTML 表格，支持表格嵌套。

    非贪婪正则会在内层表格的 </table> 处提前截断，导致外层表格其余单元格
    丢失。这里用深度计数定位每个顶层 <table>...</table>（含嵌套），整体交给
    table_html_to_markdown 递归处理。
    """
    out = []
    pos = 0
    while True:
        match = _TABLE_TAG_RE.search(html, pos)
        if not match:
            out.append(html[pos:])
            break
        if match.group(0).startswith("</"):
            # 孤立的结束标签：原样保留，继续扫描
            out.append(html[pos:match.end()])
            pos = match.end()
            continue

        depth = 1
        scan = match.end()
        end = None
        while depth > 0:
            nxt = _TABLE_TAG_RE.search(html, scan)
            if not nxt:
                break
            if nxt.group(0).startswith("</"):
                depth -= 1
                if depth == 0:
                    end = nxt.end()
                    break
            else:
                depth += 1
            scan = nxt.end()

        if end is None:
            # 没有配对的 </table>（如文本里未配对的字面 <table> 标签）：
            # 该标签按字面放行并继续扫描，不得把该位置到文末的内容整段
            # 当作表格交给 table_html_to_markdown 吞掉（见 BUG-040）
            out.append(html[pos:match.end()])
            pos = match.end()
            continue

        out.append(html[pos:match.start()])
        out.append(table_html_to_markdown(html[match.start():end]))
        pos = end
    return "".join(out)


def promote_numbered_bold_headings(markdown: str) -> str:
    """将“编号 + 加粗标题”段落提升为 Markdown 标题。"""
    pattern = re.compile(
        r"^(?P<num>\d+(?:\.\d+)*)(?P<dot>\.)?\s+\*\*(?P<title>[^*\n]+)\*\*\s*$",
        flags=0,
    )
    heading_pattern = re.compile(r"^(#{1,6})\s+")

    lines = markdown.splitlines()
    out = []
    previous_heading_level = 0
    previous_promoted_depth = None

    for line in lines:
        heading_match = heading_pattern.match(line)
        if heading_match:
            previous_heading_level = len(heading_match.group(1))
            previous_promoted_depth = None
            out.append(line)
            continue

        match = pattern.match(line.strip())
        if not match:
            out.append(line)
            continue

        num = match.group("num")
        dot = match.group("dot") or ""
        title = match.group("title").strip()
        depth = num.count(".") + 1
        level = min(depth, 6)  # 1级编号 -> #，2级编号 -> ##

        # 在深层章节下的“1. **小节**”更接近子标题，避免被抬到过高层级。
        if depth == 1 and previous_heading_level >= 2:
            if previous_promoted_depth == 1:
                level = previous_heading_level
            else:
                level = min(previous_heading_level + 1, 6)

        promoted = f"{'#' * level} {num}{dot} {title}"
        out.append(promoted)
        previous_heading_level = level
        previous_promoted_depth = depth

    # 保持原始编号，不做自动重排，避免双语并列标题或手工编号被误改。
    return "\n".join(out)


def promote_leading_bold_title(markdown: str) -> str:
    """将文档开头“整行加粗标题”提升为一级标题（保守触发）。"""
    lines = markdown.splitlines()
    first_idx = None
    for i, line in enumerate(lines):
        if line.strip():
            first_idx = i
            break
    if first_idx is None:
        return markdown

    first_line = lines[first_idx].strip()
    m = re.match(r"^\*\*(?P<title>.+?)\*\*$", first_line)
    if not m:
        return markdown

    # 仅在后续存在“编号章节标题”时触发，降低把普通强调段误判成标题的风险。
    section_heading_re = re.compile(r"^#{1,6}\s+\d+(?:\.\d+)*\.?\s+")
    has_numbered_section_heading = any(section_heading_re.match(line.strip()) for line in lines[first_idx + 1 :])
    if not has_numbered_section_heading:
        return markdown

    title = m.group("title").strip()
    if not title:
        return markdown

    lines[first_idx] = f"# {title}"
    return "\n".join(lines)


def sanitize_stem(stem: str) -> str:
    raw = stem  # 保留原始值用于 hash
    normalized = unicodedata.normalize("NFKC", raw or "")
    # NFKC 归一化（如全角→半角）不加 hash：中文文档场景过于普遍，
    # 由此产生的罕见碰撞由 sentinel 的源哈希校验兜底（不一致即重转）。
    no_quotes = normalized
    for ch in _QUOTE_CHARS:
        no_quotes = no_quotes.replace(ch, "")
    substituted = _FORBIDDEN_FILENAME_CHARS_RE.sub("_", no_quotes)
    # 引号删除或非法字符替换是强丢失映射（如 a:b 与 a_b 同映射），
    # 需附加 hash 防止不同原始名称共享同一输出目录
    lossy = no_quotes != normalized or substituted != no_quotes
    stem = _WHITESPACE_RE.sub(" ", substituted).strip()
    stem = stem.strip(". ").strip()
    if not stem:
        return "document"
    if len(stem) <= 120 and not lossy:
        return stem
    # 清洗发生丢失或超长截断时，附加原始全名的短 hash，避免不同文件名映射到同一输出目录
    suffix = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:8]
    return f"{stem[:111]}_{suffix}"


def extract_heading_level_map(docx_path: str) -> Dict[str, int]:
    """解析 DOCX 的 heading bookmark 段落样式，映射为 Markdown 标题层级。"""
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tag_p = f"{{{ns_w}}}p"
    tag_ppr = f"{{{ns_w}}}pPr"
    tag_pstyle = f"{{{ns_w}}}pStyle"
    tag_bm = f"{{{ns_w}}}bookmarkStart"
    attr_name = f"{{{ns_w}}}name"
    attr_val = f"{{{ns_w}}}val"
    tag_t = f"{{{ns_w}}}t"

    def style_to_level(style_val: str) -> Optional[int]:
        if not style_val:
            return None
        raw = str(style_val).strip()
        m = re.search(r"(\d+)$", raw)
        if not m:
            return None
        n = int(m.group(1))
        if n <= 0:
            return None
        # style=1/2/3 对应一/二/三级标题。
        return min(n, 6)

    def infer_level_from_text(text: str) -> Optional[int]:
        m = re.match(r"^\s*(\d+(?:\.\d+)*)\s*\.?\s+", text or "")
        if not m:
            return None
        depth = m.group(1).count(".") + 1
        return min(depth, 6)

    level_map: Dict[str, int] = {}
    try:
        with zipfile.ZipFile(docx_path, "r") as zip_ref:
            doc_xml = _safe_xml_fromstring(zip_ref.read("word/document.xml"))
        for p in doc_xml.findall(f".//{tag_p}"):
            bm = p.find(f".//{tag_bm}")
            if bm is None:
                continue
            name = bm.get(attr_name)
            if not name or not name.startswith("heading_"):
                continue

            ppr = p.find(tag_ppr)
            style_val = ""
            if ppr is not None:
                pstyle = ppr.find(tag_pstyle)
                if pstyle is not None:
                    style_val = pstyle.get(attr_val, "")

            level = style_to_level(style_val)
            if level is None:
                text = "".join((t.text or "") for t in p.findall(f".//{tag_t}"))
                level = infer_level_from_text(text)
            if level is not None:
                level_map[name] = level
    except TimeoutError:
        raise  # 超时信号不得被降级捕获吞掉（见 BUG-025）
    except Exception:
        return {}

    return level_map


def resolve_part_path(target: str, source_part: str = "word/document.xml") -> str:
    """将 relationship target 解析为 docx zip 内的规范路径（normpath 归一）。

      - 以 / 开头的是包根绝对路径（如 /custom/image1.png -> custom/image1.png）；
      - 其余按 source_part 所在目录解析相对路径（如 word/document.xml 的
        media/image1.png -> word/media/image1.png）；
      - 兼容历史输入里已带 word/ 前缀的相对写法。

    注意 Mammoth 实际读图并不归一（uri_to_zip_entry_name 只做字面拼接，
    含 . / .. 段的条目两边口径不同，见 BUG-038）——配额统计等需要与
    Mammoth 对齐的场景必须同时计入字面名与这里的规范名。
    """
    target = (target or "").replace("\\", "/").strip()
    if not target:
        return ""
    if target.startswith("/"):
        return posixpath.normpath(target[1:])
    if target.startswith("word/"):
        return posixpath.normpath(target)
    base_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base_dir, target))


def parse_relationships(docx_path):
    """解析docx中的关系文件，找出Excel嵌入和对应预览图的映射。

    策略：
      1. 优先从 document.xml 中解析 <w:object> 节点，提取 OLEObject rId
         和 imagedata rId 的真实配对关系（最可靠）。
      2. 对方法1未覆盖的项，使用 "rId相邻" 启发式补全（兼容）。
    """
    excel_to_preview = {}  # Excel路径 -> 预览图路径
    preview_to_excel = {}  # 预览图路径 -> Excel路径
    ordered_pairs = []  # [(Excel路径, 预览图路径)]，按文档出现顺序

    # --- 公共：解析 rels 文件，建立 rId -> target 映射 ---
    NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
    relationships = {}  # rId -> {'type': ..., 'target': ...}

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        try:
            rels_content = zip_ref.read('word/_rels/document.xml.rels')
            rels_root = _safe_xml_fromstring(rels_content)
            for rel in rels_root.findall(f'.//{{{NS_REL}}}Relationship'):
                rid = rel.get('Id')
                rel_type = rel.get('Type', '').split('/')[-1]
                target = rel.get('Target', '')
                relationships[rid] = {'type': rel_type, 'target': target}
        except TimeoutError:
            raise  # 超时信号不得被降级捕获吞掉（见 BUG-025）
        except Exception as e:
            logger.warning("解析关系文件失败: %s", e)
            return excel_to_preview, preview_to_excel, ordered_pairs

        # --- 方法1：从 document.xml 解析 OLE 对象的真实引用 ---
        NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        NS_V = "urn:schemas-microsoft-com:vml"
        NS_O = "urn:schemas-microsoft-com:office:office"
        NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        try:
            doc_xml = zip_ref.read('word/document.xml')
            doc_root = _safe_xml_fromstring(doc_xml)

            # 查找所有 <w:object> 节点（可能嵌套在 mc:AlternateContent 等下面）
            for obj_node in doc_root.iter(f'{{{NS_W}}}object'):
                ole_rid = None
                img_rid = None

                # <o:OLEObject r:id="rIdX" />
                for ole in obj_node.iter(f'{{{NS_O}}}OLEObject'):
                    ole_rid = ole.get(f'{{{NS_R}}}id')

                # <v:imagedata r:id="rIdY" />
                for imgdata in obj_node.iter(f'{{{NS_V}}}imagedata'):
                    img_rid = imgdata.get(f'{{{NS_R}}}id')

                if ole_rid and img_rid and ole_rid in relationships and img_rid in relationships:
                    ole_target = resolve_part_path(relationships[ole_rid]['target'])
                    img_target = resolve_part_path(relationships[img_rid]['target'])
                    if ole_target.lower().endswith('.xlsx'):
                        excel_to_preview[ole_target] = img_target
                        preview_to_excel[img_target] = ole_target
                        ordered_pairs.append((ole_target, img_target))
        except TimeoutError:
            raise  # 超时信号不得被降级捕获吞掉（见 BUG-025）
        except Exception:
            pass  # document.xml 解析失败不影响后续

        # --- 方法2（补全）：rId 相邻启发式，补全方法1未覆盖的 Excel ---
        def rid_sort_key(rid: str) -> int:
            m = re.fullmatch(r"rId(\d+)", rid or "")
            return int(m.group(1)) if m else 10**9

        sorted_rids = sorted(relationships.keys(), key=rid_sort_key)

        for i, rid in enumerate(sorted_rids):
            rel = relationships[rid]
            if rel['type'] == 'package' and rel['target'].lower().endswith('.xlsx'):
                excel_file = resolve_part_path(rel['target'])
                if excel_file in excel_to_preview:
                    continue  # 已被方法1覆盖，跳过
                if i + 1 < len(sorted_rids):
                    next_rid = sorted_rids[i + 1]
                    next_rel = relationships[next_rid]
                    if next_rel['type'] == 'image':
                        preview_file = resolve_part_path(next_rel['target'])
                        excel_to_preview[excel_file] = preview_file
                        preview_to_excel[preview_file] = excel_file
                        ordered_pairs.append((excel_file, preview_file))

    return excel_to_preview, preview_to_excel, ordered_pairs


def _format_cell_value(cell) -> str:
    """将 openpyxl 单元格值转换为友好的字符串表示。"""
    if cell is None:
        return ''
    import datetime as _dt
    if isinstance(cell, _dt.datetime):
        if cell.hour == 0 and cell.minute == 0 and cell.second == 0 and cell.microsecond == 0:
            return cell.strftime("%Y-%m-%d")
        return cell.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(cell, _dt.date):
        return cell.strftime("%Y-%m-%d")
    if isinstance(cell, _dt.time):
        return cell.strftime("%H:%M:%S")
    if isinstance(cell, float) and math.isfinite(cell) and cell.is_integer():
        return str(int(cell))
    return str(cell)


def excel_to_markdown(xlsx_data):
    """将Excel数据转换为Markdown表格（仅依赖 openpyxl，无需 pandas）"""
    try:
        import openpyxl

        # XLSX 本身也是 ZIP：DOCX 外层的 entry 限制只能约束
        # xlsx 字节大小，无法防止其内部条目解压膨胀。交给
        # openpyxl 前先对内层 ZIP 无条件执行恶意特征校验。
        with zipfile.ZipFile(io.BytesIO(xlsx_data), "r") as xlsx_zip:
            validate_docx_zip_security(xlsx_zip, on_limit="reject")

        def normalize_rows(raw_rows: List[List[str]]) -> List[List[str]]:
            if not raw_rows:
                return []

            rows = [r for r in raw_rows if any(c.strip() for c in r)]
            if not rows:
                return []

            col_count = max(len(r) for r in rows)
            rows = [r + [''] * (col_count - len(r)) for r in rows]
            non_empty_cols = [j for j in range(col_count) if any(rows[i][j].strip() for i in range(len(rows)))]
            if not non_empty_cols:
                return []
            return [[r[j] for j in non_empty_cols] for r in rows]

        def apply_merged_cells(ws, raw_rows: List[List[str]]) -> List[List[str]]:
            """将合并单元格展开为 Markdown 管道表可读的全展开网格。"""
            if not raw_rows:
                return raw_rows

            for merged in ws.merged_cells.ranges:
                min_row, max_row = merged.min_row, merged.max_row
                min_col, max_col = merged.min_col, merged.max_col

                row_idx = min_row - 1
                col_idx = min_col - 1
                if row_idx >= len(raw_rows):
                    continue
                if col_idx >= len(raw_rows[row_idx]):
                    continue

                anchor_value = raw_rows[row_idx][col_idx]
                if not anchor_value:
                    continue

                for row_no in range(min_row, max_row + 1):
                    i = row_no - 1
                    if i >= len(raw_rows):
                        continue
                    if len(raw_rows[i]) < max_col:
                        raw_rows[i].extend([''] * (max_col - len(raw_rows[i])))
                    for col_no in range(min_col, max_col + 1):
                        raw_rows[i][col_no - 1] = anchor_value
            return raw_rows

        def sheet_to_rows(ws) -> tuple[List[List[str]], List[str], int]:
            raw_rows: List[List[str]] = []
            for row in ws.iter_rows(values_only=True):
                raw_rows.append([
                    _normalize_markdown_cell_text(_format_cell_value(cell))
                    for cell in row
                ])
            score_rows = normalize_rows(raw_rows)
            score = sum(1 for r in score_rows for c in r if c.strip())
            merge_ranges = [str(rng) for rng in ws.merged_cells.ranges]
            raw_rows = apply_merged_cells(ws, raw_rows)
            return normalize_rows(raw_rows), merge_ranges, score

        # 不使用 read_only=True：部分嵌入工作簿的维度元数据异常，read_only 模式会把表格截断成 1x1。
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_data), read_only=False, data_only=True)
        best_rows = []
        best_merge_ranges: List[str] = []
        best_score = -1
        for ws in wb.worksheets:
            rows, merge_ranges, score = sheet_to_rows(ws)
            if not rows:
                continue
            if score > best_score:
                best_rows = rows
                best_merge_ranges = merge_ranges
                best_score = score
        wb.close()

        if not best_rows:
            return None

        header = '| ' + ' | '.join(best_rows[0]) + ' |'
        separator = '| ' + ' | '.join(['---'] * len(best_rows[0])) + ' |'
        body_lines = ['| ' + ' | '.join(r) + ' |' for r in best_rows[1:]]
        table_text = header + '\n' + separator + '\n' + '\n'.join(body_lines)
        if best_merge_ranges:
            ranges_text = ", ".join(best_merge_ranges)
            return f"> merge_ranges: {ranges_text}\n\n{table_text}"
        return table_text

    except DocxSecurityError:
        raise
    except TimeoutError:
        # 批处理 SIGALRM 的单文档超时必须继续上抛：被降级为“无表格”会把
        # 超时文档误记为转换成功并写 sentinel（见 BUG-025）
        raise
    except Exception as e:
        logger.warning("Excel转Markdown失败: %s", e)
        return None


def detect_image_format(image_data):
    """检测图片的真实格式；无法识别时返回 None（调用方应保留原扩展名）"""
    if image_data[:8] == b'\x89PNG\r\n\x1a\n':
        return '.png'
    elif image_data[:2] == b'\xff\xd8':
        return '.jpeg'
    elif image_data[:6] in (b'GIF87a', b'GIF89a'):
        return '.gif'
    elif image_data[:4] == b'RIFF' and image_data[8:12] == b'WEBP':
        return '.webp'
    elif image_data[:2] == b'BM':
        return '.bmp'
    elif image_data[:4] in (b'II*\x00', b'MM\x00*'):
        return '.tiff'
    elif image_data[:4] == b'\xd7\xcd\xc6\x9a':
        return '.wmf'
    elif len(image_data) >= 44 and image_data[40:44] == b' EMF':
        return '.emf'
    return None


def extract_content_from_docx(docx_path, assets_dir, on_limit="reject", skip_state=None):
    """从docx中提取图片和Excel数据，并构建“内容hash -> 内容”的映射

    Args:
        on_limit: 可降级资源（图片数量/单图大小/单图像素/嵌入 Excel 大小）
            超限时的处置。"reject"（默认）抛 ResourceLimitExceeded 整篇拒绝；
            "skip" 仅跳过该资源继续转换。
        skip_state: 可选的可变 dict 输出参数。传入时写入 skipped 与
            media_processed（配额内的媒体条目数）、allowed_media_paths
            （提取与 Mammoth 回调共用的物理媒体白名单）、media_part_names
            （包内全部物理图片条目，含未引用的 word/media 及关系目标，
            供 Mammoth 回调识别媒体路径）、empty_directory_parts（空目录
            占位名，供回调排除伪图片）；省略时保持历史三项返回值契约。
            ZIP 级恶意特征不在此降级（validate 阶段已无条件拒绝）。

    返回:
        image_by_hash: { sha256_hex: "assets/xxx.png" }
        table_queue_by_hash: { sha256_hex: ["<md_table1>", "<md_table2>", ...] }
        table_repeat_by_hash: { sha256_hex: "<md_table>" }  # 队列耗尽时的稳定兜底
    """
    validate_on_limit(on_limit)
    skip_mode = on_limit == "skip"
    limits = DOCX_SECURITY_LIMITS
    skipped = []
    media_processed = 0
    image_by_hash = {}
    table_queue_by_hash = defaultdict(list)
    table_repeat_by_hash = {}

    # 解析关系，找出Excel和预览图的对应
    excel_to_preview, preview_to_excel, ordered_pairs = parse_relationships(docx_path)

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        empty_directory_parts = {
            info.filename for info in zip_ref.infolist()
            if info.is_dir() and info.file_size == 0
        }
        excel_md_by_path = {}
        table_preview_paths = set()
        # 物理图片条目 = word/media/* ∪ 各 Mammoth 正文 part（动态解析，
        # 见 _document_image_part_names）关系文件的 image 关系目标，保持
        # ZIP 条目顺序（skip 配额按此顺序截取）。仅按目录前缀或硬编码
        # part 清单枚举会被改写 Target 的文档绕过数量/大小/像素防线
        # （见 BUG-026/BUG-042）。
        image_part_names = _document_image_part_names(zip_ref)
        media_paths = [
            info.filename for info in zip_ref.filelist
            if info.filename in image_part_names
        ]
        allowed_media_paths = set(media_paths)
        if skip_mode:
            allowed_media_paths = set(media_paths[:limits["image_count"]])
            media_processed = len(allowed_media_paths)
            if len(media_paths) > limits["image_count"]:
                skipped.append((
                    "图片条目*",
                    f"图片数量超过上限 {limits['image_count']}，剩余图片停止提取",
                ))
                logger.warning(
                    "图片数量超过上限 %d，配额外图片条目不会读取",
                    limits["image_count"],
                )

        # 先提取所有 Excel 文件的数据并转换为 Markdown
        for file_info in zip_ref.filelist:
            if file_info.filename.startswith('word/embeddings/') and file_info.filename.lower().endswith('.xlsx'):
                excel_file = file_info.filename
                try:
                    xlsx_data = read_zip_entry_bounded(
                        zip_ref, excel_file, limits["embedded_excel_size"],
                        error_cls=ResourceLimitExceeded,
                    )
                except ResourceLimitExceeded:
                    if not skip_mode:
                        raise
                    skipped.append((excel_file, "嵌入 Excel 超过大小上限"))
                    logger.warning("跳过超大嵌入 Excel（表格不转换，正文保留）: %s", excel_file)
                    continue

                markdown_table = excel_to_markdown(xlsx_data)
                if markdown_table:
                    excel_md_by_path[excel_file] = markdown_table
                else:
                    logger.warning("Excel表格转换失败（将保留预览图）: %s", excel_file)

        # 建立预览图 hash -> 表格队列（同一预览图内容可对应多个表格）
        pairs = ordered_pairs if ordered_pairs else [(e, p) for e, p in excel_to_preview.items()]
        for excel_path, preview_path in pairs:
            table_md = excel_md_by_path.get(excel_path)
            if not table_md:
                continue
            if preview_path not in zip_ref.namelist():
                continue
            table_preview_paths.add(preview_path)
            if skip_mode and preview_path not in allowed_media_paths:
                continue
            try:
                preview_data = _read_media_image(zip_ref, preview_path)
            except ResourceLimitExceeded as exc:
                if not skip_mode:
                    raise
                skipped.append((preview_path, str(exc)))
                logger.warning("跳过超限预览图（对应表格不注入，原位置显示跳过说明）: %s", exc)
                continue
            digest = hashlib.sha256(preview_data).hexdigest()
            table_queue_by_hash[digest].append(table_md)
            table_repeat_by_hash[digest] = table_md
            logger.info("转换Excel为表格: %s", excel_path)

        # 处理图片（关系目标可能位于 word/media 之外的任意包内路径）
        for file_info in zip_ref.filelist:
            if file_info.filename in image_part_names:
                # 目录名条目（word/media/image1.png/）去掉尾 / 取文件名
                image_name = os.path.basename(file_info.filename.rstrip("/"))

                # 检查这个图片是否是Excel的预览图
                if file_info.filename in table_preview_paths:
                    continue

                if skip_mode and file_info.filename not in allowed_media_paths:
                    continue

                # 普通图片，直接提取（读取路径上执行大小/像素防线）
                try:
                    image_data = _read_media_image(zip_ref, file_info.filename)
                except ResourceLimitExceeded as exc:
                    if not skip_mode:
                        raise
                    skipped.append((file_info.filename, str(exc)))
                    logger.warning("跳过超限图片: %s", exc)
                    continue
                digest = hashlib.sha256(image_data).hexdigest()
                
                # 检测真实的图片格式并修正扩展名；无法识别时保留原扩展名，
                # 避免把 WMF/EMF 等格式误写成 .png 造成文件损坏
                original_ext = os.path.splitext(image_name)[1].lower()
                actual_ext = detect_image_format(image_data) or original_ext or ".png"
                base_name = os.path.splitext(image_name)[0]

                # 候选路径循环分配：只有普通文件且字节一致才复用；符号链接
                # （含悬空）、目录或不同内容的占用继续换名，避免跟随链接写
                # 入外部或引用错误文件（见 BUG-023/BUG-027）。分配与独占
                # 写入之间被第三方占用的候选同样换名重试（见 BUG-047）。
                while True:
                    corrected_name, image_path = _allocate_asset_path(
                        assets_dir, base_name, actual_ext, digest, image_data)
                    try:
                        _write_asset_file_exclusive(image_path, image_data)
                        break
                    except FileExistsError:
                        continue

                image_by_hash.setdefault(digest, f"assets/{corrected_name}")
                logger.info("提取图片: %s", corrected_name)
    
    if skip_state is not None:
        skip_state.clear()
        skip_state.update({
            "skipped": skipped,
            "media_processed": media_processed,
            "allowed_media_paths": allowed_media_paths,
            "media_part_names": image_part_names,
            "empty_directory_parts": empty_directory_parts,
        })
    return image_by_hash, table_queue_by_hash, table_repeat_by_hash


def extract_textbox_content(docx_path: str) -> List[str]:
    """从 DOCX 的 document.xml 中提取文本框 (<w:txbxContent>) 内的纯文本。

    mammoth 通常会忽略 text box / shape 中的内容，此函数作为补充。
    返回非空文本块列表。
    """
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns_wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    tag_txbx = f"{{{ns_w}}}txbxContent"
    tag_txbx_wps = f"{{{ns_wps}}}txbxContent"
    tag_t = f"{{{ns_w}}}t"
    tag_p = f"{{{ns_w}}}p"

    blocks: List[str] = []
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            doc_xml = _safe_xml_fromstring(zf.read("word/document.xml"))

        for txbx_tag in (tag_txbx, tag_txbx_wps):
            for txbx in doc_xml.iter(txbx_tag):
                paras = []
                for p in txbx.findall(f".//{tag_p}"):
                    text = "".join((t.text or "") for t in p.findall(f".//{tag_t}"))
                    text = text.strip()
                    if text:
                        paras.append(text)
                if paras:
                    blocks.append("\n".join(paras))
    except TimeoutError:
        raise  # 超时信号不得被降级捕获吞掉（见 BUG-025）
    except Exception:
        pass
    return blocks


def extract_math_text(docx_path: str) -> List[str]:
    """从 DOCX 的 document.xml 中提取 OMML 数学公式的纯文本内容。

    完整的 OMML→LaTeX 转换极为复杂，此函数仅提取公式中的文本节点，
    用 $ 包裹作为占位标记，便于下游人工校正。
    """
    ns_m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tag_omath = f"{{{ns_m}}}oMath"
    tag_omath_para = f"{{{ns_m}}}oMathPara"
    tag_t_m = f"{{{ns_m}}}t"
    tag_t_w = f"{{{ns_w}}}t"

    formulas: List[str] = []
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            doc_xml = _safe_xml_fromstring(zf.read("word/document.xml"))

        seen = set()
        for parent_tag in (tag_omath_para, tag_omath):
            for node in doc_xml.iter(parent_tag):
                node_id = id(node)
                if node_id in seen:
                    continue
                seen.add(node_id)
                parts = []
                for t in node.iter():
                    if t.tag in (tag_t_m, tag_t_w) and t.text:
                        parts.append(t.text)
                text = "".join(parts).strip()
                if text:
                    formulas.append(text)
                for child in node.iter(tag_omath):
                    seen.add(id(child))
    except TimeoutError:
        raise  # 超时信号不得被降级捕获吞掉（见 BUG-025）
    except Exception:
        pass
    return formulas


def convert_docx_to_markdown(docx_path, output_dir, create_subfolder=True, output_name=None,
                             on_limit="reject"):
    """将docx转换为markdown

    Args:
        docx_path: DOCX 文件路径
        output_dir: 输出目录路径
        create_subfolder: 是否在输出目录下创建以文件名命名的子文件夹（默认 True）
        output_name: 自定义输出命名（默认 None 用源文件名）。末尾 .docx 自动去除，
            其他点号后缀保留。
            经 sanitize_stem 清洗后统一用于子文件夹名、.md 文件名与 sentinel 的
            folder_name 字段，三处保持一致。适合 Web 上传等需要以用户原始
            文件名命名的场景；批处理不使用本参数（按源文件名命名）
        on_limit: 可降级资源（图片数量/单图大小/单图像素/嵌入 Excel 大小）超限
            处置。"reject"（默认）抛 DocxSecurityError 整篇拒绝，与历史行为一致；
            "skip" 仅跳过超限资源继续转换（超限图片原位置写入可见跳过说明，
            不落盘）。ZIP bomb 等恶意特征在任何模式下都整篇拒绝
    """
    validate_on_limit(on_limit)
    skip_mode = on_limit == "skip"
    
    # 先校验输入，避免 BadZipFile 直接中断并泄漏底层异常。
    # 安全校验（DocxSecurityError）在结构校验之后执行：结构非法报格式错误，
    # 结构合法但资源超限报安全错误。二者均为 ValueError 子类，上层可按需
    # 区分——安全错误表示输入恶意/异常，不可降级重试。
    try:
        with zipfile.ZipFile(docx_path, "r") as zip_ref:
            if "word/document.xml" not in zip_ref.namelist():
                raise ValueError(f"输入文件不是有效的 DOCX（缺少 word/document.xml）: {docx_path}")
            validate_docx_zip_security(zip_ref, on_limit=on_limit)
    except zipfile.BadZipFile as exc:
        raise ValueError(f"输入文件不是有效的 DOCX/ZIP: {docx_path}") from exc

    # 记录源文件哈希：转换全部成功后写入 .converted sentinel，
    # 供批处理判断“输出完整且与当前源一致”（源变更后自动重转）。
    source_sha256 = sha256_file(docx_path)

    # 输出命名：优先显式指定的 output_name（如 Web 上传场景的用户原始文件名），
    # 统一经 sanitize_stem 清洗后作为 folder_name，下游目录/文件名/sentinel 均引用它
    if output_name is not None:
        if not isinstance(output_name, str) or not output_name.strip():
            raise ValueError("output_name 必须是非空字符串")
        normalized_output_name = output_name.strip()
        if normalized_output_name.lower().endswith(".docx"):
            normalized_output_name = normalized_output_name[:-5]
        if not normalized_output_name.strip():
            raise ValueError("output_name 去除 .docx 后不能为空")
        folder_name = sanitize_stem(normalized_output_name)
    else:
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        folder_name = sanitize_stem(base_name)
    
    # 确定最终输出目录
    if create_subfolder:
        final_output_dir = os.path.join(output_dir, folder_name)
    else:
        final_output_dir = output_dir

    # 输出写入边界（见 BUG-019）：文档输出子目录必须是普通目录，不允许
    # 符号链接——否则预置链接会把 assets 清理、Markdown 与 sentinel 写入
    # 导向输出目录之外，prune_stale_assets 甚至会删除外部文件。最终真实
    # 路径还须仍位于输出根目录内（防御中间组件被替换为链接）。
    # 目录状态一律用 _path_status 直接 lstat：os.path.lexists/islink/isdir
    # 内部捕获 OSError（含 TimeoutError），会吞掉批处理超时信号（见 BUG-035）。
    if create_subfolder:
        st = _path_status(final_output_dir)
        if st is not None and not stat.S_ISDIR(st.st_mode):
            raise ValueError(
                f"输出子目录已存在且不是普通目录（不允许符号链接）: {final_output_dir}")
    # realpath 一律走 _safe_realpath：os.path.realpath 非 strict 模式
    # 内部 except OSError 会吞掉组件 lstat 的 TimeoutError（见 BUG-039）
    output_root_real = _safe_realpath(output_dir)
    final_real = _safe_realpath(final_output_dir)
    if os.path.commonpath([output_root_real, final_real]) != output_root_real:
        raise ValueError(
            f"输出目录解析后的真实路径越出输出根目录: {final_output_dir} -> {final_real}")

    # 创建输出目录：仅在缺失时调用 makedirs，避免其 exist_ok 兜底里的
    # isdir 吞掉超时信号（见 BUG-035）
    if _path_status(final_output_dir) is None:
        os.makedirs(final_output_dir)
    assets_dir = os.path.join(final_output_dir, 'assets')
    st = _path_status(assets_dir)
    if st is not None and not stat.S_ISDIR(st.st_mode):
        raise ValueError(f"assets 目录必须是普通目录（不允许符号链接）: {assets_dir}")
    if st is None:
        os.makedirs(assets_dir)
    
    # 提取图片和Excel表格
    logger.info("正在提取内容...")
    skip_state = {}
    image_by_hash, table_queue_by_hash, table_repeat_by_hash = extract_content_from_docx(
        docx_path, assets_dir, on_limit=on_limit, skip_state=skip_state
    )
    skipped_resources = skip_state["skipped"]
    # Mammoth 回调必须复用 ZIP 提取阶段选定的同一份物理媒体
    # 白名单，不能按回调顺序重新分配配额，否则可落盘 2x 上限。
    allowed_media_paths = set(skip_state["allowed_media_paths"])
    media_part_names = set(skip_state["media_part_names"])
    empty_directory_parts = set(skip_state["empty_directory_parts"])
    media_quota_exceeded = any(
        "图片数量超过上限" in reason for _, reason in skipped_resources
    )
    # Mammoth 回调兜底写盘的剩余配额：扫描侧已确认的物理图片之外的未知
    # 图片（zip 条目与回调数据不一致、或动态解析口径与 mammoth 实际读取
    # 出现偏差时）也不得使总数突破 image_count（见 BUG-042 fail-closed）。
    fallback_image_budget = [
        DOCX_SECURITY_LIMITS["image_count"] - len(allowed_media_paths)]
    noted_skipped = set()
    if any("图片数量超过上限" in reason for _, reason in skipped_resources):
        # 提取阶段已用一条记录汇总“剩余图片”；Mammoth 仍会
        # 逐引用回调，这里只返回占位符，不再按图片 hash 增长清单。
        noted_skipped.add("count")

    def _skipped_image_result(reason, image_data):
        """记录一次回调侧资源跳过，并返回可见跳过占位的 img src。"""
        marker = reason if reason == "count" else hashlib.sha256(image_data).hexdigest()[:8]
        if marker not in noted_skipped:
            noted_skipped.add(marker)
            skipped_resources.append(
                (f"文档引用图片#{marker}", SKIPPED_IMAGE_NOTE[reason]))
        return {"src": f"__SKIPPED_IMAGE_{reason}__"}

    table_md_by_placeholder = {}
    table_seq = [0]
    heading_level_map = extract_heading_level_map(docx_path)
    
    # 使用mammoth转换为HTML
    logger.info("正在转换文档...")

    def convert_image(image):
        """根据图片内容hash，返回对应的assets路径或表格占位符"""
        # 空目录不是图片：两种策略均在配额判断及回调读取前排除，防止
        # 已被扫描排除的目录经兜底写盘复活（见 BUG-049）。
        if _mammoth_embedded_media_key(image, empty_directory_parts) is not None:
            return _skipped_image_result("directory", b"")
        if skip_mode:
            media_key = _mammoth_embedded_media_key(image, media_part_names)
            if media_key is not None and media_key not in allowed_media_paths:
                return _skipped_image_result("count", b"")
            if media_key is None and media_quota_exceeded:
                # Mammoth 版本/图片类型无法暴露物理路径时选择安全回退：
                # 不在已确认超配额的文档中允许未知回调兜底落盘。
                return _skipped_image_result("count", b"")

        # 两种策略的回调读取都执行大小/像素防线：reject 直接整篇拒绝；
        # skip 模式下已跳过的超限图片不会出现在 image_by_hash，若放行到
        # 下方兜底写盘分支即绕过防线，故超限直接返回可见跳过占位（见
        # BUG-026：reject 回调此前无界 read 且不检查像素）。
        with image.open() as image_bytes:
            image_data = image_bytes.read(DOCX_SECURITY_LIMITS["image_file_size"] + 1)
        if len(image_data) > DOCX_SECURITY_LIMITS["image_file_size"]:
            if skip_mode:
                return _skipped_image_result("size", image_data)
            raise ResourceLimitExceeded(
                "图片超过单图大小上限 "
                f"{_fmt_bytes(DOCX_SECURITY_LIMITS['image_file_size'])}: mammoth 回调"
            )
        pixels = image_pixel_count(image_data)
        if pixels is not None and pixels > DOCX_SECURITY_LIMITS["image_pixels"]:
            if skip_mode:
                return _skipped_image_result("pixels", image_data)
            raise ResourceLimitExceeded(
                f"图片像素超过上限 {DOCX_SECURITY_LIMITS['image_pixels']}: "
                f"mammoth 回调（{pixels} 像素）"
            )
        digest = hashlib.sha256(image_data).hexdigest()

        table_queue = table_queue_by_hash.get(digest)
        if table_queue:
            # 若仅剩一个元素则不再弹出，确保同一预览图多次出现时仍稳定替换为表格
            table_md = table_queue[0] if len(table_queue) == 1 else table_queue.pop(0)
            placeholder = f"__TABLE_PLACEHOLDER_{digest}_{table_seq[0]}__"
            table_seq[0] += 1
            table_md_by_placeholder[placeholder] = table_md
            return {"src": placeholder}

        # 防御性兜底：当前队列逻辑保证最后一个元素不会被弹出，因此此分支在
        # 正常流程中不会触发。保留此分支作为安全网，以防未来队列策略调整后
        # 队列被完全消耗的情况，确保仍能稳定替换为表格而非退化为普通图片。
        if digest in table_repeat_by_hash:
            table_md = table_repeat_by_hash[digest]
            placeholder = f"__TABLE_PLACEHOLDER_{digest}_{table_seq[0]}__"
            table_seq[0] += 1
            table_md_by_placeholder[placeholder] = table_md
            return {"src": placeholder}

        image_src = image_by_hash.get(digest)
        if image_src:
            return {"src": image_src}

        # 兜底：某些情况下zip里的图片与mammoth回调数据不一致，直接按hash写入assets。
        # 候选分配与提取循环同一套规则（见 BUG-027）
        if fallback_image_budget[0] <= 0:
            # 扫描侧未见过的图片也不得使总数突破上限（见 BUG-042）
            if skip_mode:
                return _skipped_image_result("count", b"")
            raise ResourceLimitExceeded(
                "图片数量超过上限 "
                f"{DOCX_SECURITY_LIMITS['image_count']}: mammoth 回调兜底"
            )
        fallback_image_budget[0] -= 1
        ext = detect_image_format(image_data)
        if not ext:
            # 依据 mammoth 提供的 content_type 推断扩展名，仍未知则保留二进制原名
            content_subtype = (getattr(image, "content_type", "") or "").split("/")[-1].lower()
            ext = {
                "jpeg": ".jpeg", "jpg": ".jpeg", "png": ".png", "gif": ".gif",
                "webp": ".webp", "bmp": ".bmp", "tiff": ".tiff",
                "x-wmf": ".wmf", "x-emf": ".emf",
            }.get(content_subtype, ".bin")
        # 分配与独占写入之间被第三方占用的候选换名重试（见 BUG-047）
        while True:
            filename, image_path = _allocate_asset_path(
                assets_dir, f"image_{digest[:16]}", ext, digest, image_data)
            try:
                _write_asset_file_exclusive(image_path, image_data)
                break
            except FileExistsError:
                continue
        image_by_hash[digest] = f"assets/{filename}"
        return {"src": f"assets/{filename}"}
    
    with open(docx_path, 'rb') as docx_file:
        import mammoth

        result = mammoth.convert_to_html(
            docx_file,
            convert_image=mammoth.images.img_element(convert_image)
        )
        html = result.value
        for msg in getattr(result, "messages", []) or []:
            logger.debug("mammoth提示: %s", msg)
    
    # 将HTML转换为Markdown
    markdown = html_to_markdown(html, heading_level_map)
    
    # 替换表格占位符
    for placeholder_key, table_md in table_md_by_placeholder.items():
        placeholder = f"![]({placeholder_key})"
        markdown = markdown.replace(placeholder, f"\n\n{table_md}\n\n")

    # 自检：占位符未被替换通常意味着 mammoth 输出的 HTML 结构与预期不符，
    # 保留占位符并告警，便于发现未知文档结构的退化情况。
    leftover = re.findall(r"__TABLE_PLACEHOLDER_[0-9a-f]+_\d+__", markdown)
    if leftover:
        logger.warning("有 %d 个表格占位符未能替换为表格（请检查输出）", len(leftover))

    # 将 on_limit=skip 的超限图片、两种策略的空目录占位替换为可见说明
    # （不引用不存在的资源文件，跳过在输出中可见、可审计）
    markdown = re.sub(
        r"!\[[^\]]*\]\(__SKIPPED_IMAGE_([a-z]+)__\)",
        lambda m: f"*【图片已跳过：{SKIPPED_IMAGE_NOTE.get(m.group(1), m.group(1))}】*",
        markdown,
    )

    # 移除嵌入 Excel 替换后残留的预览图说明文本
    markdown = re.sub(
        r"\n+\*{0,2}点击图片可查看完整电子表格\*{0,2}\s*\n",
        "\n",
        markdown,
    )

    # 追加 mammoth 未能提取的文本框内容
    textbox_blocks = extract_textbox_content(docx_path)
    if textbox_blocks:
        # 检查主体中是否已包含文本框文本（mammoth 有时也能提取部分文本框）；
        # 主体经实体保护管线后以实体形式存在，两种形式都查避免重复追加
        def _already_in_markdown(line: str) -> bool:
            return line in markdown or _escape_html(line, quote=False) in markdown

        missing = [b for b in textbox_blocks if not _already_in_markdown(b.splitlines()[0])]
        if missing:
            markdown += "\n\n---\n\n> **\\[文本框内容\\]**\n\n"
            for block in missing:
                # 文本框内容在此拼接，已绕过 html_to_markdown 的实体保护
                # 管线：字面 <time>/<table> 等必须实体化，否则渲染时被当作
                # 原始 HTML 吞掉显示（见 BUG-046，处理方式与脚注体一致）
                markdown += f"> {_escape_html(block, quote=False)}\n>\n"
            logger.info("追加了 %d 个文本框内容", len(missing))

    # 追加 mammoth 未能提取的数学公式
    math_formulas = extract_math_text(docx_path)
    if math_formulas:
        missing_math = [f for f in math_formulas if f not in markdown]
        if missing_math:
            markdown += "\n\n---\n\n> **\\[数学公式\\]**\n\n"
            for formula in missing_math:
                markdown += f"> $$ {formula} $$\n>\n"
            logger.info("追加了 %d 个数学公式", len(missing_math))

    md_path = os.path.join(final_output_dir, f"{folder_name}.md")

    # 随机独占临时文件 + 原子替换：目标位置的预置符号链接只被替换目录项，
    # 不会被跟随写入外部文件（见 BUG-023）
    _atomic_write_text(md_path, markdown)

    # 子目录模式的 assets 归当前文档独占，可清理旧产物。
    # 平铺模式可能由多份 Markdown 共享 assets，不删除未引用文件。
    if create_subfolder:
        prune_stale_assets(assets_dir, image_by_hash.values())
    write_conversion_sentinel(
        final_output_dir, folder_name, source_sha256, on_limit=on_limit)

    if skip_mode and skipped_resources:
        logger.warning(
            "on_limit=skip：共跳过 %d 项超限资源（正文与其余资源已保留）", len(skipped_resources))
        for name, reason in skipped_resources:
            logger.warning("  已跳过: %s（%s）", name, reason)

    logger.info("转换完成: %s", md_path)
    return md_path


class _FootnoteHTMLParser(HTMLParser):
    """按 li 嵌套深度提取完整脚注，仅记录需从原文删除的区间。"""

    def __init__(self, source: str):
        super().__init__(convert_charrefs=True)
        self.source = source
        self.line_offsets = [0] + [match.end() for match in re.finditer("\n", source)]
        self.footnote_bodies: Dict[str, str] = {}
        self.spans: List[Tuple[int, int]] = []
        self.fid: Optional[str] = None
        self.li_depth = 0
        self.start = 0
        self.body: List[str] = []

    def _source_offset(self) -> int:
        line, column = self.getpos()
        return self.line_offsets[line - 1] + column

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if self.fid is None:
            fid = re.fullmatch(r"footnote-(\d+)", attrs.get("id") or "", flags=re.IGNORECASE)
            if tag != "li" or not fid:
                return
            self.fid = fid.group(1)
            self.li_depth = 1
            self.start = self._source_offset()
            self.body = []
            return

        if tag == "li":
            self.li_depth += 1
        if tag in {"p", "div", "li", "ul", "ol", "br"}:
            self.body.append(" ")
        elif tag == "img" and attrs.get("src") is not None:
            # 和数据节点一样保护实体，避免路径被后续 HTML 管线重新解释。
            self.body.append(_escape_html(f"![]({attrs['src']})", quote=False))

    def handle_data(self, data):
        if self.fid is not None:
            # HTMLParser 已解码实体；重新转义以保留字面 <time>/<table>，
            # 由 html_to_markdown 的最终实体保护管线统一处理（BUG-031）。
            self.body.append(_escape_html(data, quote=False))

    def handle_endtag(self, tag):
        if self.fid is None:
            return
        if tag in {"p", "div", "li", "ul", "ol"}:
            self.body.append(" ")
        if tag != "li":
            return
        self.li_depth -= 1
        if self.li_depth:
            return

        # 空白来自原文本、字符实体或块级边界，统一压成单行定义（BUG-032）。
        body = _WHITESPACE_RE.sub(" ", "".join(self.body).replace("↑", "")).strip()
        if body:
            self.footnote_bodies[self.fid] = body
        end = self.source.index(">", self._source_offset()) + 1
        self.spans.append((self.start, end))
        self.fid = None


def _convert_footnotes(html: str) -> str:
    """将 mammoth 生成的脚注 HTML 转换为 Markdown 脚注语法。

    mammoth 输出格式：
      正文引用: <sup><a href="#footnote-N" id="footnote-ref-N">[N]</a></sup>
      文末列表: <li id="footnote-N"><p>text <a href="#footnote-ref-N">↑</a></p></li>

    脚注体在其余 HTML 转换之前按完整 li 节点抽出，支持内部嵌套列表。
    图片收成内联 Markdown，skip 占位沿用正文替换；文本实体保留到
    管线末尾，全部空白归一化，使图片与正文都保持在 [^N]: 同一行。
    """
    parser = _FootnoteHTMLParser(html)
    parser.feed(html)
    parser.close()
    footnote_bodies = parser.footnote_bodies
    parts = []
    cursor = 0
    for start, end in parser.spans:
        parts.append(html[cursor:start])
        cursor = end
    parts.append(html[cursor:])
    html = "".join(parts)

    html = re.sub(
        r"<sup>\s*<a\b[^>]*href\s*=\s*[\"']?#footnote-(\d+)[\"']?[^>]*>"
        r"\s*\[\d+\]\s*</a>\s*</sup>",
        lambda m: f"[^{m.group(1)}]",
        html,
        flags=re.IGNORECASE,
    )

    if footnote_bodies:
        footer = "\n\n---\n\n"
        for fid in sorted(footnote_bodies, key=int):
            footer += f"[^{fid}]: {footnote_bodies[fid]}\n"
        html += footer

    return html


def html_to_markdown(html, heading_level_map: Optional[Dict[str, int]] = None):
    """将HTML转换为Markdown"""

    html = _convert_footnotes(html)

    # 处理标题
    html = re.sub(r'<h1[^>]*>(.*?)</h1>', r'# \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h2[^>]*>(.*?)</h2>', r'## \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h3[^>]*>(.*?)</h3>', r'### \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h4[^>]*>(.*?)</h4>', r'#### \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h5[^>]*>(.*?)</h5>', r'##### \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h6[^>]*>(.*?)</h6>', r'###### \1\n\n', html, flags=re.DOTALL)
    
    # 优先按 DOCX 原始 heading 样式提升标题层级（主流程），避免纯文本启发式误判。
    if heading_level_map:
        def _replace_anchored_heading(match):
            heading_id = match.group("id1") or match.group("id2") or match.group("id3") or ""
            content = match.group("content")
            level = heading_level_map.get(heading_id)
            if not level:
                return match.group(0)
            text = re.sub(r"<[^>]+>", "", content, flags=re.DOTALL)
            text = unescape(text).strip()
            if not text:
                return ""
            # unescape 会把标题里的字面标签文本（&lt;table&gt; 等）还原成
            # 裸标签：裸 <table> 会被表格转换吞掉其后全部内容（见 BUG-040），
            # <a>/<time> 等会被当作真实标签删除或改写成链接（见 BUG-044）。
            # 重新实体化，由管线末尾的实体保护还原为可见的字面文本。
            text = text.replace("<", "&lt;").replace(">", "&gt;")
            return f"{'#' * level} {text}\n\n"

        html = re.sub(
            (
                r"<p[^>]*>\s*"
                r"<a[^>]*\bid\s*=\s*(?:\"(?P<id1>heading_\d+)\"|'(?P<id2>heading_\d+)'|(?P<id3>heading_\d+))[^>]*>"
                r"\s*</a>(?P<content>.*?)</p>"
            ),
            _replace_anchored_heading,
            html,
            flags=re.DOTALL | re.IGNORECASE,
        )

    # 处理粗体和斜体
    html = re.sub(r'<strong>(.*?)</strong>', r'**\1**', html, flags=re.DOTALL)
    html = re.sub(r'<b>(.*?)</b>', r'**\1**', html, flags=re.DOTALL)
    html = re.sub(r'<em>(.*?)</em>', r'*\1*', html, flags=re.DOTALL)
    html = re.sub(r'<i>(.*?)</i>', r'*\1*', html, flags=re.DOTALL)
    
    def _img_markdown(match):
        src = match.group("src1") or match.group("src2") or match.group("src3") or ""
        return f"![]({src})"

    # 处理链接（支持双引号、单引号、无引号三种 href 写法）。
    # 先于图片处理：链接内的图片用内联语法就地替换，避免图片自带的块级换行
    # 落在 `[` 与 `](url)` 之间，把链接语法拆断成非法 Markdown。
    def _replace_link(match):
        href = match.group("href1") or match.group("href2") or match.group("href3") or ""
        text = _IMG_TAG_RE.sub(_img_markdown, match.group("text"))
        return f"[{text}]({href})"

    html = re.sub(
        (
            r"<a\b[^>]*\bhref\s*=\s*"
            r"(?:\"(?P<href1>[^\"]*)\"|'(?P<href2>[^']*)'|(?P<href3>[^\s\"'=<>`]+))"
            r"[^>]*>(?P<text>.*?)</a>"
        ),
        _replace_link,
        html,
        flags=re.DOTALL | re.IGNORECASE,
    )

    # 处理链接之外的图片（块级语法，保持既有排版）
    html = _IMG_TAG_RE.sub(lambda match: _img_markdown(match) + "\n\n", html)

    # 先把HTML里的换行标签转为文本换行（需早于表格转换，避免改写表格里的 <br> 文本）
    html = re.sub(r'<br\s*/?>', '\n', html)

    # 先处理表格（必须在段落/列表转换之前）；按配对标签匹配以支持嵌套表格
    html = replace_html_tables(html)

    # 使用结构化解析处理列表，避免正则顺序导致的嵌套层级破坏。
    html = transform_html_lists_to_markdown(html)
    
    # 处理段落
    html = re.sub(r'<p[^>]*>(.*?)</p>', r'\1\n\n', html, flags=re.DOTALL)
    
    # 移除剩余的HTML标签（保留 <br> 供 Markdown 单元格换行显示）
    html = re.sub(r'<(?!br\s*/?)[^>]+>', '', html, flags=re.IGNORECASE)
    
    # 清理多余的空行
    html = re.sub(r'\n{3,}', '\n\n', html)
    
    html = html.replace('&nbsp;', ' ')
    # 统一 unescape 前保护“实体形式的标签样文本”（如 &lt;table&gt;、&lt;time&gt;、
    # &lt;采暖&gt;）：还原成裸 <...> 会被 Python Markdown 当作原始 HTML，
    # PDF 渲染时槽位文字不可见；标记后在 unescape 后恢复为实体，渲染为
    # 可见字面文本。首字符覆盖 /!?（<!-- -->、<? ?>、<![CDATA[]]> 等，
    # 见 BUG-045）与非 ASCII（中文槽位标签，见 BUG-048），仅排除空白
    # 首字符（"a < b" 类普通比较文本不是标签样，unescape 后保持裸形式
    # 维持 md 源文件可读性）；真正需要透传的 <br>（单元格换行）与
    # mammoth 的 sup/sub 等原始标签不在实体形式，不受影响
    # （见 BUG-020 验收补充）。
    html = re.sub(r"&lt;([^\s\x00\x01][^\x00\x01]*?)&gt;", "\x00\\1\x01", html)
    html = unescape(html)
    html = html.replace("\x00", "&lt;").replace("\x01", "&gt;")

    html = promote_numbered_bold_headings(html)
    html = promote_leading_bold_title(html)
    return html.strip()


if __name__ == '__main__':
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    parser = argparse.ArgumentParser(
        description="将单个 DOCX 文档转换为 Markdown（提取图片与嵌入 Excel 表格）")
    parser.add_argument("docx_path", help="DOCX 文件路径")
    parser.add_argument("output_dir", help="输出目录路径")
    parser.add_argument("--output-name", default=None,
                        help="自定义输出子文件夹与 .md 文件名（末尾 .docx 自动去除；默认用源文件名）")
    parser.add_argument("--on-limit", choices=("reject", "skip"), default="reject",
                        help="资源超限处置：reject 整篇拒绝（默认）；skip 仅跳过超限资源继续转换"
                             "（ZIP bomb 等恶意特征仍整篇拒绝）")
    args = parser.parse_args()

    docx_path = args.docx_path

    if not os.path.exists(docx_path):
        logger.error("文件不存在 - %s", docx_path)
        sys.exit(1)

    try:
        convert_docx_to_markdown(
            docx_path, args.output_dir,
            output_name=args.output_name, on_limit=args.on_limit)
    except DocxSecurityError as exc:
        logger.error("安全拒绝（输入恶意或资源超限，不重试）: %s", exc)
        sys.exit(2)
    except ValueError as exc:
        logger.error("输入错误: %s", exc)
        sys.exit(2)
