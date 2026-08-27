#!/usr/bin/env python3
"""
批量将目录下的所有docx文档转换为markdown格式
每个文档生成一个同名文件夹，包含md文件和assets子文件夹

跳过与清理语义：
  - 跳过需满足：输出目录 + md + 有效 `.converted` sentinel 齐备，
    且 sentinel 记录的源 SHA-256 与当前源文件一致（源变更自动重转，无需 --force）
  - 转换失败（含超时/安全拒绝）时清理输出目录，避免半成品被误判为已完成
"""

import logging
import os
import shutil
import signal
import sys
import glob

# 支持从同目录或作为模块导入
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from convert_docx import (
    DocxSecurityError,
    convert_docx_to_markdown,
    read_conversion_sentinel,
    sanitize_stem,
    sha256_file,
)

logger = logging.getLogger(__name__)

DEFAULT_TIMEOUT_SECONDS = 300


def remove_path(path):
    """删除文件/目录/符号链接；不存在时静默通过。"""
    if os.path.islink(path):
        os.remove(path)
    elif os.path.isdir(path):
        shutil.rmtree(path)
    elif os.path.exists(path):
        os.remove(path)


def _cleanup_failed_output(target_dir):
    """转换失败后清理半成品输出目录。"""
    if os.path.exists(target_dir) or os.path.islink(target_dir):
        logger.info("  清理未完成的输出目录: %s", target_dir)
        remove_path(target_dir)


def _is_output_complete(target_dir, folder_name, docx_path):
    """判断输出是否可信：目录 + md + 有效 sentinel，且源哈希与当前源一致。

    源哈希仅在 sentinel 存在且目录名匹配时才计算，避免对未转换文件白做 IO。
    """
    if not os.path.isdir(target_dir):
        return False
    if not os.path.isfile(os.path.join(target_dir, f"{folder_name}.md")):
        return False
    sentinel = read_conversion_sentinel(target_dir)
    if not sentinel:
        return False
    if sentinel.get("folder_name") != folder_name:
        return False
    return sentinel.get("source_sha256") == sha256_file(docx_path)


def _run_with_timeout(func, timeout_seconds):
    """POSIX 平台用 SIGALRM 给 func 加超时；不支持的环境降级为直接执行。

    Windows 无 SIGALRM 自动跳过；非主线程无法安装信号 handler（ValueError）
    同样降级。无论是否触发，finally 中恢复原 handler 并清除未决 alarm。
    """
    if not timeout_seconds or timeout_seconds <= 0:
        return func()
    if not hasattr(signal, "SIGALRM") or not hasattr(signal, "alarm"):
        return func()  # Windows 等无 SIGALRM 平台自动跳过

    def _on_alarm(signum, frame):
        raise TimeoutError(f"单文档转换超时（>{timeout_seconds}秒）")

    try:
        old_handler = signal.signal(signal.SIGALRM, _on_alarm)
    except (ValueError, OSError):
        return func()  # 非主线程安装失败，降级为无超时

    signal.alarm(timeout_seconds)
    try:
        return func()
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, old_handler)


def batch_convert(source_dir, output_dir, force=False, timeout=DEFAULT_TIMEOUT_SECONDS):
    """批量转换目录下的所有docx文件

    Args:
        source_dir: 源文件目录
        output_dir: 输出目录
        force: 为 True 时强制重新转换已存在的输出目录
        timeout: 单文档转换超时秒数（默认 300；<=0 不限制，
            仅 POSIX 主线程生效，Windows 自动跳过）
    """

    # 合并两种大小写扩展名并去重（macOS 大小写不敏感时 *.docx 已包含 .DOCX）
    seen = set()
    docx_files = []
    for path in glob.glob(os.path.join(source_dir, '*.docx')) + glob.glob(os.path.join(source_dir, '*.DOCX')):
        real = os.path.realpath(path)
        if real not in seen:
            seen.add(real)
            docx_files.append(path)

    if not docx_files:
        logger.warning("在 %s 中没有找到docx文件", source_dir)
        return

    logger.info("找到 %d 个docx文件待处理%s", len(docx_files),
                "（强制重新转换）" if force else "")

    success_count = 0
    fail_count = 0
    skip_count = 0

    os.makedirs(output_dir, exist_ok=True)

    for i, docx_path in enumerate(sorted(docx_files), 1):
        # 获取文件名（不含扩展名）作为输出文件夹名
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        folder_name = sanitize_stem(base_name)
        target_dir = os.path.join(output_dir, folder_name)

        logger.info("[%d/%d] 正在处理: %s", i, len(docx_files), base_name)

        # 检查是否已经完整转换且源未变更（--force 时跳过此检查）
        if not force and _is_output_complete(target_dir, folder_name, docx_path):
            logger.info("  已完成且源文件未变更，跳过（源变更会自动重转；强制重转用 --force）")
            skip_count += 1
            continue

        # 既有输出不可信（sentinel 缺失/无效/哈希不匹配）或 --force：清理后重转
        if os.path.exists(target_dir) or os.path.islink(target_dir):
            if not force:
                logger.info("  既有输出不完整或与当前源不一致，清理后重新转换")
            remove_path(target_dir)

        try:
            _run_with_timeout(
                lambda: convert_docx_to_markdown(docx_path, output_dir, create_subfolder=True),
                timeout,
            )
            logger.info("  完成")
            success_count += 1
        except TimeoutError as exc:
            logger.error("  超时: %s", exc)
            _cleanup_failed_output(target_dir)
            fail_count += 1
        except DocxSecurityError as exc:
            # 安全拒绝表示输入恶意/资源超限，清理后计为失败，不降级不重试
            logger.error("  安全拒绝（不重试）: %s", exc)
            _cleanup_failed_output(target_dir)
            fail_count += 1
        except Exception as exc:
            logger.error("  失败: %s", exc)
            _cleanup_failed_output(target_dir)
            fail_count += 1

    logger.info("处理完成: 成功 %d 个, 跳过 %d 个, 失败 %d 个",
                success_count, skip_count, fail_count)

if __name__ == '__main__':
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    parser = argparse.ArgumentParser(description="批量将目录下的所有docx文档转换为Markdown")
    parser.add_argument("source_dir", nargs="?", default="1-Reference", help="源文件目录（默认 1-Reference）")
    parser.add_argument("output_dir", nargs="?", default="2-Temp", help="输出目录（默认 2-Temp）")
    parser.add_argument("--force", action="store_true", help="强制重新转换已存在的输出目录")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS,
                        help="单文档转换超时秒数（默认 300；<=0 不限制；Windows 无 SIGALRM 自动跳过）")
    args = parser.parse_args()

    batch_convert(args.source_dir, args.output_dir, force=args.force, timeout=args.timeout)
