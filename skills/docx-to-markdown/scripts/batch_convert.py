#!/usr/bin/env python3
"""
批量将目录下的所有docx文档转换为markdown格式
每个文档生成一个同名文件夹，包含md文件和assets子文件夹

跳过与清理语义：
  - 跳过需满足：输出目录 + md + 有效 `.converted` sentinel 齐备，
    且 sentinel 记录的源 SHA-256 与当前源文件、on_limit 与当前请求一致
    （源或策略变更自动重转，无需 --force）
  - 转换失败（含超时/安全拒绝）时清理输出目录，避免半成品被误判为已完成

命名与退出码：
  - 扩展名大小写不敏感匹配（.docx/.DOCX/.Docx 等）
  - 批内两个文件名清洗后相同时（如 NFKC 归一化的 A 与 Ａ），后来者附加
    原始名短 hash 消歧，避免共用目录互相覆盖
  - 任一文档失败时 CLI 退出码为 1，全部成功/跳过为 0
"""

import hashlib
import logging
import os
import shutil
import signal
import sys

# 支持从同目录或作为模块导入
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from convert_docx import (  # noqa: E402  sys.path 调整必须先于本 import
    DocxSecurityError,
    convert_docx_to_markdown,
    read_conversion_sentinel,
    sanitize_stem,
    sha256_file,
    validate_on_limit,
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


def _is_output_complete(target_dir, folder_name, docx_path, on_limit="reject"):
    """判断输出是否可信：目录 + md + sentinel，且源哈希与策略一致。

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
    if sentinel.get("on_limit", "reject") != on_limit:
        return False
    return sentinel.get("source_sha256") == sha256_file(docx_path)


def _allocate_unique_folder_names(docx_files):
    """为批内文件分配互不冲突的输出 folder_name。

    sanitize_stem 的 NFKC 归一化（如全角Ａ→半角 A）是有意不加 hash 的弱映射，
    单文件场景的碰撞由 sentinel 哈希校验兜底；但同一批次内两个源文件若映射到
    同一目录，后处理者会删除并覆盖先处理者的完整结果，sentinel 无法阻止。

    这里在转换前统一分配：先到者保留自然名，后来者附加原始文件名的短 hash
    消歧（仅依赖文件名，跨批次稳定，不影响 skip/增量语义）；极端情况下仍
    冲突则再加序号。返回 [(docx_path, folder_name, output_name)]，
    output_name 仅在发生消歧时非 None，且必须满足
    sanitize_stem(output_name) == folder_name——转换器以 output_name 重算
    目录名，二者不一致时会写回旧目录造成覆盖（含 _N 序号推进的场景）。
    """
    used = set()
    plan = []
    for docx_path in docx_files:
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        folder_name = sanitize_stem(base_name)
        output_name = None
        if folder_name in used:
            digest = hashlib.sha256(base_name.encode("utf-8")).hexdigest()[:8]
            candidate_base = f"{folder_name}_{digest}"
            candidate = candidate_base
            folder_name = sanitize_stem(candidate)
            n = 2
            while folder_name in used:
                candidate = f"{candidate_base}_{n}"
                folder_name = sanitize_stem(candidate)
                n += 1
            output_name = candidate
            logger.warning(
                "  %s 与先前文件的输出名冲突，输出到: %s", base_name, folder_name)
        used.add(folder_name)
        plan.append((docx_path, folder_name, output_name))
    return plan


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


def batch_convert(source_dir, output_dir, force=False, timeout=DEFAULT_TIMEOUT_SECONDS,
                  on_limit="reject"):
    """批量转换目录下的所有docx文件

    Args:
        source_dir: 源文件目录
        output_dir: 输出目录
        force: 为 True 时强制重新转换已存在的输出目录
        timeout: 单文档转换超时秒数（默认 300；<=0 不限制，
            仅 POSIX 主线程生效，Windows 自动跳过）
        on_limit: 透传给 convert_docx_to_markdown 的资源超限处置。
            "reject"（默认）超限整篇拒绝计失败；"skip" 仅跳过超限资源继续
            转换（带超大附件的正常文档也能转出）。ZIP bomb 等恶意特征
            任何模式下都计失败不重试

    Returns:
        dict: {"success": 成功数, "skipped": 跳过数, "failed": 失败数}。
        任意文档转换失败时 failed > 0，CLI 据此返回退出码 1。
    """

    validate_on_limit(on_limit)

    # 大小写不敏感匹配 .docx/.DOCX/.Docx 等混合大小写扩展名；
    # realpath 去重吸收大小写不敏感文件系统（macOS）下的重复条目
    try:
        names = sorted(os.listdir(source_dir))
    except OSError:
        names = []
    seen = set()
    docx_files = []
    for name in names:
        if not name.lower().endswith(".docx"):
            continue
        path = os.path.join(source_dir, name)
        if not os.path.isfile(path):
            continue
        real = os.path.realpath(path)
        if real not in seen:
            seen.add(real)
            docx_files.append(path)

    if not docx_files:
        logger.warning("在 %s 中没有找到docx文件", source_dir)
        return {"success": 0, "skipped": 0, "failed": 0}

    logger.info("找到 %d 个docx文件待处理%s", len(docx_files),
                "（强制重新转换）" if force else "")

    success_count = 0
    fail_count = 0
    skip_count = 0

    os.makedirs(output_dir, exist_ok=True)

    # 转换前统一分配输出目录名，防止批内 NFKC 等弱映射碰撞互相覆盖
    plan = _allocate_unique_folder_names(sorted(docx_files))

    for i, (docx_path, folder_name, output_name) in enumerate(plan, 1):
        # 获取文件名（不含扩展名）用于日志展示
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        target_dir = os.path.join(output_dir, folder_name)

        logger.info("[%d/%d] 正在处理: %s", i, len(docx_files), base_name)

        # 检查是否已经完整转换且源未变更（--force 时跳过此检查）
        if not force and _is_output_complete(
            target_dir, folder_name, docx_path, on_limit=on_limit
        ):
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
                lambda: convert_docx_to_markdown(
                    docx_path, output_dir, create_subfolder=True,
                    output_name=output_name, on_limit=on_limit),
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
    return {"success": success_count, "skipped": skip_count, "failed": fail_count}

if __name__ == '__main__':
    import argparse

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    parser = argparse.ArgumentParser(description="批量将目录下的所有docx文档转换为Markdown")
    parser.add_argument("source_dir", nargs="?", default="1-Reference", help="源文件目录（默认 1-Reference）")
    parser.add_argument("output_dir", nargs="?", default="2-Temp", help="输出目录（默认 2-Temp）")
    parser.add_argument("--force", action="store_true", help="强制重新转换已存在的输出目录")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS,
                        help="单文档转换超时秒数（默认 300；<=0 不限制；Windows 无 SIGALRM 自动跳过）")
    parser.add_argument("--on-limit", choices=("reject", "skip"), default="reject",
                        help="资源超限处置：reject 整篇拒绝计失败（默认）；skip 仅跳过超限资源"
                             "继续转换（ZIP bomb 等恶意特征仍计失败）")
    args = parser.parse_args()

    summary = batch_convert(args.source_dir, args.output_dir, force=args.force,
                            timeout=args.timeout, on_limit=args.on_limit)
    # 有文档转换失败时以非零退出码结束，供 CI/自动化判定批次结果
    sys.exit(1 if summary["failed"] else 0)
