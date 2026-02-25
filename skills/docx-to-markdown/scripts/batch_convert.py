#!/usr/bin/env python3
"""
批量将目录下的所有docx文档转换为markdown格式
每个文档生成一个同名文件夹，包含md文件和assets子文件夹
"""

import logging
import os
import sys
import glob

# 支持从同目录或作为模块导入
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from convert_docx import convert_docx_to_markdown, sanitize_stem

logger = logging.getLogger(__name__)


def batch_convert(source_dir, output_dir, force=False):
    """批量转换目录下的所有docx文件

    Args:
        source_dir: 源文件目录
        output_dir: 输出目录
        force: 为 True 时强制重新转换已存在的输出目录
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
        
        # 检查是否已经处理过（--force 时跳过此检查）
        if not force and os.path.exists(target_dir):
            logger.info("  已存在，跳过（使用 --force 强制重新转换）")
            skip_count += 1
            continue

        try:
            if force and os.path.exists(target_dir):
                import shutil
                if os.path.isdir(target_dir) and not os.path.islink(target_dir):
                    shutil.rmtree(target_dir)
                else:
                    os.remove(target_dir)

            convert_docx_to_markdown(docx_path, output_dir, create_subfolder=True)
            logger.info("  完成")
            success_count += 1
        except Exception as e:
            logger.error("  失败: %s", e)
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
    args = parser.parse_args()

    batch_convert(args.source_dir, args.output_dir, force=args.force)
