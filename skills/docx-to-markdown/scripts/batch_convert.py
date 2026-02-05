#!/usr/bin/env python3
"""
批量将1-Reference目录下的所有docx文档转换为markdown格式
每个文档生成一个同名文件夹，包含md文件和assets子文件夹
"""

import os
import sys
import glob
from convert_docx import convert_docx_to_markdown

def batch_convert(source_dir, output_dir):
    """批量转换目录下的所有docx文件"""
    
    # 获取所有docx文件
    docx_files = glob.glob(os.path.join(source_dir, '*.docx'))
    
    if not docx_files:
        print(f"在 {source_dir} 中没有找到docx文件")
        return
    
    print(f"找到 {len(docx_files)} 个docx文件待处理\n")
    
    success_count = 0
    fail_count = 0
    
    for i, docx_path in enumerate(sorted(docx_files), 1):
        # 获取文件名（不含扩展名）作为输出文件夹名
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        # 移除文件名中的特殊引号，保持文件夹名称简洁
        folder_name = base_name.replace('"', '').replace('"', '').replace('"', '')
        target_dir = os.path.join(output_dir, folder_name)
        
        print(f"[{i}/{len(docx_files)}] 正在处理: {base_name}")
        
        # 检查是否已经处理过
        if os.path.exists(target_dir):
            print(f"  ⏭️  已存在，跳过")
            success_count += 1
            continue
        
        try:
            # 使用 create_subfolder=False，因为我们已经计算好了目标目录
            convert_docx_to_markdown(docx_path, output_dir, create_subfolder=True)
            print(f"  ✅ 完成\n")
            success_count += 1
        except Exception as e:
            print(f"  ❌ 失败: {e}\n")
            fail_count += 1
    
    print(f"\n{'='*50}")
    print(f"处理完成: 成功 {success_count} 个, 失败 {fail_count} 个")

if __name__ == '__main__':
    source_dir = '1-Reference'
    output_dir = '2-Temp'
    
    # 支持命令行参数覆盖默认值
    if len(sys.argv) >= 3:
        source_dir = sys.argv[1]
        output_dir = sys.argv[2]
    
    batch_convert(source_dir, output_dir)
