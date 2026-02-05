#!/usr/bin/env python3
"""
将docx文档转换为markdown格式，并提取所有图片到assets文件夹
支持将嵌入的Excel表格转换为Markdown表格
"""

import os
import sys
import zipfile
import shutil
from pathlib import Path
import mammoth
import re
import xml.etree.ElementTree as ET
import io
import pandas as pd
import warnings
warnings.filterwarnings('ignore')


def parse_relationships(docx_path):
    """解析docx中的关系文件，找出Excel嵌入和对应预览图的映射"""
    excel_to_preview = {}  # Excel文件 -> 预览图文件
    preview_to_excel = {}  # 预览图文件 -> Excel文件
    
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        try:
            rels_content = zip_ref.read('word/_rels/document.xml.rels')
            root = ET.fromstring(rels_content)
            
            # 收集所有关系
            relationships = {}
            for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rid = rel.get('Id')
                rel_type = rel.get('Type', '').split('/')[-1]
                target = rel.get('Target', '')
                relationships[rid] = {'type': rel_type, 'target': target}
            
            # 按 rId 数字排序，找出 Excel 和紧随其后的 image 的对应关系
            sorted_rids = sorted(relationships.keys(), key=lambda x: int(x.replace('rId', '')))
            
            for i, rid in enumerate(sorted_rids):
                rel = relationships[rid]
                if rel['type'] == 'package' and rel['target'].endswith('.xlsx'):
                    excel_file = rel['target']
                    # 查找下一个关系是否是 image
                    if i + 1 < len(sorted_rids):
                        next_rid = sorted_rids[i + 1]
                        next_rel = relationships[next_rid]
                        if next_rel['type'] == 'image':
                            preview_file = next_rel['target']
                            excel_to_preview[excel_file] = preview_file
                            preview_to_excel[preview_file] = excel_file
                            
        except Exception as e:
            print(f"  警告: 解析关系文件失败: {e}")
    
    return excel_to_preview, preview_to_excel


def excel_to_markdown(xlsx_data):
    """将Excel数据转换为Markdown表格"""
    try:
        df = pd.read_excel(io.BytesIO(xlsx_data))
        
        # 清理空行空列
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        if df.empty:
            return None
        
        # 填充NaN为空字符串
        df = df.fillna('')
        
        # 转换为Markdown表格
        markdown_table = df.to_markdown(index=False)
        return markdown_table
        
    except Exception as e:
        print(f"    Excel转Markdown失败: {e}")
        return None


def detect_image_format(image_data):
    """检测图片的真实格式"""
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
    else:
        return '.png'  # 默认


def extract_content_from_docx(docx_path, assets_dir):
    """从docx中提取图片和Excel数据"""
    images = {}  # 图片路径映射
    excel_tables = {}  # Excel预览图路径 -> Markdown表格内容
    
    # 解析关系，找出Excel和预览图的对应
    excel_to_preview, preview_to_excel = parse_relationships(docx_path)
    
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        # 先提取所有Excel文件的数据并转换为Markdown
        for file_info in zip_ref.filelist:
            if file_info.filename.startswith('word/embeddings/') and file_info.filename.endswith('.xlsx'):
                excel_file = file_info.filename.replace('word/', '')
                xlsx_data = zip_ref.read(file_info.filename)
                
                # 找到对应的预览图
                if excel_file in excel_to_preview:
                    preview_file = excel_to_preview[excel_file]
                    full_preview_path = f"word/{preview_file}"
                    
                    # 转换Excel为Markdown表格
                    markdown_table = excel_to_markdown(xlsx_data)
                    if markdown_table:
                        excel_tables[full_preview_path] = markdown_table
                        print(f"  转换Excel为表格: {excel_file}")
        
        # 处理图片
        for file_info in zip_ref.filelist:
            if file_info.filename.startswith('word/media/'):
                image_name = os.path.basename(file_info.filename)
                
                # 检查这个图片是否是Excel的预览图
                if file_info.filename in excel_tables:
                    # 这是Excel预览图，用Markdown表格替代
                    # 记录一个特殊标记，后续在转换时替换
                    images[file_info.filename] = ('TABLE', excel_tables[file_info.filename])
                    continue
                
                # 普通图片，直接提取
                image_data = zip_ref.read(file_info.filename)
                
                # 检测真实的图片格式并修正扩展名
                actual_ext = detect_image_format(image_data)
                base_name = os.path.splitext(image_name)[0]
                corrected_name = f"{base_name}{actual_ext}"
                
                image_path = os.path.join(assets_dir, corrected_name)
                with open(image_path, 'wb') as f:
                    f.write(image_data)
                
                images[file_info.filename] = ('IMAGE', f"assets/{corrected_name}")
                print(f"  提取图片: {corrected_name}")
    
    return images


def convert_docx_to_markdown(docx_path, output_dir, create_subfolder=True):
    """将docx转换为markdown
    
    Args:
        docx_path: DOCX 文件路径
        output_dir: 输出目录路径
        create_subfolder: 是否在输出目录下创建以文件名命名的子文件夹（默认 True）
    """
    
    # 获取文件名（不含扩展名）
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    # 移除文件名中的特殊引号
    folder_name = base_name.replace('"', '').replace('"', '').replace('"', '')
    
    # 确定最终输出目录
    if create_subfolder:
        final_output_dir = os.path.join(output_dir, folder_name)
    else:
        final_output_dir = output_dir
    
    # 创建输出目录
    os.makedirs(final_output_dir, exist_ok=True)
    assets_dir = os.path.join(final_output_dir, 'assets')
    os.makedirs(assets_dir, exist_ok=True)
    
    # 提取图片和Excel表格
    print(f"正在提取内容...")
    content_map = extract_content_from_docx(docx_path, assets_dir)
    
    # 使用mammoth转换为HTML
    print(f"正在转换文档...")
    
    # 内容计数器和列表
    content_counter = [0]
    content_list = list(content_map.values())
    
    def convert_image(image):
        """处理图片转换"""
        if content_counter[0] < len(content_list):
            content_type, content_value = content_list[content_counter[0]]
            content_counter[0] += 1
            if content_type == 'IMAGE':
                return {"src": content_value}
            else:
                # 表格，返回一个特殊的占位符
                # 使用一个不太可能出现在正常文本中的标记
                return {"src": f"__TABLE_PLACEHOLDER_{content_counter[0]-1}__"}
        return {}
    
    with open(docx_path, 'rb') as docx_file:
        result = mammoth.convert_to_html(
            docx_file,
            convert_image=mammoth.images.img_element(convert_image)
        )
        html = result.value
    
    # 将HTML转换为Markdown
    markdown = html_to_markdown(html)
    
    # 替换表格占位符
    for i, (content_type, content_value) in enumerate(content_list):
        if content_type == 'TABLE':
            placeholder = f"![](__TABLE_PLACEHOLDER_{i}__)"
            markdown = markdown.replace(placeholder, f"\n\n{content_value}\n\n")
    
    md_path = os.path.join(final_output_dir, f"{folder_name}.md")
    
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(markdown)
    
    print(f"转换完成: {md_path}")
    return md_path


def html_to_markdown(html):
    """将HTML转换为Markdown"""
    
    # 处理标题
    html = re.sub(r'<h1[^>]*>(.*?)</h1>', r'# \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h2[^>]*>(.*?)</h2>', r'## \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h3[^>]*>(.*?)</h3>', r'### \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h4[^>]*>(.*?)</h4>', r'#### \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h5[^>]*>(.*?)</h5>', r'##### \1\n\n', html, flags=re.DOTALL)
    html = re.sub(r'<h6[^>]*>(.*?)</h6>', r'###### \1\n\n', html, flags=re.DOTALL)
    
    # 处理粗体和斜体
    html = re.sub(r'<strong>(.*?)</strong>', r'**\1**', html, flags=re.DOTALL)
    html = re.sub(r'<b>(.*?)</b>', r'**\1**', html, flags=re.DOTALL)
    html = re.sub(r'<em>(.*?)</em>', r'*\1*', html, flags=re.DOTALL)
    html = re.sub(r'<i>(.*?)</i>', r'*\1*', html, flags=re.DOTALL)
    
    # 处理图片
    html = re.sub(r'<img[^>]*src="([^"]*)"[^>]*/?>',  r'![](\1)\n\n', html)
    
    # 处理链接
    html = re.sub(r'<a[^>]*href="([^"]*)"[^>]*>(.*?)</a>', r'[\2](\1)', html, flags=re.DOTALL)
    
    # 处理段落
    html = re.sub(r'<p[^>]*>(.*?)</p>', r'\1\n\n', html, flags=re.DOTALL)
    
    # 处理换行
    html = re.sub(r'<br\s*/?>', '\n', html)
    
    # 处理列表
    html = re.sub(r'<ul[^>]*>', '\n', html)
    html = re.sub(r'</ul>', '\n', html)
    html = re.sub(r'<ol[^>]*>', '\n', html)
    html = re.sub(r'</ol>', '\n', html)
    html = re.sub(r'<li[^>]*>(.*?)</li>', r'- \1\n', html, flags=re.DOTALL)
    
    # 处理表格 - 简化处理
    def convert_table(match):
        table_html = match.group(0)
        rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_html, re.DOTALL)
        if not rows:
            return ''
        
        markdown_table = []
        for i, row in enumerate(rows):
            cells = re.findall(r'<t[hd][^>]*>(.*?)</t[hd]>', row, re.DOTALL)
            # 清理单元格内容
            cells = [re.sub(r'<[^>]+>', '', cell).strip() for cell in cells]
            markdown_table.append('| ' + ' | '.join(cells) + ' |')
            
            # 在第一行后添加分隔符
            if i == 0:
                markdown_table.append('| ' + ' | '.join(['---'] * len(cells)) + ' |')
        
        return '\n'.join(markdown_table) + '\n\n'
    
    html = re.sub(r'<table[^>]*>.*?</table>', convert_table, html, flags=re.DOTALL)
    
    # 移除剩余的HTML标签
    html = re.sub(r'<[^>]+>', '', html)
    
    # 清理多余的空行
    html = re.sub(r'\n{3,}', '\n\n', html)
    
    # 清理HTML实体
    html = html.replace('&nbsp;', ' ')
    html = html.replace('&lt;', '<')
    html = html.replace('&gt;', '>')
    html = html.replace('&amp;', '&')
    html = html.replace('&quot;', '"')
    
    return html.strip()


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("用法: python convert_docx.py <docx文件路径> <输出目录>")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2]
    
    if not os.path.exists(docx_path):
        print(f"错误: 文件不存在 - {docx_path}")
        sys.exit(1)
    
    convert_docx_to_markdown(docx_path, output_dir)
