#!/usr/bin/env python3
"""
将docx文档转换为markdown格式，并提取所有图片到assets文件夹
支持将嵌入的Excel表格转换为Markdown表格
"""

import hashlib
import os
import sys
import zipfile
import re
import xml.etree.ElementTree as ET
import io
import unicodedata
from collections import defaultdict
import posixpath


_FORBIDDEN_FILENAME_CHARS_RE = re.compile(r'[\\/:*?"<>|]')
_WHITESPACE_RE = re.compile(r"\s+")
_QUOTE_CHARS = '"“”‘’‚‛„‟«»‹›'


def sanitize_stem(stem: str) -> str:
    raw = stem  # 保留原始值用于 hash
    stem = unicodedata.normalize("NFKC", stem or "")
    for ch in _QUOTE_CHARS:
        stem = stem.replace(ch, "")
    stem = _FORBIDDEN_FILENAME_CHARS_RE.sub("_", stem)
    stem = _WHITESPACE_RE.sub(" ", stem).strip()
    stem = stem.strip(". ").strip()
    if not stem:
        return "document"
    if len(stem) <= 120:
        return stem
    # 截断时附加原始全名的短 hash，避免不同长文件名映射到同一输出目录
    suffix = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:8]
    return f"{stem[:111]}_{suffix}"


def resolve_part_path(target: str) -> str:
    """将 relationship target 解析为 docx zip 内的规范路径（如 word/media/image1.png）"""
    target = (target or "").replace("\\", "/").strip()
    if not target:
        return ""
    if target.startswith("/"):
        target = target[1:]
    if target.startswith("word/"):
        return posixpath.normpath(target)
    return posixpath.normpath(posixpath.join("word", target))


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
            rels_root = ET.fromstring(rels_content)
            for rel in rels_root.findall(f'.//{{{NS_REL}}}Relationship'):
                rid = rel.get('Id')
                rel_type = rel.get('Type', '').split('/')[-1]
                target = rel.get('Target', '')
                relationships[rid] = {'type': rel_type, 'target': target}
        except Exception as e:
            print(f"  警告: 解析关系文件失败: {e}")
            return excel_to_preview, preview_to_excel, ordered_pairs

        # --- 方法1：从 document.xml 解析 OLE 对象的真实引用 ---
        NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        NS_V = "urn:schemas-microsoft-com:vml"
        NS_O = "urn:schemas-microsoft-com:office:office"
        NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        try:
            doc_xml = zip_ref.read('word/document.xml')
            doc_root = ET.fromstring(doc_xml)

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


def excel_to_markdown(xlsx_data):
    """将Excel数据转换为Markdown表格（仅依赖 openpyxl，无需 pandas）"""
    try:
        import openpyxl

        wb = openpyxl.load_workbook(io.BytesIO(xlsx_data), read_only=True, data_only=True)
        ws = wb.active
        if ws is None:
            return None

        # 读取所有行
        raw_rows = []
        for row in ws.iter_rows(values_only=True):
            raw_rows.append([str(cell) if cell is not None else '' for cell in row])
        wb.close()

        if not raw_rows:
            return None

        # 去掉全空行
        rows = [r for r in raw_rows if any(c.strip() for c in r)]
        if not rows:
            return None

        # 去掉全空列
        col_count = max(len(r) for r in rows)
        # 补齐短行
        rows = [r + [''] * (col_count - len(r)) for r in rows]
        non_empty_cols = [j for j in range(col_count) if any(rows[i][j].strip() for i in range(len(rows)))]
        if not non_empty_cols:
            return None
        rows = [[r[j] for j in non_empty_cols] for r in rows]

        # 构建 Markdown 表格
        header = '| ' + ' | '.join(rows[0]) + ' |'
        separator = '| ' + ' | '.join(['---'] * len(rows[0])) + ' |'
        body_lines = ['| ' + ' | '.join(r) + ' |' for r in rows[1:]]
        return header + '\n' + separator + '\n' + '\n'.join(body_lines)

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
    """从docx中提取图片和Excel数据，并构建“内容hash -> 内容”的映射

    返回:
        image_by_hash: { sha256_hex: "assets/xxx.png" }
        table_queue_by_hash: { sha256_hex: ["<md_table1>", "<md_table2>", ...] }
        table_repeat_by_hash: { sha256_hex: "<md_table>" }  # 队列耗尽时的稳定兜底
    """
    image_by_hash = {}
    table_queue_by_hash = defaultdict(list)
    table_repeat_by_hash = {}
    
    # 解析关系，找出Excel和预览图的对应
    excel_to_preview, preview_to_excel, ordered_pairs = parse_relationships(docx_path)
    
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        excel_md_by_path = {}
        table_preview_paths = set()

        # 先提取所有 Excel 文件的数据并转换为 Markdown
        for file_info in zip_ref.filelist:
            if file_info.filename.startswith('word/embeddings/') and file_info.filename.lower().endswith('.xlsx'):
                excel_file = file_info.filename
                xlsx_data = zip_ref.read(file_info.filename)

                markdown_table = excel_to_markdown(xlsx_data)
                if markdown_table:
                    excel_md_by_path[excel_file] = markdown_table

        # 建立预览图 hash -> 表格队列（同一预览图内容可对应多个表格）
        pairs = ordered_pairs if ordered_pairs else [(e, p) for e, p in excel_to_preview.items()]
        for excel_path, preview_path in pairs:
            table_md = excel_md_by_path.get(excel_path)
            if not table_md:
                continue
            if preview_path not in zip_ref.namelist():
                continue
            preview_data = zip_ref.read(preview_path)
            digest = hashlib.sha256(preview_data).hexdigest()
            table_queue_by_hash[digest].append(table_md)
            table_repeat_by_hash[digest] = table_md
            table_preview_paths.add(preview_path)
            print(f"  转换Excel为表格: {excel_path}")

        # 处理图片
        for file_info in zip_ref.filelist:
            if file_info.filename.startswith('word/media/'):
                image_name = os.path.basename(file_info.filename)
                
                # 检查这个图片是否是Excel的预览图
                if file_info.filename in table_preview_paths:
                    continue
                
                # 普通图片，直接提取
                image_data = zip_ref.read(file_info.filename)
                digest = hashlib.sha256(image_data).hexdigest()
                
                # 检测真实的图片格式并修正扩展名
                actual_ext = detect_image_format(image_data)
                base_name = os.path.splitext(image_name)[0]
                corrected_name = f"{base_name}{actual_ext}"
                
                image_path = os.path.join(assets_dir, corrected_name)
                # 扩展名修正后可能与已有文件同名，若内容不同则附加hash后缀避免覆盖
                if os.path.exists(image_path):
                    try:
                        with open(image_path, "rb") as f:
                            existing = f.read()
                        if existing != image_data:
                            corrected_name = f"{base_name}_{digest[:8]}{actual_ext}"
                            image_path = os.path.join(assets_dir, corrected_name)
                    except Exception:
                        corrected_name = f"{base_name}_{digest[:8]}{actual_ext}"
                        image_path = os.path.join(assets_dir, corrected_name)

                if not os.path.exists(image_path):
                    with open(image_path, 'wb') as f:
                        f.write(image_data)

                image_by_hash.setdefault(digest, f"assets/{corrected_name}")
                print(f"  提取图片: {corrected_name}")
    
    return image_by_hash, table_queue_by_hash, table_repeat_by_hash


def convert_docx_to_markdown(docx_path, output_dir, create_subfolder=True):
    """将docx转换为markdown
    
    Args:
        docx_path: DOCX 文件路径
        output_dir: 输出目录路径
        create_subfolder: 是否在输出目录下创建以文件名命名的子文件夹（默认 True）
    """
    
    # 获取文件名（不含扩展名）
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    folder_name = sanitize_stem(base_name)
    
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
    image_by_hash, table_queue_by_hash, table_repeat_by_hash = extract_content_from_docx(docx_path, assets_dir)
    table_md_by_placeholder = {}
    table_seq = [0]
    
    # 使用mammoth转换为HTML
    print(f"正在转换文档...")

    def convert_image(image):
        """根据图片内容hash，返回对应的assets路径或表格占位符"""
        with image.open() as image_bytes:
            image_data = image_bytes.read()
        digest = hashlib.sha256(image_data).hexdigest()

        table_queue = table_queue_by_hash.get(digest)
        if table_queue:
            # 若仅剩一个元素则不再弹出，确保同一预览图多次出现时仍稳定替换为表格
            table_md = table_queue[0] if len(table_queue) == 1 else table_queue.pop(0)
            placeholder = f"__TABLE_PLACEHOLDER_{digest}_{table_seq[0]}__"
            table_seq[0] += 1
            table_md_by_placeholder[placeholder] = table_md
            return {"src": placeholder}

        # 队列被消耗完时，继续复用最后一次已知表格，避免退化为普通图片
        if digest in table_repeat_by_hash:
            table_md = table_repeat_by_hash[digest]
            placeholder = f"__TABLE_PLACEHOLDER_{digest}_{table_seq[0]}__"
            table_seq[0] += 1
            table_md_by_placeholder[placeholder] = table_md
            return {"src": placeholder}

        image_src = image_by_hash.get(digest)
        if image_src:
            return {"src": image_src}

        # 兜底：某些情况下zip里的图片与mammoth回调数据不一致，直接按hash写入assets
        ext = detect_image_format(image_data)
        filename = f"image_{digest[:16]}{ext}"
        image_path = os.path.join(assets_dir, filename)
        if not os.path.exists(image_path):
            with open(image_path, "wb") as f:
                f.write(image_data)
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
            print(f"  mammoth提示: {msg}")
    
    # 将HTML转换为Markdown
    markdown = html_to_markdown(html)
    
    # 替换表格占位符
    for placeholder_key, table_md in table_md_by_placeholder.items():
        placeholder = f"![]({placeholder_key})"
        markdown = markdown.replace(placeholder, f"\n\n{table_md}\n\n")
    
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
        print("用法: python scripts/convert_docx.py <docx文件路径> <输出目录>  (在skill目录执行)")
        print("或:   python convert_docx.py <docx文件路径> <输出目录>          (在scripts目录执行)")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2]
    
    if not os.path.exists(docx_path):
        print(f"错误: 文件不存在 - {docx_path}")
        sys.exit(1)
    
    convert_docx_to_markdown(docx_path, output_dir)
