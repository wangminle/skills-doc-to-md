#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 Markdown 文件转换为 PDF，支持中文字体回退。

用法:
  python scripts/md_to_pdf.py <input.md> [output.pdf]
"""

import argparse
import os
import shutil
import subprocess
from html.parser import HTMLParser
from pathlib import Path
from typing import Optional


def _pick_chinese_font() -> str:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    # 1) 尝试系统字体（能注册就用；ttc 在部分 reportlab 版本可能不支持）
    candidates = [
        "/System/Library/Fonts/PingFang.ttc",  # macOS
        "/System/Library/Fonts/STHeiti Light.ttc",  # macOS
        "/System/Library/Fonts/Supplemental/Songti.ttc",  # macOS
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",  # Linux
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",  # Linux
        "C:/Windows/Fonts/msyh.ttc",  # Windows
        "C:/Windows/Fonts/simsun.ttc",  # Windows
    ]
    for path in candidates:
        if not os.path.exists(path):
            continue
        try:
            pdfmetrics.registerFont(TTFont("ChineseFont", path))
            return "ChineseFont"
        except Exception:
            continue

    # 2) ReportLab 内置 CID 字体（无需本地字体文件）
    try:
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont

        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        return "STSong-Light"
    except Exception:
        return "Helvetica"


class MarkdownToPDFParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.content = []
        self.heading_level = 0
        self.current_text = []
        self.in_table = False
        self.table_rows = []
        self.current_row = []
        self.in_cell = False
        self.cell_text = []
        self.in_pre = False
        self.pre_text = []

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.in_table = True
            self.table_rows = []
            self.current_row = []
            return
        if self.in_table and tag == "tr":
            self.current_row = []
            return
        if self.in_table and tag in {"td", "th"}:
            self.in_cell = True
            self.cell_text = []
            return
        if tag == "pre":
            self.in_pre = True
            self.pre_text = []
            return

        if tag in {"h1", "h2", "h3", "h4", "h5", "h6", "li", "p"}:
            self.current_text = []
        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            self.heading_level = int(tag[1])

    def handle_endtag(self, tag):
        if self.in_table and tag in {"td", "th"}:
            text = "".join(self.cell_text).strip()
            self.current_row.append(text)
            self.in_cell = False
            self.cell_text = []
            return
        if self.in_table and tag == "tr":
            if self.current_row:
                self.table_rows.append(self.current_row)
            self.current_row = []
            return
        if tag == "table" and self.in_table:
            lines = []
            for row in self.table_rows:
                lines.append("| " + " | ".join(c.strip() for c in row) + " |")
            table_text = "\n".join(lines).strip()
            if table_text:
                self.content.append(("code", table_text))
            self.in_table = False
            self.table_rows = []
            return

        if tag == "pre" and self.in_pre:
            text = "".join(self.pre_text).rstrip()
            if text:
                self.content.append(("code", text))
            self.in_pre = False
            self.pre_text = []
            return

        # 仅处理块级标签；内联标签（strong/em/a/span 等）直接跳过，
        # 让文本继续在 current_text 中累积。
        if tag not in {"h1", "h2", "h3", "h4", "h5", "h6", "li", "p"}:
            return

        text = "".join(self.current_text).strip()
        self.current_text = []
        if not text:
            return

        if tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            self.content.append((f"h{self.heading_level}", text))
            self.heading_level = 0
        elif tag == "li":
            self.content.append(("bullet", text))
        elif tag == "p":
            self.content.append(("paragraph", text))

    def handle_data(self, data):
        if self.in_cell:
            self.cell_text.append(data)
        elif self.in_pre:
            self.pre_text.append(data)
        else:
            self.current_text.append(data)


def convert_md_to_pdf(input_file: str, output_file: Optional[str] = None) -> str:
    import markdown
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import Paragraph, Preformatted, SimpleDocTemplate, Spacer

    chinese_font = _pick_chinese_font()

    with open(input_file, "r", encoding="utf-8") as f:
        md_content = f.read()

    html_content = markdown.markdown(md_content, extensions=["tables", "fenced_code"])

    out_path = Path(output_file) if output_file else Path(input_file).with_suffix(".pdf")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Heading1"],
        fontName=chinese_font,
        fontSize=18,
        textColor=colors.HexColor("#1a5490"),
        spaceAfter=20,
        alignment=TA_CENTER,
        leading=24,
    )
    heading1_style = ParagraphStyle(
        "CustomH1",
        parent=styles["Heading1"],
        fontName=chinese_font,
        fontSize=14,
        textColor=colors.HexColor("#1a5490"),
        spaceAfter=12,
        spaceBefore=20,
        leading=20,
    )
    heading2_style = ParagraphStyle(
        "CustomH2",
        parent=styles["Heading2"],
        fontName=chinese_font,
        fontSize=12,
        textColor=colors.HexColor("#2c3e50"),
        spaceAfter=10,
        spaceBefore=15,
        leading=18,
    )
    heading3_style = ParagraphStyle(
        "CustomH3",
        parent=styles["Heading3"],
        fontName=chinese_font,
        fontSize=11,
        textColor=colors.HexColor("#34495e"),
        spaceAfter=8,
        spaceBefore=12,
        leading=16,
    )
    heading4_style = ParagraphStyle(
        "CustomH4",
        parent=styles["Normal"],
        fontName=chinese_font,
        fontSize=10,
        textColor=colors.HexColor("#34495e"),
        spaceAfter=6,
        spaceBefore=10,
        leading=15,
    )
    normal_style = ParagraphStyle(
        "CustomNormal",
        parent=styles["Normal"],
        fontName=chinese_font,
        fontSize=10,
        leading=16,
        spaceAfter=8,
        alignment=TA_LEFT,
    )
    bullet_style = ParagraphStyle(
        "CustomBullet",
        parent=styles["Normal"],
        fontName=chinese_font,
        fontSize=10,
        leading=16,
        spaceAfter=6,
        leftIndent=20,
        alignment=TA_LEFT,
    )
    code_style = ParagraphStyle(
        "CustomCode",
        parent=styles.get("Code", styles["Normal"]),
        fontName=chinese_font,
        fontSize=9,
        leading=12,
        backColor=colors.whitesmoke,
        leftIndent=10,
        rightIndent=10,
        spaceAfter=8,
    )

    parser = MarkdownToPDFParser()
    parser.feed(html_content)

    heading_style_map = {
        "h1": title_style,
        "h2": heading1_style,
        "h3": heading2_style,
        "h4": heading3_style,
        "h5": heading4_style,
        "h6": heading4_style,
    }

    story = []
    for item_type, text in parser.content:
        if item_type in heading_style_map:
            story.append(Paragraph(text, heading_style_map[item_type]))
            story.append(Spacer(1, 0.3 * cm if item_type == "h1" else 0.2 * cm))
        elif item_type == "bullet":
            story.append(Paragraph("• " + text, bullet_style))
        elif item_type == "paragraph":
            story.append(Paragraph(text, normal_style))
            story.append(Spacer(1, 0.2 * cm))
        elif item_type == "code":
            story.append(Preformatted(text, code_style))
            story.append(Spacer(1, 0.2 * cm))

    doc.build(story)
    return str(out_path)


def convert_md_to_pdf_with_pandoc(input_file: str, output_file: str):
    cmd = ["pandoc", input_file, "-o", output_file]
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        stderr = (e.stderr or "").strip()
        raise RuntimeError(f"pandoc 转换失败: {stderr or e}") from e


def run_conversion(input_file: str, output_file: str, engine: str) -> str:
    # 独立脚本模式：优先使用系统 pandoc（无 Python 三方依赖）
    if engine in {"auto", "pandoc"}:
        if shutil.which("pandoc"):
            convert_md_to_pdf_with_pandoc(input_file, output_file)
            return "pandoc"
        if engine == "pandoc":
            raise RuntimeError("未找到 pandoc，请先安装 pandoc，或改用 --engine python。")

    # Python 回退模式：仅当用户环境已有 markdown/reportlab 时可用
    if engine in {"auto", "python"}:
        try:
            convert_md_to_pdf(input_file, output_file)
            return "python"
        except ModuleNotFoundError as e:
            missing = getattr(e, "name", "依赖")
            raise RuntimeError(
                f"缺少 Python 依赖: {missing}。可执行 `pip install markdown reportlab`，"
                "或安装 pandoc 后使用默认模式。"
            ) from e

    raise RuntimeError(f"不支持的转换引擎: {engine}")


def main():
    parser = argparse.ArgumentParser(description="将Markdown转换为PDF（支持中文字体回退）")
    parser.add_argument("input", help="输入Markdown文件路径")
    parser.add_argument("output", nargs="?", help="输出PDF路径（默认同名.pdf）")
    parser.add_argument(
        "--engine",
        choices=["auto", "pandoc", "python"],
        default="auto",
        help="转换引擎：auto(默认，优先pandoc)、pandoc、python",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"错误: 文件不存在 - {input_path}")

    output_path = Path(args.output) if args.output else input_path.with_suffix(".pdf")

    try:
        used_engine = run_conversion(str(input_path), str(output_path), args.engine)
        print(f"✅ PDF已成功生成: {output_path}")
        print(f"🔧 使用引擎: {used_engine}")
        print(f"📄 文件大小: {os.path.getsize(output_path)/1024:.1f} KB")
    except Exception as e:
        raise SystemExit(f"❌ 生成PDF时出错: {e}")


if __name__ == "__main__":
    main()
