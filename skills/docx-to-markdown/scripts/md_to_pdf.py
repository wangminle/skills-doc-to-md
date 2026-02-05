#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import markdown
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# 注册中文字体（使用系统字体）
try:
    # 尝试注册常见的中文字体
    font_paths = [
        '/System/Library/Fonts/PingFang.ttc',  # macOS
        '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',  # Linux
        'C:/Windows/Fonts/msyh.ttc',  # Windows
    ]

    chinese_font = None
    for font_path in font_paths:
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
            chinese_font = 'ChineseFont'
            break

    if not chinese_font:
        chinese_font = 'Helvetica'  # 回退到默认字体

except Exception as e:
    print(f"字体注册警告: {e}")
    chinese_font = 'Helvetica'

# 读取markdown文件
input_file = '3-Results/2-分布式唤醒需求体系Review报告-v2.md'
output_file = '3-Results/2-分布式唤醒需求体系Review报告-v2.pdf'

with open(input_file, 'r', encoding='utf-8') as f:
    md_content = f.read()

# 转换markdown到HTML
html_content = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])

# 创建PDF
doc = SimpleDocTemplate(
    output_file,
    pagesize=A4,
    rightMargin=2*cm,
    leftMargin=2*cm,
    topMargin=2*cm,
    bottomMargin=2*cm
)

# 创建样式
styles = getSampleStyleSheet()

# 标题样式
title_style = ParagraphStyle(
    'CustomTitle',
    parent=styles['Heading1'],
    fontName=chinese_font,
    fontSize=18,
    textColor=colors.HexColor('#1a5490'),
    spaceAfter=20,
    alignment=TA_CENTER,
    leading=24
)

heading1_style = ParagraphStyle(
    'CustomH1',
    parent=styles['Heading1'],
    fontName=chinese_font,
    fontSize=14,
    textColor=colors.HexColor('#1a5490'),
    spaceAfter=12,
    spaceBefore=20,
    leading=20
)

heading2_style = ParagraphStyle(
    'CustomH2',
    parent=styles['Heading2'],
    fontName=chinese_font,
    fontSize=12,
    textColor=colors.HexColor('#2c3e50'),
    spaceAfter=10,
    spaceBefore=15,
    leading=18
)

normal_style = ParagraphStyle(
    'CustomNormal',
    parent=styles['Normal'],
    fontName=chinese_font,
    fontSize=10,
    leading=16,
    spaceAfter=8,
    alignment=TA_LEFT
)

bullet_style = ParagraphStyle(
    'CustomBullet',
    parent=styles['Normal'],
    fontName=chinese_font,
    fontSize=10,
    leading=16,
    spaceAfter=6,
    leftIndent=20,
    alignment=TA_LEFT
)

story = []

# 处理HTML内容并添加到story
import re
from html.parser import HTMLParser

class MarkdownToPDFParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.content = []
        self.in_heading = False
        self.heading_level = 0
        self.in_list = False
        self.in_paragraph = False
        self.current_text = []

    def handle_starttag(self, tag, attrs):
        if tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            self.in_heading = True
            self.heading_level = int(tag[1])
        elif tag == 'li':
            self.in_list = True
        elif tag == 'p':
            self.in_paragraph = True
            self.current_text = []

    def handle_endtag(self, tag):
        if tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            self.in_heading = False
            text = ''.join(self.current_text).strip()
            if text:
                if self.heading_level == 1:
                    self.content.append(('h1', text))
                elif self.heading_level == 2:
                    self.content.append(('h2', text))
                else:
                    self.content.append(('h3', text))
            self.current_text = []
        elif tag == 'li':
            self.in_list = False
            text = ''.join(self.current_text).strip()
            if text:
                self.content.append(('bullet', text))
            self.current_text = []
        elif tag == 'p':
            self.in_paragraph = False
            text = ''.join(self.current_text).strip()
            if text:
                self.content.append(('paragraph', text))
            self.current_text = []

    def handle_data(self, data):
        self.current_text.append(data)

parser = MarkdownToPDFParser()
parser.feed(html_content)

# 根据解析结果构建PDF
for item_type, text in parser.content:
    if item_type == 'h1':
        story.append(Paragraph(text, title_style))
        story.append(Spacer(1, 0.3*cm))
    elif item_type == 'h2':
        story.append(Paragraph(text, heading1_style))
        story.append(Spacer(1, 0.2*cm))
    elif item_type == 'h3':
        story.append(Paragraph(text, heading2_style))
        story.append(Spacer(1, 0.2*cm))
    elif item_type == 'bullet':
        story.append(Paragraph('• ' + text, bullet_style))
    elif item_type == 'paragraph':
        story.append(Paragraph(text, normal_style))
        story.append(Spacer(1, 0.2*cm))

# 构建PDF
try:
    doc.build(story)
    print(f'✅ PDF已成功生成: {output_file}')
    print(f'📄 文件大小: {os.path.getsize(output_file)/1024:.1f} KB')
except Exception as e:
    print(f'❌ 生成PDF时出错: {e}')
