#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检查生成的 Word 文档内容"""

from docx import Document
import os

# 使用相对路径
rel_path = "访谈正文 Word/受访者 001_访谈正文.docx"
abs_path = os.path.join(os.getcwd(), rel_path)
print(f"Current dir: {os.getcwd()}")
print(f"Checking file: {abs_path}")
print(f"File exists: {os.path.exists(abs_path)}")

# 列出目录内容
print(f"\nFiles in 访谈正文 Word: {len(os.listdir('访谈正文 Word'))}")

doc = Document(rel_path)
print("\n=== 受访者 001_访谈正文.docx 内容预览 ===\n")
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
for i, text in enumerate(paragraphs[:15]):
    print(f"{text}\n")
