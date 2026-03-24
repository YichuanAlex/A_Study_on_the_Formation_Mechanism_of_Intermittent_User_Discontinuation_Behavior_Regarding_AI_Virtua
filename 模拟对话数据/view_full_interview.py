#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""查看完整的访谈文档内容"""

import os
from docx import Document

# 查看第 5 份文档的完整内容
file = "受访者 005_访谈正文_详细版.docx"
cwd = os.getcwd()
file_path = os.path.join(cwd, '访谈正文 Word_详细版', file)

print(f"完整文件：{file}\n")
print(f"文件路径：{file_path}")
print(f"文件存在：{os.path.exists(file_path)}\n")
print("=" * 70)

doc = Document(file_path)
full_text = []
for para in doc.paragraphs:
    if para.text.strip():
        full_text.append(para.text)

# 显示所有内容
for text in full_text:
    print(text)
    print()

print("=" * 70)
print(f"\n总段落数：{len(full_text)}")
