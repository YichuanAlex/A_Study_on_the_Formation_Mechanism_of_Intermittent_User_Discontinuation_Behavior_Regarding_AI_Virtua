#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""查看生成的访谈文档内容"""

import os
from docx import Document

# 查看第 10 份和第 30 份文档
for file_num in [10, 30]:
    file = f"受访者{file_num:03d}_访谈正文_详细版.docx"
    file_path = '访谈正文 Word_详细版/' + file
    
    print(f"\n{'='*60}")
    print(f"=== {file} 内容预览 ===")
    print(f"{'='*60}\n")
    
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text)
    
    # 显示核心问题部分
    print("【核心问题访谈部分】\n")
    for i, text in enumerate(full_text):
        if "访谈者：" in text and ("隐私" in text or "情感" in text or "不满意" in text or "信任" in text or "特殊群体" in text or "渠道" in text or "动机" in text or "社会" in text or "中辍" in text or "年龄" in text):
            print(f"问：{text}")
            if i+1 < len(full_text):
                print(f"答：{full_text[i+1]}\n")
    
    print(f"\n总段落数：{len(full_text)}")
