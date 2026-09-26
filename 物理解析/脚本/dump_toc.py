#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 教师目录.docx 的段落流导出到 toc_dump.txt，用于确定合并顺序。"""
import sys
from docx import Document

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

doc = Document('教师目录.docx')
paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
with open('toc_dump.txt', 'w', encoding='utf-8') as f:
    f.write(f'总段落: {len(paras)}\n')
    for i, t in enumerate(paras):
        f.write(f'[{i:03d}] {t}\n')
print(f'总段落: {len(paras)} -> toc_dump.txt')
