#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_output.py — 回读生成的小册子，打印目录与正文前 160 段，并统计结构。"""
import sys
from docx import Document

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

doc = Document('物理解析合集_小册子.docx')
paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
with open('verify_output.txt', 'w', encoding='utf-8') as fo:
    fo.write(f'总非空段落: {len(paras)}\n')
    fo.write(f'表格: {len(doc.tables)}\n')
    fo.write('-' * 70 + '\n【前 200 段】\n')
    for i, t in enumerate(paras[:200]):
        fo.write(f'[{i:03d}] {t[:150]}\n')
    fo.write('-' * 70 + '\n【关键词计数】\n')
    fo.write(f'含"目　录": {sum(1 for t in paras if t.startswith("目　录"))}\n')
    tab00 = '\t00'
    fo.write(f'含"00"(目录占位): {sum(1 for t in paras if tab00 in t)}\n')
    fo.write(f'含"（与"同源提示": {sum(1 for t in paras if "同源" in t)}\n')
    fo.write(f'含"无条目": {sum(1 for t in paras if "无条目" in t)}\n')
print('已写出 verify_output.txt')
