#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_questions.py — 检查「数字开头的行」哪些是题、哪些是知识小标题，
为避免把知识小标题当成题号（生物项目踩过的坑）。
输出 probe_questions.txt。
"""
import os
import re
import sys
from docx import Document

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

RE_NUM = re.compile(r'^(\d+)[.、．]\s*(.*)$')
RE_YEAR = re.compile(r'[（(]\s*\d{4}')
RE_EMPTY = re.compile(r'[（(]\s*[　 ]*[）)]')
RE_WORDS = re.compile(r'下列|正确|错误|判断|说法|叙述|符合|分析|计算|回答|说明|探究|验证|假设|设计|理由|原因|求|则|若|如图所示|如图|大小|方向|多少|可知|根据')

files = sorted(f for f in os.listdir('.') if f.endswith('.docx') and f != '教师目录.docx')
out = []
for f in files:
    doc = Document(f)
    for t in [p.text.strip() for p in doc.paragraphs if p.text.strip()]:
        m = RE_NUM.match(t)
        if not m:
            continue
        rest = m.group(2)
        if RE_YEAR.search(t) or RE_EMPTY.search(t):
            # 强信号：数字题
            continue
        out.append(f'{f}\t{t}')

with open('probe_questions.txt', 'w', encoding='utf-8') as fo:
    fo.write(f'无年份/无空括号 的“数字开头”行（需人工判定是题还是知识）: {len(out)}\n')
    fo.write('-' * 80 + '\n')
    for line in out:
        fo.write(line + '\n')
print(f'写出 probe_questions.txt: {len(out)} 行')
