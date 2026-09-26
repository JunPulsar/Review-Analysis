#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_titles.py — 从各 docx 提取正文章节标题（13pt 加粗），生成 titles 文件供 WPS 定位页码

覆盖目标（全部位于「成册解析」文件夹）：
  成册解析/大本解析_全书.docx     → book
  成册解析/小本解析_合集.docx     → xb
  成册解析/大本小本合订本.docx    → cb
"""
import os
import re
from docx import Document

ROOT = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(ROOT, '_toc_tmp')
os.makedirs(TMP, exist_ok=True)
D = os.path.join(ROOT, '成册解析')

targets = []
for fn, key in (('大本解析_全书.docx', 'book'),
                ('小本解析_合集.docx', 'xb'),
                ('大本小本合订本.docx', 'cb')):
    p = os.path.join(D, fn)
    if os.path.exists(p):
        targets.append((p, key))

for path, key in targets:
    if not os.path.exists(path):
        print('!! 不存在', path)
        continue
    doc = Document(path)
    titles = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        for r in p.runs:
            if r.bold and r.font.size and abs(r.font.size.pt - 13) < 0.01:
                if t not in titles:
                    titles.append(t)
                break
    with open(os.path.join(TMP, 'titles_%s.txt' % key), 'w', encoding='utf-8') as f:
        f.write('\n'.join(titles) + '\n')
    print('%s: %d 个标题' % (key, len(titles)))
