#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""peek.py <file.docx> — 把文档全部非空段落写到 peek_<stem>.txt 供人工查看结构。"""
import os
import sys
from docx import Document

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

f = sys.argv[1]
doc = Document(f)
paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
out = f'peek_{os.path.splitext(os.path.basename(f))[0][:30]}.txt'
with open(out, 'w', encoding='utf-8') as fo:
    fo.write(f'文件: {f}\n总非空段落: {len(paras)}\n')
    fo.write('-' * 60 + '\n')
    for i, t in enumerate(paras):
        fo.write(f'[{i:03d}] {t}\n')
print(f'{out} 写好了 ({len(paras)} 段)')
