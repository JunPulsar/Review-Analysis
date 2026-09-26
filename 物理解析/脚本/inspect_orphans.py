#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""inspect_orphans.py <file.docx> — 打印该文件抽取的无题号条目。"""
import sys
from docx import Document
import merge_physics as MP

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

f = sys.argv[1]
entries = MP.extract_entries(f)
print(f'文件: {f}  条目 {len(entries)}')
k = 0
for e in entries:
    if not e['num'] and not e.get('src'):
        k += 1
        print(f'--- 孤立 #{k} (num={e["num"]!r}, src={e.get("src")!r}) ---')
        for ln in e['lines']:
            print('   ', ln[:120])
print(f'孤立条目合计: {k}')
