#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fill_toc_pages_physics.py — 用 WPS 实测页码回填三本目录的 00 占位。

  book = 物理解析_大本.docx
  xb   = 物理解析_小本.docx
  cb   = 物理解析_大本小本合订本.docx

页码来源：_toc_tmp/pagemap_<key>.txt（wps_pagemap_physics.ps1 实测产物，
由 export_titles_physics.py 生成的 titles_<key>.txt 定位）。
"""
import os
import re
import sys
from docx import Document

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

TMP = '_toc_tmp'
BOOKS = [
    ('book', '物理解析_大本.docx'),
    ('xb', '物理解析_小本.docx'),
    ('cb', '物理解析_大本小本合订本.docx'),
]


def load_map(path):
    m = {}
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if '|' in line:
                k, v = line.rsplit('|', 1)
                m[k.strip().lstrip('\ufeff')] = int(v)
    return m


def load_entries(path):
    out = []
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.rstrip('\n')
            if '\t' in line:
                disp, head = line.split('\t', 1)
                out.append((disp, head))
    return out


def fill(docx_path, map_path, ent_path):
    page = load_map(map_path)
    entries = load_entries(ent_path)
    doc = Document(docx_path)
    toc = [p for p in doc.paragraphs if re.match(r'^\d{2,3}\. .*\t\d+$', p.text)]
    if len(toc) != len(entries):
        print(f'  [警告] 目录段落 {len(toc)} vs 条目 {len(entries)}')

    n, missed = 0, []
    for p, (disp, head) in zip(toc, entries):
        pg = page.get(head)
        if not pg or pg <= 0:
            missed.append((disp, head))
            continue
        for r in reversed(p.runs):
            if r.text == '00':
                r.text = str(pg)
                n += 1
                break
        else:
            missed.append((disp, head))
    # 只有真的回填了才写盘；空跑也重写会无谓改变文件指纹
    if n:
        doc.save(docx_path)
    return n, len(entries), missed


def main():
    for key, fn in BOOKS:
        if not os.path.exists(fn):
            print(f'skip 缺文件: {fn}')
            continue
        mp = os.path.join(TMP, f'pagemap_{key}.txt')
        ep = os.path.join(TMP, f'toc_entries_{key}.txt')
        if not (os.path.exists(mp) and os.path.exists(ep)):
            print(f'skip 缺清单: {key}')
            continue
        n, total, missed = fill(fn, mp, ep)
        print(f'{fn}: 回填 {n}/{total}')
        for d, h in missed[:10]:
            print(f'    未匹配: {d}  ->  {h}')


if __name__ == '__main__':
    main()
