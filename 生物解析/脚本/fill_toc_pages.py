#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fill_toc_pages.py — 把目录中的页码占位 00 替换为 WPS 实测页码（只读页码映射，写回目录 run 文本）"""
import os
import re
from docx import Document

ROOT = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(ROOT, '_toc_tmp')


def load_map(path):
    """返回 (页码表, 前置偏移)

    pagemap 文件里 WPS 给的是**物理页**；封面/说明/目录不编号、正文从 1 重新开始，
    所以显示页 = 物理页 - 偏移。偏移由 wps_pagemap 脚本以 '__OFFSET__|n' 写入。
    """
    m = {}
    offset = 0
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if '|' in line:
                k, v = line.rsplit('|', 1)
                k = k.strip().lstrip('\ufeff')
                if k == '__OFFSET__':
                    offset = int(v)
                    continue
                m[k] = int(v)
    return m, offset


def fill(docx_path, pagemap_path):
    page, offset = load_map(pagemap_path)
    doc = Document(docx_path)
    n = 0
    for p in doc.paragraphs:
        t = p.text
        m = re.match(r'^(\d{2,3})\. (.*?)\t00$', t)
        if not m:
            continue
        title = m.group(2)
        if title not in page:
            print(f'  未找到页码: {title}')
            continue
        pg = page[title] - offset          # 物理页 → 显示页
        if pg < 1:
            print(f'  页码异常(偏移 {offset}): {title} -> {page[title]}')
            continue
        # 把最后 run（占位 00）改掉
        runs = [r for r in p.runs]
        if not runs:
            continue
        # 最后 run 可能是 '00'，也可能是 '\t' + '00' 拆分成两个 run
        if runs[-1].text == '00':
            runs[-1].text = str(pg)
        else:
            # 兜底：拼一个新 run
            for r in runs:
                if r.text == '00':
                    r.text = str(pg)
                    break
        n += 1
    if n:
        doc.save(docx_path)   # 只有真回填了才写盘；空跑也重写会无谓改变文件指纹
    return n


targets = []
D = os.path.join(ROOT, '成册解析')
for fn, key in (('大本解析_全书.docx', 'book'),
                ('小本解析_合集.docx', 'xb'),
                ('大本小本合订本.docx', 'cb')):
    p = os.path.join(D, fn)
    if os.path.exists(p):
        targets.append((p, os.path.join(TMP, 'pagemap_%s.txt' % key)))

for docx_path, map_path in targets:
    if not os.path.exists(map_path):
        print('skip 无映射:', docx_path)
        continue
    try:
        n = fill(docx_path, map_path)
        print(f'{os.path.basename(docx_path)}: 回填 {n} 个目录页码')
    except Exception as e:
        print(f'FAIL {docx_path}: {e}')
