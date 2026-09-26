#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scan_physics.py — 全量扫描 95 个 docx 的"解析可挖性"：
统计每个文件的答案/解析/提示/题目起始行数量，找异常（0 答案、0 解析、无题号等）。
输出 scan_physics.txt（UTF-8）。
"""
import os
import re
import sys
from docx import Document

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

OUT = 'scan_physics.txt'

RE_ANS = re.compile(r'^答案\s*[　 ]')
RE_ANA = re.compile(r'^解析\s*[　 ]')
RE_PRM = re.compile(r'^提示\s*[　 ]')
RE_Q_EX = re.compile(r'^例\d+[　 ]')
RE_Q_NUM = re.compile(r'^\d+\.[　 ]?')
RE_Q_THK = re.compile(r'^(思考|拓展|方法|技巧|反思)\s*[　 ]')
RE_Q_VAR = re.compile(r'^\[变式\]\s*')
RE_SRC = re.compile(r'[（(]\s*\d{4}')
RE_BARE = re.compile(r'^(答案|解析|提示)[：:]?$')

rows = []
files = sorted(f for f in os.listdir('.') if f.endswith('.docx') and f != '教师目录.docx')
for f in files:
    try:
        doc = Document(f)
    except Exception as e:
        rows.append((f, 'ERR', str(e)))
        continue
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    n_ans = sum(1 for t in paras if RE_ANS.match(t))
    n_ana = sum(1 for t in paras if RE_ANA.match(t))
    n_prm = sum(1 for t in paras if RE_PRM.match(t))
    n_qex = sum(1 for t in paras if RE_Q_EX.match(t))
    n_qnum = sum(1 for t in paras if RE_Q_NUM.match(t))
    n_qthk = sum(1 for t in paras if RE_Q_THK.match(t))
    n_qvar = sum(1 for t in paras if RE_Q_VAR.match(t))
    n_src = sum(1 for t in paras if RE_SRC.search(t))
    n_bare = sum(1 for t in paras if RE_BARE.match(t))
    n_tbl = len(doc.tables)
    n_img = doc.element.body.xml.count('r:embed')
    rows.append((f, len(paras), n_tbl, n_img, n_ans, n_ana, n_prm, n_qex, n_qnum, n_qthk, n_qvar, n_src, n_bare))

with open(OUT, 'w', encoding='utf-8') as fo:
    fo.write('文件 | 段落 | 表 | 图 | 答案 | 解析 | 提示 | 例N | 序号题 | 思考/拓展等 | [变式] | 含年份 | 裸标签\n')
    fo.write('-' * 120 + '\n')
    for r in rows:
        fo.write(' | '.join(str(x) for x in r) + '\n')
    fo.write('\n--- 汇总 ---\n')
    non_err = [r for r in rows if not (len(r) > 1 and r[1] == 'ERR')]
    idx = {k: i for i, k in enumerate(['段落', '表', '图', '答案', '解析', '提示', '例N', '序号题', '思考/拓展等', '[变式]', '含年份', '裸标签'])}
    for k, i in idx.items():
        vals = [r[i + 1] for r in non_err]
        fo.write(f'{k}: 合计 {sum(vals)}，文件数 {sum(1 for v in vals if v > 0)}，最大值 {max(vals) if vals else 0}\n')
    fo.write('\n--- 异常文件 ---\n')
    for r in non_err:
        if r[4] == 0 and r[5] == 0:
            fo.write(f'无答案也无解析: {r[0]}\n')
        elif r[4] == 0:
            fo.write(f'无答案: {r[0]}\n')
        elif r[5] == 0:
            fo.write(f'无解析: {r[0]}\n')
print(f'完成 -> {OUT}（文件数 {len(files)}）')
