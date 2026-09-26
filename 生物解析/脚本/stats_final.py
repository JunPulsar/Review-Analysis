# -*- coding: utf-8 -*-
"""stats_final.py — 最终条目构成统计（只读）"""
import os
import re
from docx import Document

import merge_jiexi as MJ

ROOT = os.path.dirname(os.path.abspath(__file__))

total = 0
tf = 0
example = 0
choice = 0
big = 0
no_num = 0
lines = 0
src_count = 0

BIG_FEAT = re.compile(r'^\s*答案\s*[　 ]*(（\d+）|\(\d+\)|①|②|实验|遗传图解|设计|方案|思路|步骤|现象|结果|基因型为|请|简述|写出|分析|计算|比例|概率|原因|理由)')
SRC_FEAT = re.compile(r'\d{4}·')


for name in sorted(os.listdir(ROOT)):
    d = os.path.join(ROOT, name)
    if not os.path.isdir(d) or not name.startswith('大概念'):
        continue
    for f in sorted(os.listdir(d)):
        if not f.endswith('.docx') or f.startswith('~$') or f.endswith('_解析.docx'):
            continue
        entries = MJ.process_one_docx(os.path.join(d, f))
        for e in entries:
            total += 1
            lines += len(e['lines'])
            if SRF_ECC := MJ.RE_SRC.search(e['lines'][0] if False else ''):
                pass
            first = e['lines'][0]
            num = e.get('num', '')
            src = e.get('src', '')
            if src:
                src_count += 1
            if re.match(r'^[（(]\s*\d+\s*[）)]', first) and re.search(r'[×√]', first):
                tf += 1
            elif num.startswith('['):
                example += 1
            elif num == '':
                no_num += 1
            elif BIG_FEAT.match(first):
                big += 1
            else:
                choice += 1

print(f'条目总数: {total}  统计行数(lines合计): {lines}')
print(f'判断题条目: {tf}')
print(f'[例N]/[练N] 例题: {example}')
print(f'编号选择题(num=N. 且答案行无大题特征): {choice}')
print(f'大题/填空/实验题(num=N. 且答案含(1)/①/实验思路等): {big}')
print(f'无编号条目(判断块外的孤立答案/无编号真题): {no_num}')
print(f'带年份来源(src)的条目: {src_count}')
