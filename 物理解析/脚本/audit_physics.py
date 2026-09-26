#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_physics.py — 全量审计：抽出来的 答案/解析/提示 行是否与原文 1:1 全覆盖（不漏、不多）。
输出 audit_physics.txt（UTF-8）。
"""
import os
import sys
from docx import Document

import merge_physics as MP

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

files = sorted(f for f in os.listdir('.')
               if f.endswith('.docx') and f != '教师目录.docx'
               and not f.startswith('物理解析') and not f.endswith('_新.docx'))
rows = []
tot_raw_ans = tot_raw_ana = tot_get_ans = tot_get_ana = 0
tot_entries = 0
tot_orphans = 0

for f in files:
    doc = Document(f)
    raw = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    raw_ans = sum(1 for t in raw if MP.RE_ANSWER.match(t))
    raw_ana = sum(1 for t in raw if MP.RE_ANALYSIS.match(t))
    raw_prm = sum(1 for t in raw if MP.RE_PROMPT.match(t))
    entries = MP.extract_entries(f)
    lines = [ln for e in entries for ln in e['lines']]
    got_ans = sum(1 for ln in lines if MP.RE_ANSWER.match(ln))
    got_ana = sum(1 for ln in lines if MP.RE_ANALYSIS.match(ln))
    got_prm = sum(1 for ln in lines if MP.RE_PROMPT.match(ln))
    orphans = sum(1 for e in entries if not e['num'] and not e.get('src'))
    diff = (raw_ans - got_ans) + (raw_ana - got_ana) + (raw_prm - got_prm)
    rows.append((f, raw_ans, got_ans, raw_ana, got_ana, raw_prm, got_prm, len(entries), orphans, diff))
    tot_raw_ans += raw_ans
    tot_raw_ana += raw_ana
    tot_get_ans += got_ans
    tot_get_ana += got_ana
    tot_entries += len(entries)
    tot_orphans += orphans

with open('audit_physics.txt', 'w', encoding='utf-8') as fo:
    fo.write('文件 | 原答案 | 抽答案 | 原解析 | 抽解析 | 原提示 | 抽提示 | 条目 | 孤立条目 | 差值\n')
    fo.write('-' * 100 + '\n')
    for r in rows:
        fo.write(' | '.join(str(x) for x in r) + '\n')
    fo.write('\n=== 汇总 ===\n')
    fo.write(f'原始: 答案 {tot_raw_ans}，解析 {tot_raw_ana}；抽取: 答案 {tot_get_ans}，解析 {tot_get_ana}\n')
    fo.write(f'总条目 {tot_entries}，孤立条目（无题号/来源） {tot_orphans}\n')
    fo.write(f'全量差值（原文-抽取）：答案 {tot_raw_ans - tot_get_ans}，解析 {tot_raw_ana - tot_get_ana}\n')
    bad = [r for r in rows if r[9] != 0]
    fo.write(f'差异文件数: {len(bad)}\n')
    for r in bad:
        fo.write(f'  ! {r[0]}  原(答案{ r[1]} 解析{r[3]}) 抽(答案{r[2]} 解析{r[4]}) 差{r[9]}\n')
print(f'审计完成 -> audit_physics.txt；差值：答案 {tot_raw_ans - tot_get_ans}，解析 {tot_raw_ana - tot_get_ana}，孤立条目 {tot_orphans}')
