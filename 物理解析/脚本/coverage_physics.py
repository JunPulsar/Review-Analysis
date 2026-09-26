#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""coverage_physics.py — 条目级解析覆盖率统计。"""
import os
import sys
import merge_physics as MP

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

files = sorted(f for f in os.listdir('.')
               if f.endswith('.docx') and f != '教师目录.docx'
               and not f.startswith('物理解析') and not f.endswith('_新.docx'))

total_entries = 0
with_ans = 0
with_ana = 0
both = 0
only_ans = 0
only_ana = 0
anchored = 0

for f in files:
    entries = MP.extract_entries(f)
    for e in entries:
        total_entries += 1
        lines = e['lines']
        a = any(MP.RE_ANSWER.match(ln) for ln in lines)
        n = any(MP.RE_ANALYSIS.match(ln) for ln in lines)
        if a:
            with_ans += 1
        if n:
            with_ana += 1
        if a and n:
            both += 1
        elif a:
            only_ans += 1
        elif n:
            only_ana += 1
        if e['num'] or e.get('src'):
            anchored += 1

with open('coverage_physics.txt', 'w', encoding='utf-8') as fo:
    fo.write(f'总条目: {total_entries}\n')
    fo.write(f'含答案行条目: {with_ans}（{with_ans / total_entries * 100:.1f}%）\n')
    fo.write(f'含解析行条目: {with_ana}（{with_ana / total_entries * 100:.1f}%）\n')
    fo.write(f'答案+解析都有的条目: {both}（{both / total_entries * 100:.1f}%）\n')
    fo.write(f'只有答案(无解析行)的条目: {only_ans}（{only_ans / total_entries * 100:.1f}%）\n')
    fo.write(f'只有解析(无答案行)的条目: {only_ana}（{only_ana / total_entries * 100:.1f}%）\n')
    fo.write(f'带题号/来源锚点的条目: {anchored}（{anchored / total_entries * 100:.1f}%）\n')
print(f'条目级覆盖率 -> coverage_physics.txt（总 {total_entries}，答案+解析双有 {both}，仅答案 {only_ans}，仅解析 {only_ana}，锚点 {anchored}）')
