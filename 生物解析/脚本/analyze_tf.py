#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""analyze_tf.py — 统计判断题中“没有解析/提示”的题（只读）"""
import os
import re
import collections
from docx import Document

import merge_jiexi as MJ

ROOT = os.path.dirname(os.path.abspath(__file__))

RE_TF_START_P = re.compile(r'^[（(]\s*\d+\s*[）)]')
RE_MARK = re.compile(r'[\(（][×√][\)）]')
RE_KEEP_LINE = re.compile(r'^(提示|答案|解析)\s*[　 ]')
RE_TF_JUDGE = re.compile(r'[\(（][×√][\)）]')


def analyze_one(path):
    """返回 (总条数, 有提示条数, [(题号, 题干, 标记)] 无提示列表)"""
    doc = Document(path)
    texts = MJ.normalize_split_judges([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    total = 0
    has = 0
    no_hint = []
    i = 0
    n = len(texts)
    while i < n:
        t = texts[i]
        if RE_TF_START_P.match(t) and RE_MARK.search(t):
            total += 1
            i += 1
            got = False
            while i < n and RE_KEEP_LINE.match(texts[i]):
                got = True
                i += 1
            if got:
                has += 1
            else:
                m = re.match(r'^[（(]\s*(\d+)\s*[）)]', t)
                mm = re.search(r'[\(（]([×√])[\)）]\s*$', t)
                mark = mm.group(1) if mm else '?'
                no_hint.append((m.group(1) if m else '?', t[:90], mark))
            continue
        i += 1
    return total, has, no_hint


def main():
    files = []
    for name in sorted(os.listdir(ROOT)):
        d = os.path.join(ROOT, name)
        if not os.path.isdir(d) or not name.startswith('大概念'):
            continue
        for f in sorted(os.listdir(d)):
            if not f.endswith('.docx') or f.startswith('~$') or f.endswith('_解析.docx'):
                continue
            files.append((name, f, os.path.join(d, f)))

    total = 0
    has = 0
    no_hint_items = []
    no_hint_x = []
    no_hint_ok = []
    by_day = collections.Counter()
    by_day_has = collections.Counter()

    for day, f, path in files:
        t, h, nh = analyze_one(path)
        total += t
        has += h
        by_day[day] += t
        by_day_has[day] += h
        for num, text, mark in nh:
            no_hint_items.append((day, f, num, text, mark))
            if mark == '×':
                no_hint_x.append((day, f, num, text))
            else:
                no_hint_ok.append((day, f, num, text))

    lines = []
    lines.append('判断题“无解析/提示”统计（仅原始 97 篇文档）')
    lines.append('=' * 70)
    lines.append(f'判断题条目总数: {total}')
    lines.append(f'有解析/提示: {has}  ({has/total*100:.1f}%)' if total else '有解析/提示: 0')
    lines.append(f'无解析/提示(孤题): {total - has}  ({(total-has)/total*100:.1f}%)' if total else '无解析/提示: 0')
    lines.append(f'  其中 (×) 错题无提示: {len(no_hint_x)} 条')
    lines.append(f'  其中 (√) 对题无提示: {len(no_hint_ok)} 条（原书常态：对题一般不配提示）')
    lines.append('')
    lines.append('【按大概念分布】')
    for day in sorted(by_day):
        lines.append(f'  {day}: 判断题 {by_day[day]} 条，无提示 {by_day[day] - by_day_has[day]} 条')
    lines.append('')
    lines.append('【(×) 错题却无提示（重点检查）】')
    for day, f, num, text in no_hint_x:
        lines.append(f'  {day}\\{f}  ({num}) {text[:80]}')
    lines.append('')
    lines.append(f'【(√) 对题无提示，共 {len(no_hint_ok)} 条（示例前 15 条）】')
    for day, f, num, text in no_hint_ok[:15]:
        lines.append(f'  {day}\\{f}  ({num}) {text[:70]}')

    with open('analyze_tf.txt', 'w', encoding='utf-8') as fw:
        fw.write('\n'.join(lines) + '\n')
    print('\n'.join(lines))


if __name__ == '__main__':
    main()
