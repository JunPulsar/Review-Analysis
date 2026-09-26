#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe2.py — 判定“纯文字续段”规则（只读）：输出 audit_probe2.txt
候选：前一条 = 答案/解析 行（非编号开头），本条不匹配题号/标题/编号标记，
统计前条结尾特征并抽样，判断哪些是真解析续段、哪些是小标题/题干。
"""
import os
import re
from docx import Document

import audit_direct as AD

OUT = os.path.join(AD.ROOT, 'audit_probe2.txt')
MAX = 120

HEADING_RE = re.compile(
    r'^(考向|考点|限时练|阶段|周末必刷|第\d+讲|高考共鸣点|回扣落实|'
    r'【|】|综合提升|巩固|基础|大概念|本讲|课标|目录|练\s*[··]|'
    r'思维|方法|技巧|易错|警示|归纳|总结|拓展|延伸|微网|构建)'
)
NUM_MARK = re.compile(r'^[（(]\s*\d+\s*[）)]|^[①②③④⑤⑥⑦⑧⑨⑩]|^\d+[.、．]')


def paras(path):
    return [p.text.strip() for p in Document(path).paragraphs if p.text.strip()]


def main():
    rows = []
    for name in sorted(os.listdir(AD.ROOT)):
        d = os.path.join(AD.ROOT, name)
        if not os.path.isdir(d) or not name.startswith('大概念'):
            continue
        for fname in sorted(os.listdir(d)):
            if not fname.endswith('.docx') or fname.startswith('~$') or fname.endswith('_解析.docx'):
                continue
            ps = paras(os.path.join(d, fname))
            for i in range(1, len(ps)):
                prev, cur = ps[i - 1], ps[i]
                if not (AD.RE_KEEP_STRICT.match(prev) and
                        (prev.startswith('答案') or prev.startswith('解析'))):
                    continue
                if AD.RE_KEEP_STRICT.match(cur) or NUM_MARK.match(cur) or HEADING_RE.match(cur) \
                        or AD.is_title(cur) or AD.RE_Q_NOYEAR.match(cur):
                    continue
                # 上一条是否编号开头（规则A已覆盖）
                if re.match(r'^(答案|解析)\s*[　 ]*[（(]?\d+[）)]', prev):
                    continue
                rows.append((name + '\\' + fname, prev, cur))

    # 统计前条结尾
    end_ok = sum(1 for _, p, _ in rows if p[-1] in '。？！？)）】')
    end_other = len(rows) - end_ok

    out = []
    out.append(f'候选纯文字续段: {len(rows)} 条；前条以 。？！) 结尾: {end_ok} 条；其他结尾: {end_other} 条')
    out.append('=' * 70)
    out.append('【其他结尾（可能是真续段）】')
    for f, p, c in rows:
        if p[-1] in '。？！？)）】':
            continue
        out.append(f'--- {f}')
        out.append(f'    前尾: ...{p[-40:]}')
        out.append(f'    本条: {c[:MAX]}')
    out.append('=' * 70)
    out.append('【。结尾（抽样前 80）】')
    cnt = 0
    for f, p, c in rows:
        if p[-1] in '。？！？)）】':
            if cnt < 80:
                out.append(f'--- {f}')
                out.append(f'    前尾: ...{p[-40:]}')
                out.append(f'    本条: {c[:MAX]}')
            cnt += 1

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out) + '\n')
    print('done ->', OUT)


if __name__ == '__main__':
    main()
