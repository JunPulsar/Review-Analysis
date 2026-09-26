#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe.py — 结构探查（只读）：输出 audit_probe.txt
1) 关键文件全段落顺序（前 260 段）：第1讲、第22讲、第30讲、第38讲、周末必刷16
2) 全部“续段·待人工看”（前一条 strict + 本条无编号无标题）行，供判定规则
"""
import os
import re
from docx import Document

import audit_direct as AD

OUT = os.path.join(AD.ROOT, 'audit_probe.txt')
MAX = 200

TARGETS = [
    '大概念一 细胞是生物体结构与生命活动的基本单位\\第1讲　走近细胞.docx',
    '大概念四　遗传信息控制生物性状并代代相传——遗传规律与人类遗传病\\第22讲　人类遗传病及遗传系谱图分析.docx',
    '大概念七　个体生命活动的调节\\第30讲　人体的内环境与稳态.docx',
    '大概念七　个体生命活动的调节\\第38讲　其他植物激素.docx',
    '大概念七　个体生命活动的调节\\周末必刷16.docx',
]

HEADING_RE = re.compile(
    r'^(考向|考点|限时练|阶段|周末必刷|第\d+讲|高考共鸣点|回扣落实|'
    r'【|】|综合提升|巩固|基础|大概念|本讲|课标|目录|练\s*[··]|'
    r'思维|方法|技巧|易错|警示|归纳|总结|拓展|延伸|微网|构建)'
)


def paras(path):
    return [p.text.strip() for p in Document(path).paragraphs if p.text.strip()]


def main():
    out = []
    for rel in TARGETS:
        p = os.path.join(AD.ROOT, rel.replace('\\', os.sep))
        if not os.path.exists(p):
            out.append(f'!! 不存在 {rel}')
            continue
        ps = paras(p)
        out.append('=' * 70)
        out.append(f'【全段落】{rel}  ({len(ps)} 段)')
        for i, t in enumerate(ps[:260]):
            flag = ''
            kind, _ = AD.classify(t)
            if kind:
                flag = f'[{kind}]'
            elif AD.is_title(t):
                flag = '[标题]'
            elif i > 0 and AD.RE_KEEP_STRICT.match(ps[i - 1]) and not HEADING_RE.match(t):
                flag = '[待看]'
            elif HEADING_RE.match(t):
                flag = '[标题类]'
            out.append(f'{i:03d} {flag} {t[:MAX]}')

    # 全部待看续段
    out.append('=' * 70)
    out.append('【全部“待看”续段 汇总】')
    for name in sorted(os.listdir(AD.ROOT)):
        d = os.path.join(AD.ROOT, name)
        if not os.path.isdir(d) or not name.startswith('大概念'):
            continue
        for fname in sorted(os.listdir(d)):
            if not fname.endswith('.docx') or fname.startswith('~$') or fname.endswith('_解析.docx'):
                continue
            ps = paras(os.path.join(d, fname))
            for i in range(1, len(ps)):
                if AD.RE_KEEP_STRICT.match(ps[i - 1]):
                    kind2, _ = AD.classify(ps[i])
                    if not kind2 and not AD.is_title(ps[i]) and not AD.RE_Q_NOYEAR.match(ps[i]) \
                            and not HEADING_RE.match(ps[i]):
                        out.append(f'--- {name}\\{fname}')
                        out.append(f'    前: ...{ps[i-1][-70:]}')
                        out.append(f'    本: {ps[i][:MAX]}')

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out) + '\n')
    print('done ->', OUT)


if __name__ == '__main__':
    main()
