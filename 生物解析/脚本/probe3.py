#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe3.py — 探查“孤立标签”段（答案/解析/提示 单独成段）的上下文（只读）
输出 audit_probe3.txt：标签段索引、下一段（可能是真实内容）。
"""
import os
from docx import Document

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'audit_probe3.txt')
MAX = 160


def paras(path):
    return [p.text.strip() for p in Document(path).paragraphs if p.text.strip()]


def main():
    lines = []
    for name in sorted(os.listdir('.')):
        d = os.path.join('.', name)
        if not os.path.isdir(d) or not name.startswith('大概念'):
            continue
        for fname in sorted(os.listdir(d)):
            if not fname.endswith('.docx') or fname.startswith('~$') or fname.endswith('_解析.docx'):
                continue
            ps = paras(os.path.join(d, fname))
            for i, t in enumerate(ps):
                if t.strip() in ('答案', '解析', '提示') or t.strip() in ('答案：', '解析：', '提示：'):
                    lines.append(f'--- {name}\\{fname}  #{i} 标签: {t}')
                    for j in range(i + 1, min(i + 3, len(ps))):
                        lines.append(f'      #{j}: {ps[j][:MAX]}')

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print('done ->', OUT)


if __name__ == '__main__':
    main()
