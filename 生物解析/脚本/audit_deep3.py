#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_deep3.py — 第三轮：确认“多段答案/提示被截断”确凿案例与产物实缺证据。
只读。输出 audit_deep3.txt
"""

import os
import re
from docx import Document

import audit_direct as AD

OUT = os.path.join(AD.ROOT, 'audit_deep3.txt')
MAX = 220

# 确凿续段：前一条是 答案/提示/解析 (数字) 开头，本条以更大序号 (数字)/(1) 开头
PREV_NUM = re.compile(r'^(答案|提示|解析)\s*[　 ]*[（(]\s*(\d+)\s*[）)]')
CUR_NUM = re.compile(r'^[（(]\s*(\d+)\s*[）)]')


def paragraphs(path):
    return [p.text.strip() for p in Document(path).paragraphs if p.text.strip()]


def main():
    doc_files = []
    for name in sorted(os.listdir(AD.ROOT)):
        if not os.path.isdir(os.path.join(AD.ROOT, name)) or not name.startswith('大概念'):
            continue
        for fname in sorted(os.listdir(os.path.join(AD.ROOT, name))):
            if not fname.endswith('.docx') or fname.startswith('~$') or fname.endswith('_解析.docx'):
                continue
            doc_files.append(os.path.join(AD.ROOT, name, fname))

    out = []
    definite = 0
    definite_files = set()
    examples = []

    for path in doc_files:
        rel = os.path.relpath(path, AD.ROOT)
        paras = paragraphs(path)
        for i in range(1, len(paras)):
            m_prev = PREV_NUM.match(paras[i - 1])
            m_cur = CUR_NUM.match(paras[i])
            if not m_prev or not m_cur:
                continue
            pn, cn = int(m_prev.group(2)), int(m_cur.group(1))
            if cn > pn and cn - pn <= 4:
                definite += 1
                definite_files.add(rel)
                if len(examples) < 60:
                    examples.append((rel, paras[i - 1], paras[i]))

        # 产物确认：该文件 _解析.docx 是否缺这些续行
        name, _ = os.path.splitext(path)
        jp = name + '_解析.docx'
        if os.path.exists(jp):
            jiexi = [p.text.strip() for p in Document(jp).paragraphs if p.text.strip()]
            for i in range(1, len(paras)):
                m_prev = PREV_NUM.match(paras[i - 1])
                m_cur = CUR_NUM.match(paras[i])
                if not m_prev or not m_cur:
                    continue
                pn, cn = int(m_prev.group(2)), int(m_cur.group(1))
                if cn > pn and cn - pn <= 4:
                    if paras[i] not in jiexi:
                        out.append(f'【产物实缺】{rel}')
                        out.append(f'   保留: {paras[i-1][:MAX]}')
                        out.append(f'   丢失: {paras[i][:MAX]}')

    out.append('')
    out.append('=' * 70)
    out.append(f'【确凿多段答案/提示截断】共 {definite} 行 / {len(definite_files)} 个文件')
    out.append('（即：原文 答案(1)→(2)…、提示(1)→(2)… 分开成多段，产物只保留第一段）')
    out.append('')
    out.append('【样例】')
    for rel, a, b in examples[:40]:
        out.append(f'  {rel}')
        out.append(f'    保: {a[:120]}')
        out.append(f'    丢: {b[:120]}')

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out) + '\n')
    print(f'done -> {OUT}')


if __name__ == '__main__':
    main()
