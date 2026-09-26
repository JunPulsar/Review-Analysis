#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_deep.py — 深度直读：针对 audit_direct.py 标出的疑点，输出完整上下文证据。

只读，不修改任何 docx。输出 audit_deep.txt。

覆盖：
 1) 真正的“无年份题号 + 其后确有答案/解析”行（merge 典型例题会漏）—— 全量统计
 2) merge 条目覆盖不到 的 答案/解析 行（按文本消费计数）—— 全量统计
 3) process 漏判的 nospace 行及其上下文
 4) 表格内候选行全文（两脚本都丢）
 5) 疑似续段行：上一段 解析/答案 行 + 本段完整文本（人工判断真假截断）
 6) 周末必刷16.docx 全文结构（merge 抽不到任何条目）
"""

import os
import collections
from docx import Document

import audit_direct as AD
import merge_jiexi as MJ

OUT = os.path.join(AD.ROOT, 'audit_deep.txt')
MAX = 400  # 长行截断显示


def paragraphs(path):
    doc = Document(path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def real_noyear(paras):
    """返回 (题号行idx, 题号行, 其后的答案/解析行) 列表：无年份题号后确有答案/解析"""
    res = []
    for i, t in enumerate(paras):
        if not (AD.RE_Q_NOYEAR.match(t) and not AD.RE_Q_YEAR.match(t)):
            continue
        j = i + 1
        saw = None
        while j < len(paras):
            nt = paras[j]
            if AD.RE_KEEP_STRICT.match(nt) and (nt.startswith('答案') or nt.startswith('解析')):
                saw = nt
                break
            if AD.RE_Q_YEAR.match(nt) or (AD.RE_Q_NOYEAR.match(nt) and not AD.RE_KEEP_STRICT.match(nt)):
                break
            j += 1
        if saw:
            res.append((i, t, saw))
    return res


def merge_cover(paras, entries):
    """merge 条目覆盖不到的 答案/解析 行（文本消费计数）"""
    counter = collections.Counter()
    for e in entries:
        for line in e['lines']:
            counter[line] += 1
    missed = []
    for t in paras:
        if AD.RE_KEEP_STRICT.match(t) and (t.startswith('答案') or t.startswith('解析')):
            if counter[t] > 0:
                counter[t] -= 1
            else:
                missed.append(t)
    return missed


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
    tot_noyear_q = 0
    tot_noyear_files = set()
    tot_merge_miss = 0
    tot_merge_miss_files = set()
    tot_nospace = 0
    tot_table_files = set()

    for path in doc_files:
        rel = os.path.relpath(path, AD.ROOT)
        paras = paragraphs(path)

        # 1) 真正无年份题号
        ny = real_noyear(paras)
        if ny:
            tot_noyear_q += len(ny)
            tot_noyear_files.add(rel)
            out.append(f'【无年份题号(有答案)】{rel}  共 {len(ny)} 条')
            for idx, q, a in ny[:8]:
                out.append(f'   #{idx} 题号: {q[:MAX]}')
                out.append(f'       答案/解析: {a[:MAX]}')

        # 2) merge 覆盖不到的答案/解析
        entries = MJ.process_one_docx(path)
        missed = merge_cover(paras, entries)
        if missed:
            tot_merge_miss += len(missed)
            tot_merge_miss_files.add(rel)
            out.append(f'【merge 漏答案/解析】{rel}  共 {len(missed)} 行')
            for m in missed[:8]:
                out.append(f'   {m[:MAX]}')

        # 3) nospace 上下文
        for i, t in enumerate(paras):
            kind, kept = AD.classify(t)
            if kind == 'nospace':
                tot_nospace += 1
                out.append(f'【nospace】{rel}')
                out.append(f'   行#{i}: {t[:MAX]}')
                if i > 0:
                    out.append(f'   前一行: {paras[i-1][:60]}')

        # 4) 表格内候选全文
        tf = AD.table_candidates(Document(path))
        if tf:
            tot_table_files.add(rel)
            out.append(f'【表格内候选】{rel}  共 {len(tf)} 行')
            for kind, t in tf:
                out.append(f'   [{kind}] {t}')

        # 5) 续段上下文（前一条解析行 + 本段）
        for i in range(1, len(paras)):
            if AD.RE_KEEP_STRICT.match(paras[i - 1]):
                kind2, _ = AD.classify(paras[i])
                if not kind2 and not AD.is_title(paras[i]) and not AD.RE_Q_NOYEAR.match(paras[i]):
                    prev = paras[i - 1]
                    out.append(f'【续段】{rel}')
                    out.append(f'   解析行#{i-1}(尾80): ...{prev[-80:]}')
                    out.append(f'   续段#{i}(头200): {paras[i][:200]}')

    # 6) 周末必刷16 全文
    special = None
    for p in doc_files:
        if p.endswith('周末必刷16.docx'):
            special = p
            break
    if special:
        out.append('=' * 70)
        out.append('【周末必刷16.docx 全文段落（每段非空文本，最长300字）】')
        for i, t in enumerate(paragraphs(special)):
            out.append(f'#{i:03d} {t[:300]}')

    summary = (
        '\n' + '=' * 70 + '\n'
        '【深度统计】\n'
        f'真实无年份题号(后确有答案/解析): {tot_noyear_q} 行 / {len(tot_noyear_files)} 个文件\n'
        f'merge 覆盖不到的 答案/解析 行: {tot_merge_miss} 行 / {len(tot_merge_miss_files)} 个文件\n'
        f'nospace 候选行: {tot_nospace} 行\n'
        f'表格内候选文件: {len(tot_table_files)} 个\n'
    )
    out.append(summary)

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out) + '\n')
    print(f'写完成: {OUT}（统计见文末）')


if __name__ == '__main__':
    main()
