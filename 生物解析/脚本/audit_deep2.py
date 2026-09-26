#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_deep2.py — 第二轮深度核查：修正“年份括号全角/半角”误判，分类续段真伪，
并直接对照 _解析.docx / 解析合集 产物与现行代码的差异。只读，不改任何 docx。
输出 audit_deep2.txt
"""

import os
import re
import collections
from docx import Document

import audit_direct as AD

OUT = os.path.join(AD.ROOT, 'audit_deep2.txt')
MAX = 300

# 修正：年份括号同时认全角/半角
RE_YEAR2 = re.compile(r'^\d+\.\s*[（(]\d{4}')
RE_Q2 = re.compile(r'^\d+\.\s*\S')

# 续段分类：小标题特征
HEADING_RE = re.compile(
    r'^(考向|考点|限时练|阶段|周末必刷|第\d+讲|高考共鸣点|回扣落实|'
    r'【|】|综合提升|巩固|基础|大概念|本讲|课标|目录|练\s*[··]|'
    r'思维|方法|技巧|易错|警示|归纳|总结|拓展|延伸|微网|构建)'
)
ANS_PART_RE = re.compile(r'^[（(]\s*\d+\s*[）)]|^[①②③④⑤⑥⑦⑧⑨⑩]|^第[一二三四五六七八九十]|[（(]1[）)]')


def paragraphs(path):
    doc = Document(path)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def answer_analysis_lines(paras):
    return [t for t in paras if AD.RE_KEEP_STRICT.match(t) and (t.startswith('答案') or t.startswith('解析'))]


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
    files_real_noyear = []
    tot_real_noyear = 0
    files_cont_real = set()
    tot_cont_real = 0
    files_cont_prose = set()
    files_cont_heading = set()
    files_artifact_diff = []
    tot_artifact_diff_lines = 0
    files_nospace = set()
    tot_nospace = 0

    for path in doc_files:
        rel = os.path.relpath(path, AD.ROOT)
        paras = paragraphs(path)

        # ── 1) 真实无年份题号（修正全角年份后） ──
        real_noyear = []
        for i, t in enumerate(paras):
            if not (RE_Q2.match(t) and not RE_YEAR2.match(t)):
                continue
            j = i + 1
            saw = None
            while j < len(paras):
                nt = paras[j]
                if AD.RE_KEEP_STRICT.match(nt) and (nt.startswith('答案') or nt.startswith('解析')):
                    saw = nt
                    break
                if RE_YEAR2.match(nt) or (RE_Q2.match(nt) and not AD.RE_KEEP_STRICT.match(nt)):
                    break
                j += 1
            if saw:
                real_noyear.append((t, saw))
        if real_noyear:
            tot_real_noyear += len(real_noyear)
            files_real_noyear.append(rel)
            out.append(f'【真实无年份题号】{rel} 共{len(real_noyear)}条')
            for q, a in real_noyear[:5]:
                out.append(f'  题: {q[:120]}')
                out.append(f'  后: {a[:120]}')

        # ── 2) 续段分类 ──
        for i in range(1, len(paras)):
            if AD.RE_KEEP_STRICT.match(paras[i - 1]):
                kind2, _ = AD.classify(paras[i])
                if kind2 or AD.is_title(paras[i]) or RE_Q2.match(paras[i]):
                    continue
                nxt = paras[i]
                if HEADING_RE.match(nxt):
                    files_cont_heading.add(rel)
                elif ANS_PART_RE.match(nxt):
                    tot_cont_real += 1
                    files_cont_real.add(rel)
                    out.append(f'【续段·疑似真截断】{rel}')
                    out.append(f'   前条(尾60): ...{paras[i-1][-60:]}')
                    out.append(f'   本条: {nxt[:MAX]}')
                else:
                    files_cont_prose.add(rel)

        # ── 3) nospace 行 ──
        for t in paras:
            kind, _ = AD.classify(t)
            if kind == 'nospace':
                tot_nospace += 1
                files_nospace.add(rel)
                out.append(f'【nospace】{rel}: {t[:MAX]}')

        # ── 4) _解析.docx 产物 vs 现行 process 规则 ──
        name, ext = os.path.splitext(path)
        jp = name + '_解析.docx'
        if os.path.exists(jp):
            jiexi = [p.text.strip() for p in Document(jp).paragraphs if p.text.strip()]
            keep_now = set()
            for t in paras:
                if AD.PD.is_keep_line(t) or AD.is_title(t):
                    keep_now.add(t)
            missing = [t for t in keep_now if t not in set(jiexi)]
            extra = [t for t in jiexi if t not in keep_now]
            if missing or extra:
                tot_artifact_diff_lines += len(missing) + len(extra)
                files_artifact_diff.append(rel)
                out.append(f'【产物与现行规则不同】{rel} 缺{len(missing)} 多{len(extra)}')
                for m in missing[:5]:
                    out.append(f'   规则应保留但产物无: {m[:MAX]}')
                for e in extra[:5]:
                    out.append(f'   产物有但规则不应保留: {e[:MAX]}')

        # ── 5) 表格内候选 ──
        tf = AD.table_candidates(Document(path))
        if tf:
            out.append(f'【表格内】{rel}')
            for kind, t in tf:
                out.append(f'   [{kind}] {t}')

    # ── 6) 解析合集 与 现行 merge 的差异（按答案/解析行计数） ──
    heji_dir = os.path.join(AD.ROOT, '解析合集')
    if os.path.isdir(heji_dir):
        out.append('【解析合集对照】')
        for dname in sorted(os.listdir(heji_dir)):
            dpath = os.path.join(heji_dir, dname)
            if not os.path.isdir(dpath):
                continue
            for cat in ['第X讲.docx', '高考共鸣点.docx', '阶段排查.docx', '周末必刷.docx']:
                fpath = os.path.join(dpath, cat)
                if not os.path.exists(fpath):
                    continue
                try:
                    hdoc = Document(fpath)
                    h_ans = sum(1 for p in hdoc.paragraphs
                                if AD.RE_KEEP_STRICT.match(p.text.strip())
                                and (p.text.strip().startswith('答案') or p.text.strip().startswith('解析')))
                except Exception as e:
                    h_ans = -1
                # 现行 merge 对该大概念该分类可抽到的答案/解析行数
                exp = 0
                for path in doc_files:
                    if os.path.basename(os.path.dirname(path)) == dname and os.path.basename(path).startswith(cat[0]):
                        pass
                out.append(f'   {dname}/{cat}: 产物答案/解析行={h_ans}')

    summary = (
        '\n' + '=' * 70 + '\n'
        '【深度统计2】\n'
        f'真实无年份题号(修正全角年份后): {tot_real_noyear} 行 / {len(files_real_noyear)} 个文件\n'
        f'续段疑似真截断(答案续行带(1)/①): {tot_cont_real} 行 / {len(files_cont_real)} 个文件\n'
        f'续段遇到小标题(考点/考向等，误报): {len(files_cont_heading)} 个文件\n'
        f'续段待人工看: {len(files_cont_prose)} 个文件\n'
        f'nospace: {tot_nospace} 行 / {len(files_nospace)} 个文件\n'
        f'产物与现行规则差异: {tot_artifact_diff_lines} 行 / {len(files_artifact_diff)} 个文件\n'
    )
    out.append(summary)

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('\n'.join(out) + '\n')
    print(f'done -> {OUT}')


if __name__ == '__main__':
    main()
