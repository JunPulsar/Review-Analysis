#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_direct.py — 直读核查脚本（只读，不修改任何 docx）

目的：把原始 docx 的段落文本直接抽出来看，找出"看起来是 答案/解析/提示/判断题、
但按现有脚本（process_docs.py / merge_jiexi.py）的规则会被漏掉或误判"的行，
并与脚本实际产物（_解析.docx、解析合集/）对照，给出易错点证据。

判定分类（多打的"宽松候选"，比脚本严格规则更宽）：
  strict   = 答案/解析/提示 + 全角/半角空格     （脚本一般能保留）
  colon    = 答案/解析/提示 + 中文/半角冒号      （脚本识别不了 → 漏）
  nospace  = 答案/解析/提示 后面无分隔           （脚本识别不了 → 漏）
  tf_miss  = 判断题标记 (×/√)，但题号是全角/复合形式（脚本识别不了 → 漏）
  cont     = 上一条是"解析/答案/提示"行，下一条疑似续段（无前缀 → 会被截断）
  table    = 表格单元格里出现的 答案/解析/提示/判断题（两脚本都整体丢表格）
  q_noyear = 形如 "1. xxx" 但不带 (2024… 年份的题号（merge 的典型例题识别不了）

产物：
  audit_report.txt            全量统计 + 易错行分类计数 + 受影响文件清单
  audit_samples/<名>.txt      每个"有疑点"文件：原文候选行 vs _解析.docx 实际保留行
"""

import os
import re
import collections
from docx import Document

import process_docs as PD
import merge_jiexi as MJ

ROOT = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(ROOT, 'audit_samples')
REPORT_PATH = os.path.join(ROOT, 'audit_report.txt')

MAX_SHOW = 160  # 样例文本最大显示长度

# ── 宽松候选正则 ──
RE_KEEP_STRICT = re.compile(r'^(答案|解析|提示)\s*[　 ]')
RE_COLON       = re.compile(r'^(答案|解析|提示)\s*[:：]')
RE_NOSPACE     = re.compile(r'^(答案|解析|提示)(?!\s*[　 :：])')
RE_TF_HALF     = re.compile(r'^\(\d+\)')
RE_TF_ANY      = re.compile(r'^[（(]\s*\d+[）)]')
RE_TF_COMP     = re.compile(r'^\d+[、.．]\s*[（(]?\s*\d+\s*[）)]')
RE_MARK        = re.compile(r'[×√]')
RE_Q_YEAR      = re.compile(r'^\d+\.\s*\(\d{4}')
RE_Q_NOYEAR    = re.compile(r'^\d+\.\s*\S')
TITLE_RES = [
    re.compile(r'^第\d+讲'),
    re.compile(r'^周末必刷\d+'),
    re.compile(r'^阶段排查'),
    re.compile(r'^高考共鸣点\d+'),
]


def is_title(text):
    return any(p.match(text) for p in TITLE_RES)


def classify(text):
    """返回 (主类别, 是否被 process 严格规则保留)；不是候选则返回 (None, False)"""
    if RE_KEEP_STRICT.match(text):
        return ('strict', True)
    if RE_COLON.match(text):
        return ('colon', False)
    if RE_NOSPACE.match(text):
        return ('nospace', False)
    if RE_MARK.search(text):
        if RE_TF_ANY.match(text) or RE_TF_COMP.match(text):
            kept = bool(RE_TF_HALF.match(text) and not RE_TF_COMP.match(text))
            return ('tf', kept)
    if RE_Q_NOYEAR.match(text) and not RE_Q_YEAR.match(text):
        return ('q_noyear', False)
    return (None, False)


def table_candidates(doc):
    """收集表格单元格内的候选行（两脚本都会丢）"""
    found = []
    seen = set()
    try:
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    # 以单元格文本为 key 去重（lxml 代理的 id() 会被 GC 复用，不可靠）
                    key = tuple(p.text.strip() for p in cell.paragraphs if p.text.strip())
                    if not key or key in seen:
                        continue
                    seen.add(key)
                    for p in cell.paragraphs:
                        t = p.text.strip()
                        if not t:
                            continue
                        kind, kept = classify(t)
                        if kind in ('strict', 'colon', 'nospace', 'tf'):
                            found.append((kind, t[:MAX_SHOW]))
    except Exception:
        pass
    return found


def get_jiexi_paras(input_path):
    """读取对应 _解析.docx 的非空段落；文件不存在返回 None"""
    name, ext = os.path.splitext(input_path)
    out_path = name + '_解析.docx'
    if not os.path.exists(out_path):
        return None
    try:
        doc = Document(out_path)
        return [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    except Exception as e:
        return [f'<读取失败: {e}>']


def main():
    os.makedirs(SAMPLE_DIR, exist_ok=True)

    # 收集所有原始 docx
    doc_files = []
    for name in sorted(os.listdir(ROOT)):
        if not os.path.isdir(os.path.join(ROOT, name)) or not name.startswith('大概念'):
            continue
        for fname in sorted(os.listdir(os.path.join(ROOT, name))):
            if not fname.endswith('.docx') or fname.startswith('~$'):
                continue
            if fname.endswith('_解析.docx'):
                continue
            doc_files.append(os.path.join(ROOT, name, fname))

    stats = collections.Counter()       # 各分类的行数
    files_miss = collections.defaultdict(set)    # kind -> {file}
    files_table = []
    files_cont = []
    files_only_title = []
    files_missing_jiexi = []
    files_empty_merge = []
    files_noyear = set()
    samples = []  # (file, 行列表)

    for path in doc_files:
        rel = os.path.relpath(path, ROOT)
        try:
            doc = Document(path)
        except Exception as e:
            stats['open_fail'] += 1
            files_miss['open_fail'].add(rel)
            continue

        paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        flags = []
        table_found = table_candidates(doc)
        cont_flags = []

        for i, t in enumerate(paras):
            kind, kept = classify(t)
            if kind:
                stats[kind] += 1
                flags.append((kind, t[:MAX_SHOW]))
                if not kept and kind in ('colon', 'nospace', 'tf'):
                    files_miss[kind].add(rel)
                if kind == 'q_noyear':
                    files_noyear.add(rel)
            # 续段截断：上一条 strict 开头，本条无任何前缀/题号/标题特征
            if i > 0 and RE_KEEP_STRICT.match(paras[i - 1]):
                kind2, _ = classify(paras[i])
                if not kind2 and not is_title(paras[i]) and not RE_Q_NOYEAR.match(paras[i]):
                    cont_flags.append((i, paras[i][:MAX_SHOW]))

        if cont_flags:
            stats['cont'] += len(cont_flags)
            files_cont.append(rel)
            flags += [('cont', c[1]) for c in cont_flags]

        if table_found:
            stats['table'] += len(table_found)
            files_table.append(rel)
            flags += [('table', t) for _, t in table_found]

        jiexi = get_jiexi_paras(path)
        if jiexi is None:
            stats['no_jiexi'] += 1
            files_missing_jiexi.append(rel)
        elif len(jiexi) <= 1:
            stats['only_title'] += 1
            files_only_title.append(rel)

        # merge 检查：当前代码能抽多少条目/行（不写盘）
        try:
            entries = MJ.process_one_docx(path)
            stats['merge_entries'] += len(entries)
            stats['merge_lines'] += sum(len(e['lines']) for e in entries)
            if not entries:
                files_empty_merge.append(rel)
        except Exception as e:
            stats['merge_fail'] += 1
            files_miss['merge_fail'].add(rel)
            entries = []

        # 有疑点才写对照样本
        suspicious = [f for f in flags if f[0] in ('colon', 'nospace', 'tf', 'cont', 'table')] \
                     or (jiexi is not None and len(jiexi) <= 1)
        if suspicious:
            samples.append((rel, flags, jiexi, entries))

    # ── 写入样本文件 ──
    for rel, flags, jiexi, entries in samples:
        safe = rel.replace('\\', '__').replace('/', '__').replace('.docx', '')
        sp = os.path.join(SAMPLE_DIR, safe + '.txt')
        with open(sp, 'w', encoding='utf-8') as f:
            f.write(f'文件: {rel}\n')
            f.write(f'疑点行数: {len(flags)}\n' + '=' * 70 + '\n')
            f.write('【原文疑点行（类别在前）】\n')
            for kind, text in flags:
                f.write(f'[{kind}] {text}\n')
            f.write('\n【_解析.docx 实际保留内容】\n')
            if jiexi is None:
                f.write('（不存在）\n')
            else:
                for t in jiexi:
                    f.write(t[:MAX_SHOW] + '\n')
            f.write('\n【merge 当前代码抽取条目数】\n')
            f.write(f'{len(entries)} 条，共 {sum(len(e["lines"]) for e in entries)} 行\n')

    # ── 写入汇总报告 ──
    lines = []
    lines.append('audit_direct.py 直读核查报告')
    lines.append('=' * 70)
    lines.append(f'扫描原始 docx: {len(doc_files)} 个')
    lines.append(f"打开失败: {stats['open_fail']}")
    lines.append('')
    lines.append('【易错行统计（宽松候选）】')
    for k in ['strict', 'colon', 'nospace', 'tf', 'cont', 'table', 'q_noyear']:
        lines.append(f'  {k:<9}: {stats[k]} 行')
    lines.append('')
    lines.append('【merge 按当前代码可抽取】')
    lines.append(f"  条目: {stats['merge_entries']} 条 / 行: {stats['merge_lines']} 行 / 抽取失败文件: {stats['merge_fail']}")
    lines.append('')
    lines.append('【产物完整性】')
    lines.append(f'  缺 _解析.docx: {len(files_missing_jiexi)} 个 {files_missing_jiexi[:10]}')
    lines.append(f'  _解析.docx 只剩标题/近空: {len(files_only_title)} 个 {files_only_title[:10]}')
    lines.append(f'  merge 抽不到任何条目: {len(files_empty_merge)} 个 {files_empty_merge[:10]}')
    lines.append('')
    lines.append('【按类别的受影响文件】')
    for k in ['colon', 'nospace', 'tf', 'cont', 'table', 'q_noyear']:
        lines.append(f'  {k} -> {len(files_miss[k])} 个文件')
        for r in sorted(files_miss[k])[:15]:
            lines.append(f'      {r}')
    if files_table:
        lines.append(f'  表格内有候选 -> {len(files_table)} 个文件')
        for r in files_table[:15]:
            lines.append(f'      {r}')
    if files_cont:
        lines.append(f'  疑似续段截断 -> {len(files_cont)} 个文件')
        for r in files_cont[:15]:
            lines.append(f'      {r}')
    if files_noyear:
        lines.append(f'  无年份题号 -> {len(files_noyear)} 个文件')
        for r in sorted(files_noyear)[:15]:
            lines.append(f'      {r}')
    lines.append('')
    lines.append(f'【对照样本】共 {len(samples)} 个文件写入 audit_samples/，见各 txt。')

    report = '\n'.join(lines)
    with open(REPORT_PATH, 'w', encoding='utf-8') as f:
        f.write(report + '\n')

    print(report)
    print('\n详细对照见: audit_samples/')


if __name__ == '__main__':
    main()
