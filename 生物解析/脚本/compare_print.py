# -*- coding: utf-8 -*-
"""compare_print.py — 打印量对比：原始全部 vs 新整合解析全集（只读统计）"""
import os
from docx import Document

ROOT = os.path.dirname(os.path.abspath(__file__))


def chars_of(path):
    doc = Document(path)
    return sum(len(p.text) for p in doc.paragraphs)


def chars_of_dir(d, sub=None):
    """统计大概念文件夹内原始 docx（排除 _解析、~$）"""
    total = 0
    files = []
    for f in sorted(os.listdir(d)):
        if not f.endswith('.docx') or f.startswith('~$') or f.endswith('_解析.docx'):
            continue
        files.append(f)
        total += chars_of(os.path.join(d, f))
    return total, files


# ── 原始全部（97 篇，含周末必刷）──
orig_chars = 0
orig_files = 0
orig_chars_zm = 0
zm_files = []
for name in sorted(os.listdir(ROOT)):
    d = os.path.join(ROOT, name)
    if not os.path.isdir(d) or not name.startswith('大概念'):
        continue
    c, fs = chars_of_dir(d)
    orig_chars += c
    orig_files += len(fs)
    for f in fs:
        if f.startswith('周末必刷'):
            zm_files.append(f)
            orig_chars_zm += chars_of(os.path.join(d, f))

# ── 新整合全书（75 章，不含周末必刷）──
book = os.path.join(ROOT, '新整合_全书', '解析合集_全书.docx')
book_chars = chars_of(book)

# ── 周末必刷解析版（原 解析合集 的周末必刷.docx 合并）──
zm_merged = 0
zm_docs = []
for name in sorted(os.listdir(os.path.join(ROOT, '解析合集'))):
    fp = os.path.join(ROOT, '解析合集', name, '周末必刷.docx')
    if os.path.exists(fp):
        zm_merged += chars_of(fp)
        zm_docs.append(fp)

# ── 页数估算（沿用 stats.py 口径：原始教辅约1400字/页×1.67图表修正；紧凑排版约2200字/页）──
orig_pages = orig_chars / 1400 * 1.67
book_pages = book_chars / 2200
zm_pages = (orig_chars_zm / 1400 * 1.67)

lines = []
lines.append('打印量对比（估算）')
lines.append('=' * 62)
lines.append(f'【原始全部】{orig_files} 篇，段落字符合计 {orig_chars:,}')
lines.append(f'  其中 周末必刷 22 篇字符: {orig_chars_zm:,}')
lines.append(f'  按教辅排版估算页数: ~{orig_pages:,.0f} 页')
lines.append('')
lines.append(f'【新整合全书】75 章（52讲+14共鸣点+9阶段排查，不含周末必刷）')
lines.append(f'  段落字符合计 {book_chars:,}')
lines.append(f'  按紧凑排版估算页数: ~{book_pages:,.0f} 页')
lines.append('')
lines.append('【对比（新整合 vs 原始全部）】')
lines.append(f'  字符: {orig_chars:,} -> {book_chars:,}  保留 {book_chars/orig_chars*100:.1f}%')
lines.append(f'  字符减少: {orig_chars - book_chars:,}（减少 {(1-book_chars/orig_chars)*100:.1f}%）')
lines.append(f'  页数: ~{orig_pages:,.0f} -> ~{book_pages:,.0f} 页')
lines.append(f'  节省约 {orig_pages - book_pages:,.0f} 页，节省比例 {(1-book_pages/orig_pages)*100:.1f}%')
lines.append('')
lines.append('【补充口径】')
lines.append(f'  若加上周末必刷解析版（解析合集/*/周末必刷.docx，{len(zm_docs)} 份合并）:')
lines.append(f'    字符 {zm_merged:,}，紧凑排版约 ~{zm_merged/2200:,.0f} 页；')
lines.append(f'    即 新整合全书 + 周末必刷解析 ≈ {book_chars + zm_merged:,} 字符 / ~{(book_chars + zm_merged)/2200:,.0f} 页')
lines.append(f'    相对原始全部节省 ≈ {(1-(book_chars + zm_merged)/orig_chars)*100:.1f}% 字符')
lines.append('')
lines.append('注：估算口径 = 原始教辅约1400字/页（同时按 1.67 倍修正图表/留白占位），')
lines.append('    解析版紧凑排版约2200字/页；均为字符估算，实际以 Word 排版为准。')
lines.append('    统计仅含正文段落文本，不含表格内文字与图片。')

with open('compare_print.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines) + '\n')
print('\n'.join(lines))
