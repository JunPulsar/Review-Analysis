#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
stats_physics.py — 直读统计：原始 94 篇的段落/表格/图片/字符数，与小册子字符数对比。
输出 stats_physics.txt（UTF-8）。
"""
import os
import sys
from docx import Document

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

files = sorted(f for f in os.listdir('.')
               if f.endswith('.docx') and f != '教师目录.docx'
               and not f.startswith('物理解析') and not f.endswith('_新.docx'))

sum_para_chars = 0
sum_tbl_chars = 0
sum_tbls = 0
sum_imgs = 0
sum_paras = 0
row_lines = []
for f in files:
    doc = Document(f)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    pc = sum(len(t) for t in paras)
    tbls = doc.tables
    tc = 0
    for tb in tbls:
        for row in tb.rows:
            for cell in row.cells:
                tc += len(cell.text.strip())
    imgc = doc.element.body.xml.count('r:embed')
    sum_para_chars += pc
    sum_tbl_chars += tc
    sum_tbls += len(tbls)
    sum_imgs += imgc
    sum_paras += len(paras)
    row_lines.append(f'{f}|段落{len(paras)}|段字符{pc}|表{len(tbls)}|表字符{tc}|图{imgc}')

# 小册子
doc = Document('物理解析合集_小册子.docx')
body_paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
merged_chars = sum(len(t) for t in body_paras)

with open('stats_physics.txt', 'w', encoding='utf-8') as fo:
    fo.write('=== 原始 94 篇（仅段落文本，不含表格/图片）===\n')
    fo.write(f'文件数: {len(files)}\n')
    fo.write(f'段落数: {sum_paras}   段落字符: {sum_para_chars}\n')
    fo.write(f'表格数: {sum_tbls}   表格字符: {sum_tbl_chars}\n')
    fo.write(f'内嵌图片(r:embed): {sum_imgs}\n')
    fo.write(f'原始总字符(段落+表格): {sum_para_chars + sum_tbl_chars}\n')
    fo.write('\n=== 小册子 ===\n')
    fo.write(f'非空段落: {len(body_paras)}   字符: {merged_chars}\n')
    fo.write(f'字符留存率(段落文本 vs 小册子): {merged_chars / sum_para_chars * 100:.1f}%\n')
    fo.write(f'字符压缩率(表格也算): {merged_chars / (sum_para_chars + sum_tbl_chars) * 100:.1f}%\n')
    fo.write('\n=== 每篇明细 ===\n')
    for line in row_lines:
        fo.write(line + '\n')
print(f'统计完成 -> stats_physics.txt（段落字符 {sum_para_chars}，表字符 {sum_tbl_chars}，图 {sum_imgs}）')
