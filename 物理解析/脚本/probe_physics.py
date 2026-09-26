#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_physics.py — 探明物理大一轮 docx 的真实结构（为抽取答案/解析做准备）

输出:
  - 教师目录.docx 的全部段落（决定合并顺序）
  - 每类样例文档: 段落计数 / 表格数 / 图片数 / 答案-解析-提示 命中行 / 判断题样式
  - 样例文档前 150 个非空段落全文（人工观察格式）
全部写入 probe_physics.txt（UTF-8）。
"""

import os
import re
import sys
from docx import Document

# 控制台强制 UTF-8，避免 GBK 打印失败
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

OUT = 'probe_physics.txt'
lines = []
OUT_F = open(OUT, 'w', encoding='utf-8')


def log(s=''):
    lines.append(str(s))
    OUT_F.write(str(s) + '\n')
    OUT_F.flush()
    try:
        print(s)
    except Exception:
        pass


def fmt_label(label):
    return f'\n{"=" * 70}\n### {label}\n{"=" * 70}\n'


def dump_doc_header(label, path):
    try:
        doc = Document(path)
    except Exception as e:
        log(f'[打开失败] {path}: {e}')
        return None
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tables = doc.tables
    img_count = doc.element.body.xml.count('r:embed')
    log(fmt_label(f'{label}  [{os.path.basename(path)}]'))
    log(f'文件: {path}')
    log(f'非空段落: {len(paras)}   表格: {len(tables)}   内嵌图片(r:embed): {img_count}')
    # 关键行命中统计
    pat_answer = re.compile(r'^答案\s*[　 ]')
    pat_analysis = re.compile(r'^解析\s*[　 ]')
    pat_prompt = re.compile(r'^提示\s*[　 ]')
    pat_tf = re.compile(r'^[（(]\s*\d+\s*[）)]')
    pat_tf_mark = re.compile(r'[（(][×√][）)]')
    pat_src = re.compile(r'[（(]\s*\d{4}')
    hits = {
        '答案(空格)': [], '解析(空格)': [], '提示(空格)': [],
        '判断题行(数字+×√)': [], '含年份行': [],
    }
    for t in paras:
        if pat_answer.match(t):
            hits['答案(空格)'].append(t[:60])
        if pat_analysis.match(t):
            hits['解析(空格)'].append(t[:60])
        if pat_prompt.match(t):
            hits['提示(空格)'].append(t[:60])
        if pat_tf.match(t) and pat_tf_mark.search(t):
            hits['判断题行(数字+×√)'].append(t[:60])
        if pat_src.search(t):
            hits['含年份行'].append(t[:60])
    for k, v in hits.items():
        log(f'  {k}: {len(v)}')
        for it in v[:8]:
            log(f'    | {it}')
        if len(v) > 8:
            log(f'    ... 还有 {len(v) - 8} 行（详见 count 已足够）')
    # 表格文本抽样（前 2 张表的前 8 个单元格）
    if tables:
        log('--- 表格文本抽样（前2张表） ---')
        for ti, tb in enumerate(tables[:2]):
            log(f'  表{ti + 1}: {len(tb.rows)} 行 x {len(tb.columns)} 列')
            shown = 0
            for row in tb.rows:
                for cell in row.cells:
                    ct = cell.text.strip()
                    if ct:
                        log(f'    [{ct[:70]}]')
                        shown += 1
                        if shown >= 10:
                            break
                if shown >= 10:
                    break
    return paras


def main():
    log('*' * 70)
    log('PROBE 物理大一轮 docx（教师用书Word版文档）')
    log('*' * 70)
    log(f'CWD: {os.getcwd()}')
    log(f'文件数(docx): {sum(1 for f in os.listdir(".") if f.endswith(".docx"))}')

    # 1) 教师目录
    dump_doc_header('教师目录（合并顺序依据）', '教师目录.docx')

    # 2) 各类样例
    samples = [
        '第一章　第1课时　运动的描述 .docx',
        '第一章　第5课时　实验一：研究匀变速直线运动.docx',
        '第一章　微点突破1　追及相遇问题.docx',
        '第三章　阶段复习(一)　力与直线运动.docx',
        '第十一章　第55课时　磁场及其对电流的作用.docx',
        '第十五章　第75课时　实验十九：用油膜法估测油酸分子的大小　实验二十：探究等温情况下一定质量气体压强与体积的关系.docx',
    ]
    for s in samples:
        if os.path.exists(s):
            dump_doc_header('样例', s)
        else:
            log(f'[样例不存在] {s}')

    # 3) 一个完整文档的段落流（前 150 个非空段落）
    sample = '第一章　第1课时　运动的描述 .docx'
    if os.path.exists(sample):
        log(fmt_label(f'段落流全文（前150段）: {sample}'))
        doc = Document(sample)
        paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for i, t in enumerate(paras[:150]):
            log(f'[{i:03d}] {t}')

    try:
        print('\n[完成] 已写入 ' + OUT)
    except Exception:
        pass
    OUT_F.close()


if __name__ == '__main__':
    main()
