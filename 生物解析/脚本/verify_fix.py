#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_fix.py — 核对修复后的产物是否包含关键内容（只读）"""
from docx import Document


def alltext(p):
    return '\n'.join(x.text for x in Document(p).paragraphs)


checks = [
    ('第31讲 _解析 含(2)神经冲动',
     '大概念七　个体生命活动的调节\\第31讲　神经调节的结构基础和基本方式_解析.docx',
     '(2)神经冲动'),
    ('第22讲 _解析 含Aaa0续段',
     '大概念四　遗传信息控制生物性状并代代相传——遗传规律与人类遗传病\\第22讲　人类遗传病及遗传系谱图分析_解析.docx',
     'Aaa0'),
    ('第25讲 _解析 含裸解析后内容',
     '大概念五　遗传信息主要编码在DNA分子上\\第25讲　基因的表达_解析.docx',
     '由图2中tRNA'),
    ('第32讲 _解析 含解析内容',
     '大概念七　个体生命活动的调节\\第32讲　神经冲动的产生和传导_解析.docx',
     '施加适宜刺激后'),
    ('周末必刷16 合并含水稻苗期',
     '解析合集\\大概念七　个体生命活动的调节\\周末必刷.docx',
     '水稻在苗期'),
    ('高考共鸣点13 合并含酶切阶段',
     '解析合集\\大概念九　利用生物技术与工程生产对人类有用的产品\\高考共鸣点.docx',
     '在酶切阶段'),
    ('第17讲 合并含红果显性',
     '解析合集\\大概念四　遗传信息控制生物性状并代代相传——遗传规律与人类遗传病\\第X讲.docx',
     '红果为显性性状'),
    ('第20讲 合并含甲乙基因型',
     '解析合集\\大概念四　遗传信息控制生物性状并代代相传——遗传规律与人类遗传病\\第X讲.docx',
     '若甲和乙的基因型相同'),
]

res = []
for label, p, kw in checks:
    try:
        ok = kw in alltext(p)
    except Exception as e:
        ok = False
        kw = f'{kw} [读取失败: {e}]'
    res.append(f'{label}: {"OK" if ok else "MISS"}   关键词={kw[:36]}')

with open('verify_fix.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(res) + '\n')
print('\n'.join(res))
