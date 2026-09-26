#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""render_pages.py — 把 PDF 预览渲染成 PNG（供人工视觉复核）。"""
import fitz

doc = fitz.open('物理解析合集_小册子_预览.pdf')
print(f'PDF 页数: {doc.page_count}')
for i in [0, 1, 2, 39, 139]:
    if i < doc.page_count:
        page = doc[i]
        pix = page.get_pixmap(dpi=110)
        out = f'render_page_{i + 1:03d}.png'
        pix.save(out)
        print('已渲染', out)
