#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_titles_physics.py — 生成 WPS 页码定位清单（三本：大本 / 小本 / 合订本）

  _toc_tmp/titles_<key>.txt      —— 正文实际渲染的章节标题，每本一份
                                    wps_pagemap_physics.ps1 用它逐段定位真实页码
  _toc_tmp/toc_entries_<key>.txt —— 目录行「显示名 <TAB> 页码查询名」
                                    fill_toc_pages_physics.py 用它回填 00 占位

key: book=大本  xb=小本  cb=合订本

三本标题统一由 merge_physics_v2.plan_items 派生（与正文同源，不会漂移）。
合订本 = 大本 94 标题 + 小本 81 练标题；小本用「练N」命名，与大本标题天然不重名，
所以 WPS 逐段定位不会串页。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.chdir(os.path.dirname(os.path.abspath(__file__)))

import merge_physics_v2 as MV

TMP = '_toc_tmp'
os.makedirs(TMP, exist_ok=True)


def main():
    _plan, ordered, _ref = MV.build_plan()
    big, small = MV.plan_items(ordered)

    books = {
        'book': [(t, t) for _c, t, _f in big],
        'xb': [(t, t) for _c, t, _f in small],
        'cb': [(t, t) for _c, t, _f in big] + [(t, t) for _c, t, _f in small],
    }
    for key, entries in books.items():
        with open(os.path.join(TMP, 'titles_%s.txt' % key), 'w', encoding='utf-8') as f:
            f.write('\n'.join(head for _d, head in entries) + '\n')
        with open(os.path.join(TMP, 'toc_entries_%s.txt' % key), 'w', encoding='utf-8') as f:
            for disp, head in entries:
                f.write('%s\t%s\n' % (disp, head))
        print('%s: 标题 %d  目录行 %d' % (key, len(entries), len(entries)))


if __name__ == '__main__':
    main()
