# -*- coding: utf-8 -*-
"""
build_books.py — 重新生成大本解析 / 9 分册 / 小本解析（规范排版 + 正确归类）

与旧链路的区别：
  ① 富文本读取：Word 域（EQ）还原 + 上下标（vertAlign）保留 → 真上下标输出
  ② 区域切分：大本只取「主干区」，练习区段（限时练N/强化练N）剥离出去
  ③ 小本新建：57 段内嵌练习 + 22 篇周末必刷 = 79 篇

输出：
  新整合_全书/解析合集_全书.docx      大本（75 章）
  新整合_分册/<大概念>.docx           9 分册
  新整合_小本解析/小本解析合集.docx    小本（79 篇）
  新整合_全书/顺序清单.txt             大本拼装顺序
  新整合_小本解析/顺序清单.txt         小本拼装顺序
目录页码先占位 00，由 WPS 工具链回填。
"""

import os
import re

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import richtext as RT
import extract_rich as ER
import merge_jiexi as MJ
import merge_book_order as MBO

ROOT = os.path.dirname(os.path.abspath(__file__))
# 所有成册产物统一收在「成册解析」一个文件夹内
OUT_DIR = os.path.join(ROOT, '成册解析')
OUT_BOOK = OUT_DIR
OUT_PARTS = os.path.join(OUT_DIR, '大本解析_分册')
OUT_XB = OUT_DIR

BODY = MJ.BODY_FONT
SIZE = MJ.BODY_SIZE
TITLE_SIZE = MJ.TITLE_SIZE
LINE = MJ.LINE_SPACE

LECT_RE = re.compile(r'^第(\d+)讲')
GXQ_RE = re.compile(r'^高考共鸣点(\d+)')
ZM_RE = re.compile(r'^周末必刷(\d+)')
PRAC_RE = re.compile(r'^(限时练|强化练)\s*(\d+)')

EXTRA = MBO.EXTRA          # 讲号 -> [共鸣点号]

# ── 可配置：分栏数 / 纸张 / 页边距 ──（默认与既有产物一致；试验时用 configure() 切换）
COLS = 1
PAPER = 'letter'           # 'letter'(21.59×27.94) | 'a4'(21.0×29.7)
MARGIN_LR = 1.5            # 左右页边距 cm
MARGIN_TB = 0.4            # 上下页边距 cm
INCLUDE_PARTS = False      # 是否另行输出 9 册分册（默认不再输出）


def configure(outdir, cols=1, paper='letter', margin_lr=1.5, margin_tb=0.4):
    """切换输出目录 / 分栏数 / 纸张 / 页边距"""
    global OUT_DIR, OUT_BOOK, OUT_PARTS, OUT_XB, COLS, PAPER, MARGIN_LR, MARGIN_TB
    OUT_DIR = os.path.join(ROOT, outdir)
    OUT_BOOK = OUT_DIR
    OUT_PARTS = os.path.join(OUT_DIR, '大本解析_分册')
    OUT_XB = OUT_DIR
    COLS = cols
    PAPER = paper
    MARGIN_LR = margin_lr
    MARGIN_TB = margin_tb
    return OUT_DIR


def set_columns(doc, num=1, space_cm=0.6, sep=False):
    """设置分栏：改 sectPr 里的 w:cols（模板已有该元素，就地改属性最安全）"""
    for section in doc.sections:
        sectPr = section._sectPr
        cols = sectPr.find(qn('w:cols'))
        if cols is None:
            cols = OxmlElement('w:cols')
            grid = sectPr.find(qn('w:docGrid'))
            if grid is not None:
                grid.addprevious(cols)
            else:
                sectPr.append(cols)
        cols.set(qn('w:num'), str(num))
        cols.set(qn('w:space'), str(int(space_cm * 567)))   # 1cm = 567 twips
        if sep:
            cols.set(qn('w:sep'), '1')


def text_width_cm(doc):
    """正文可用宽度（cm）"""
    s = doc.sections[0]
    return s.page_width.cm - s.left_margin.cm - s.right_margin.cm


# ─────────────────────────── 版面 ───────────────────────────

def build_page(doc):
    for section in doc.sections:
        if PAPER == 'a4':
            section.page_width = Cm(21.0)
            section.page_height = Cm(29.7)
        section.top_margin = Cm(MARGIN_TB)
        section.bottom_margin = Cm(MARGIN_TB)
        section.left_margin = Cm(MARGIN_LR)
        section.right_margin = Cm(MARGIN_LR)
        section.header_distance = Cm(0)
        section.footer_distance = Cm(0.15)
    style = doc.styles['Normal']
    style.font.name = BODY
    style.font.size = SIZE
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = MJ.PARA_AFTER
    style.paragraph_format.line_spacing = LINE
    set_columns(doc, COLS)
    MJ.setup_page_numbers(doc)


def add_heading(doc, text, size, before=Pt(8), after=Pt(3), center=True, bold=True):
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = before
    p.paragraph_format.space_after = after
    p.paragraph_format.line_spacing = 1.0
    RT.set_font(p.add_run(text), BODY, size, bold=bold)
    return p


def _tight(p):
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = LINE
    return p


# ─────────────────────────── 条目渲染 ───────────────────────────

def render_entry(doc, e):
    """按条目类型渲染；全程使用富文本（真上下标）"""
    if not e.lines:
        return
    if e.num == '' and not e.src:
        for rp in e.lines:
            RT.add_rich(_tight(doc.add_paragraph()), rp.spans, BODY, SIZE)
    else:
        p = _tight(doc.add_paragraph())
        head = (e.num or '') + (e.src or '')
        if head:
            RT.set_font(p.add_run(head + ' '), BODY, SIZE, bold=True)
        RT.add_rich(p, e.lines[0].spans, BODY, SIZE)
        for rp in e.lines[1:]:
            RT.add_rich(_tight(doc.add_paragraph()), rp.spans, BODY, SIZE)


def render_chapter(doc, title, entries):
    """居中章节标题 + 按考点分段"""
    add_heading(doc, title, TITLE_SIZE)
    if not entries:
        _tight(doc.add_paragraph('（无条目）'))
        return
    last = None
    for e in entries:
        if e.katy and e.katy != last:
            add_heading(doc, e.katy, Pt(12), Pt(6), Pt(2))
        last = e.katy
        render_entry(doc, e)


def toc_tab_cm(doc):
    """目录页码右对齐制表位位置（随分栏数变化）"""
    w = text_width_cm(doc)
    if COLS > 1:
        return (w - 0.6 * (COLS - 1)) / COLS - 0.35
    return w - 0.6


def _toc_line(doc, idx, title, width=2):
    """目录行：序号. 标题 \t 00"""
    p = _tight(doc.add_paragraph())
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    try:
        p.paragraph_format.tab_stops.add_tab_stop(Cm(toc_tab_cm(doc)), WD_TAB_ALIGNMENT.RIGHT)
    except Exception:
        pass
    RT.set_font(p.add_run(('%0*d. %s' % (width, idx, title))), BODY, SIZE)
    RT.set_font(p.add_run('\t'), BODY, SIZE)
    RT.set_font(p.add_run('00'), BODY, SIZE)


def render_toc(doc, items, only_day=None):
    """items: [(大概念名, 章节标题)]；页码占位 00"""
    add_heading(doc, '目　录', Pt(16), Pt(0), Pt(8))
    idx = 0
    prev_day = None
    for day, title in items:
        if only_day and day != only_day:
            continue
        if only_day is None and day != prev_day:
            add_heading(doc, day, Pt(12), Pt(6), Pt(2), center=False)
        prev_day = day
        idx += 1
        p = _tight(doc.add_paragraph())
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        try:
            p.paragraph_format.tab_stops.add_tab_stop(Cm(toc_tab_cm(doc)), WD_TAB_ALIGNMENT.RIGHT)
        except Exception:
            pass
        RT.set_font(p.add_run('%02d. %s' % (idx, title)), BODY, SIZE)
        RT.set_font(p.add_run('\t'), BODY, SIZE)
        RT.set_font(p.add_run('00'), BODY, SIZE)


def save_doc(doc, path):
    try:
        doc.save(path)
        return path
    except PermissionError:
        alt = os.path.join(os.path.dirname(path),
                           os.path.splitext(os.path.basename(path))[0] + '_新.docx')
        doc.save(alt)
        print('  [占用] 已另存为:', alt)
        return alt


# ─────────────────────────── 素材收集 ───────────────────────────

def collect_all():
    """返回 (days, prac_items, zm_items)
    days: [{'name','dir','lect':{n:base},'gxq':{n:base},'jc':base}]
    prac_items: [{'name':'限时练3','file':path,'day':大概念名,'sort':(讲号,序)}]
    zm_items: [{'name':'周末必刷1','file':path,'day':...}]
    """
    days = []
    prac, zm = [], []
    for d in sorted(os.listdir(ROOT)):
        dp = os.path.join(ROOT, d)
        if not os.path.isdir(dp) or not d.startswith('大概念'):
            continue
        lect, gxq, jc = {}, {}, None
        for f in sorted(os.listdir(dp)):
            if not f.endswith('.docx') or f.startswith(chr(126)) or f.endswith('_解析.docx'):
                continue
            base = f[:-5]
            p = os.path.join(dp, f)
            m = LECT_RE.match(base)
            if m:
                lect[int(m.group(1))] = base
                continue
            m = GXQ_RE.match(base)
            if m:
                gxq[int(m.group(1))] = base
                continue
            if base.startswith('阶段排查'):
                jc = base
                continue
            m = ZM_RE.match(base)
            if m:
                zm.append({'name': re.sub(r'\s+', '', m.group(0)), 'file': p, 'day': d,
                           'num': int(m.group(1))})
        # 该大概念内嵌的练习
        for base in list(lect.values()) + list(gxq.values()):
            fp = os.path.join(dp, base + '.docx')
            rp = ER.normalize_split_judges_rich(ER.load_rich(fp))
            _main, pracrp = ER.split_practice(rp)
            pn = ER.practice_name(pracrp)
            if not pn:
                continue
            m = PRAC_RE.match(pn)
            kind, num = m.group(1), int(m.group(2))
            if kind == '限时练':
                sort = (num, 0)
            else:  # 强化练 → 归属到它跟随的那一讲的序号
                host = next((k for k, v in gxq.items() if v == base), None)
                lect_no = next((L for L, gs in EXTRA.items() if host in gs), 999)
                sort = (lect_no, 1)
            prac.append({'name': pn, 'file': fp, 'day': d, 'sort': sort, 'kind': kind})
        days.append({'name': d, 'dir': dp, 'lect': lect, 'gxq': gxq, 'jc': jc})
    days.sort(key=lambda x: min(x['lect']) if x['lect'] else 999)
    prac.sort(key=lambda x: (x['day'], x['sort']))
    zm.sort(key=lambda x: (x['day'], x['num']))
    return days, prac, zm


def entries_of_chapter(path, main_only=True):
    rp = ER.normalize_split_judges_rich(ER.load_rich(path))
    if main_only:
        rp, _ = ER.split_practice(rp)
    return ER.extract_entries(rp)


def entries_of_practice(prac):
    rp = ER.normalize_split_judges_rich(ER.load_rich(prac['file']))
    _main, pracrp = ER.split_practice(rp)
    return ER.extract_entries(pracrp)


def entries_of_zm(zm):
    rp = ER.normalize_split_judges_rich(ER.load_rich(zm['file']))
    return ER.extract_entries(rp)


# ─────────────────────────── 主流程 ───────────────────────────

def main():
    os.makedirs(OUT_BOOK, exist_ok=True)
    os.makedirs(OUT_XB, exist_ok=True)
    if INCLUDE_PARTS:
        os.makedirs(OUT_PARTS, exist_ok=True)

    days, prac, zm = collect_all()

    # 大本 plan
    plan = []
    for d in days:
        for n in sorted(d['lect']):
            plan.append((d['name'], 'lect', n))
            for g in EXTRA.get(n, []):
                if g in d['gxq']:
                    plan.append((d['name'], 'gxq', g))
        if d['jc']:
            plan.append((d['name'], 'jc', None))

    def title_of(day, kind, num):
        d = next(x for x in days if x['name'] == day)
        if kind == 'lect':
            return d['lect'][num]
        if kind == 'gxq':
            return d['gxq'][num]
        return d['jc']

    def path_of(day, kind, num):
        d = next(x for x in days if x['name'] == day)
        return os.path.join(d['dir'], title_of(day, kind, num) + '.docx')

    toc_items = [(day, title_of(day, k, n)) for day, k, n in plan]

    # ── 大本 ──
    book = Document()
    build_page(book)
    add_heading(book, '高中生物 解析合集（全书）', Pt(16), Pt(0), Pt(8))
    render_toc(book, toc_items)
    book.add_page_break()
    cur_day = None
    n_entry = 0
    for day, kind, num in plan:
        if day != cur_day:
            cur_day = day
            add_heading(book, day, Pt(14), Pt(12), Pt(4))
        es = entries_of_chapter(path_of(day, kind, num))
        n_entry += len(es)
        render_chapter(book, title_of(day, kind, num), es)
    p_book = save_doc(book, os.path.join(OUT_BOOK, '大本解析_全书.docx'))
    print('大本:', p_book, ' 章节', len(plan), ' 条目', n_entry)

    # ── 9 分册（默认不再输出，需要时加 --with-parts）──
    if INCLUDE_PARTS:
        for d in days:
            part = Document()
            build_page(part)
            add_heading(part, '%s —— 解析合集' % d['name'], Pt(14), Pt(0), Pt(6))
            render_toc(part, toc_items, only_day=d['name'])
            part.add_page_break()
            add_heading(part, '%s —— 解析合集' % d['name'], Pt(14), Pt(8), Pt(4))
            for day, kind, num in plan:
                if day != d['name']:
                    continue
                render_chapter(part, title_of(day, kind, num), entries_of_chapter(path_of(day, kind, num)))
            save_doc(part, os.path.join(OUT_PARTS, '%s.docx' % d['name']))
        print('分册: 9 册 ->', OUT_PARTS)

    # ── 小本 ──
    xb_items = []
    for day in [d['name'] for d in days]:
        for it in [x for x in prac if x['day'] == day]:
            xb_items.append((day, it['name'], it['file'], 'prac'))
        for it in [x for x in zm if x['day'] == day]:
            xb_items.append((day, it['name'], it['file'], 'zm'))

    xb = Document()
    build_page(xb)
    add_heading(xb, '高中生物 小本解析（限时练 · 强化练 · 周末必刷）', Pt(16), Pt(0), Pt(4))
    add_heading(xb, '目　录', Pt(16), Pt(0), Pt(8))
    for i, (day, name, _f, kind) in enumerate(xb_items, 1):
        p = _tight(xb.add_paragraph())
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        try:
            p.paragraph_format.tab_stops.add_tab_stop(Cm(toc_tab_cm(xb)), WD_TAB_ALIGNMENT.RIGHT)
        except Exception:
            pass
        RT.set_font(p.add_run('%02d. %s' % (i, name)), BODY, SIZE)
        RT.set_font(p.add_run('\t'), BODY, SIZE)
        RT.set_font(p.add_run('00'), BODY, SIZE)
    xb.add_page_break()

    n_xb = 0
    cur_day = None
    for day, name, f, kind in xb_items:
        if day != cur_day:
            cur_day = day
            add_heading(xb, day, Pt(14), Pt(12), Pt(4))
        rp = ER.normalize_split_judges_rich(ER.load_rich(f))
        if kind == 'prac':
            _main, rp = ER.split_practice(rp)
        es = ER.extract_entries(rp)
        n_xb += len(es)
        render_chapter(xb, name, es)
    p_xb = save_doc(xb, os.path.join(OUT_XB, '小本解析_合集.docx'))
    print('小本:', p_xb, ' 篇数', len(xb_items), ' 条目', n_xb)

    # ── 合订本：先大本、后小本（供胶装打印）──
    cb = Document()
    build_page(cb)
    add_heading(cb, '高中生物 解析合集', Pt(18), Pt(0), Pt(4))
    add_heading(cb, '第一部分　大本解析（%d 章）　　第二部分　小本解析（%d 篇）'
                % (len(plan), len(xb_items)), Pt(10.5), Pt(0), Pt(10))
    add_heading(cb, '目　录', Pt(16), Pt(0), Pt(6))
    add_heading(cb, '第一部分　大本解析', Pt(12), Pt(6), Pt(2), center=False)
    idx = 0
    for day, kind, num in plan:
        idx += 1
        _toc_line(cb, idx, title_of(day, kind, num), width=3)
    add_heading(cb, '第二部分　小本解析', Pt(12), Pt(8), Pt(2), center=False)
    for day, name, _f, _k in xb_items:
        idx += 1
        _toc_line(cb, idx, name, width=3)
    cb.add_page_break()

    add_heading(cb, '第一部分　大本解析', Pt(16), Pt(0), Pt(8))
    cur_day = None
    for day, kind, num in plan:
        if day != cur_day:
            cur_day = day
            add_heading(cb, day, Pt(14), Pt(12), Pt(4))
        render_chapter(cb, title_of(day, kind, num), entries_of_chapter(path_of(day, kind, num)))

    cb.add_page_break()
    add_heading(cb, '第二部分　小本解析', Pt(16), Pt(0), Pt(8))
    cur_day = None
    for day, name, f, kind in xb_items:
        if day != cur_day:
            cur_day = day
            add_heading(cb, day, Pt(14), Pt(12), Pt(4))
        rp = ER.normalize_split_judges_rich(ER.load_rich(f))
        if kind == 'prac':
            _m, rp = ER.split_practice(rp)
        render_chapter(cb, name, ER.extract_entries(rp))
    p_cb = save_doc(cb, os.path.join(OUT_DIR, '大本小本合订本.docx'))
    print('合订本:', p_cb, ' 共', len(plan) + len(xb_items), '章/篇 目录条目', idx)

    # ── 顺序清单（写到 _toc_tmp，保持交付目录只有 3 个 docx）──
    tmp = os.path.join(ROOT, '_toc_tmp')
    os.makedirs(tmp, exist_ok=True)
    with open(os.path.join(tmp, '顺序清单_大本.txt'), 'w', encoding='utf-8') as f:
        f.write('大本解析 拼装顺序（%d 章）\n%s\n' % (len(plan), '=' * 60))
        for i, (day, k, n) in enumerate(plan, 1):
            f.write('%02d. %s\n' % (i, title_of(day, k, n)))
    with open(os.path.join(tmp, '顺序清单_小本.txt'), 'w', encoding='utf-8') as f:
        f.write('小本解析 拼装顺序（%d 篇）\n%s\n' % (len(xb_items), '=' * 60))
        for i, (day, name, _f, kind) in enumerate(xb_items, 1):
            f.write('%02d. %-12s [%s]\n' % (i, name, day))

    print()
    print('=' * 70)
    print('大本章节数 :', len(plan), '（应为 52讲+14共鸣点+9阶段排查 = 75）')
    print('大本条目数 :', n_entry)
    print('小本篇数   :', len(xb_items), '（应为 52限时练+5强化练+22周末必刷 = 79）')
    print('小本条目数 :', n_xb)
    print('练习条目合计(应=486):', sum(len(entries_of_practice(x)) for x in prac))
    print('=' * 70)


if __name__ == '__main__':
    import sys
    _args = sys.argv[1:]

    def _val(flag, default):
        return _args[_args.index(flag) + 1] if flag in _args else default

    if any(f in _args for f in ('--out', '--cols', '--paper', '--margin-lr', '--margin-tb', '--with-parts')):
        configure(_val('--out', '成册解析'),
                  int(_val('--cols', '1')),
                  _val('--paper', 'letter'),
                  float(_val('--margin-lr', '1.5')),
                  float(_val('--margin-tb', '0.4')))
        if '--with-parts' in _args:
            INCLUDE_PARTS = True
        print('配置: 输出=%s  分栏=%d  纸张=%s  左右边距=%.2fcm  上下边距=%.2fcm  输出分册=%s'
              % (OUT_DIR, COLS, PAPER, MARGIN_LR, MARGIN_TB, INCLUDE_PARTS))
    main()
