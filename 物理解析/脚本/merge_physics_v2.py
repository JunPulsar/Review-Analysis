#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
merge_physics_v2.py — 物理解析小册子（规范排版版）

相对 merge_physics.py 的三处升级：
  ① 【核心】渲染改为 XML 级段落复制：原段落的 run / oMath 公式 / 上下标格式原样搬入，
     不再「只取文本、重建 run」——当前版本丢失的 3021 个公式、3548 处上下标由此找回；
     同时保留源文档「数字与字母用 Times New Roman、中文用宋体」的学术排版规范。
  ② 版面可配置：A4 / 双栏 / 左右边距（沿用生物项目实测择优参数）。
  ③ 抽取规则完全复用 merge_physics.py（前瞻结构法、续段链、目录映射），
     判定仍基于 p.text，条目选取结果与旧版一致 —— 只改「怎么画」，不改「选什么」。

用法：
  python merge_physics_v2.py                                  # 默认 A4 + 双栏 + 0.8cm
  python merge_physics_v2.py --out 试验A --cols 1 --margin-lr 1.5 --paper letter
"""

import copy
import os
import re
import sys

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

import merge_physics as MP          # 复用：全部正则、抽取判定、目录映射
import frontmatter as FM            # 封面 / 说明页 / 版本记录（与生物项目共用）

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
COPY_TAGS = {W + 'r', M + 'oMath', M + 'oMathPara', W + 'hyperlink'}

ROOT = os.path.dirname(os.path.abspath(__file__))

# ── 版面参数（默认 = 生物项目实测择优值）──
COLS = 2
PAPER = 'a4'
MARGIN_LR = 0.8
MARGIN_TB = 0.4
OUT_NAME = '物理解析合集_小册子.docx'


def configure(outname=None, cols=None, paper=None, mlr=None, mtb=None):
    global OUT_NAME, COLS, PAPER, MARGIN_LR, MARGIN_TB
    if outname:
        OUT_NAME = outname
    if cols is not None:
        COLS = cols
    if paper:
        PAPER = paper
    if mlr is not None:
        MARGIN_LR = mlr
    if mtb is not None:
        MARGIN_TB = mtb


# ────────────────────────── 版面 ──────────────────────────

def set_columns(doc, num=1, space_cm=0.6):
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
        cols.set(qn('w:space'), str(int(space_cm * 567)))


def text_width_cm(doc):
    s = doc.sections[0]
    return s.page_width.cm - s.left_margin.cm - s.right_margin.cm


def toc_tab_cm(doc):
    w = text_width_cm(doc)
    if COLS > 1:
        return (w - 0.6 * (COLS - 1)) / COLS - 0.35
    return w - 0.6


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
    style.font.name = MP.BODY_FONT
    style.font.size = MP.BODY_SIZE
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = MP.PARA_AFTER
    style.paragraph_format.line_spacing = MP.LINE_SPACE
    # 东亚字体兜底（复制过来的 run 只显式声明了 ascii 字体）
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), MP.BODY_FONT)
    set_columns(doc, COLS)
    # 页脚由 frontmatter 按节管理（封面/说明/目录/版本记录无页码，正文外侧页码）


# ────────────────────────── 抽取（返回源段落对象）──────────────────────────

# 练习区段起点。形态：`课时精练`（80 篇）/ `课时精练[A]`（第71课时，1 篇）→ 合计 81 篇
PRACTICE_RE = re.compile(r'^课时精练')


def split_boundary(paras):
    """返回「课时精练」起始段落索引；无则 None"""
    for i, p in enumerate(paras):
        if PRACTICE_RE.match(p.text.strip()):
            return i
    return None


def extract_entries_rich(path, region='all'):
    """
    与 MP.extract_entries 判定完全一致，额外带回源段落对象：
      返回 [{'num','src','texts':[...],'paras':[Paragraph,...]}]

    region:
      'all'      整篇（旧行为）
      'main'     「课时精练」之前 —— 大本（考点 / 例题）
      'practice' 「课时精练」及之后 —— 小本（81 练）；无精练的篇返回 []
    """
    doc = Document(path)
    paras = [p for p in doc.paragraphs if p.text.strip()]
    if region != 'all':
        cut = split_boundary(paras)
        if cut is None:
            if region == 'practice':
                return []
            # 微点突破 / 阶段复习 无精练 → 整篇归大本
        elif region == 'main':
            paras = paras[:cut]
        else:
            paras = paras[cut:]
    texts = [p.text.strip() for p in paras]
    entries = []
    i, n = 0, len(texts)
    while i < n:
        t = texts[i]
        if MP.is_question_start(texts, i, allow_sub=True):
            qnum = MP.collect_qnum(t)
            m_src = MP.RE_SRC.search(t)
            src = m_src.group(0) if m_src else ''
            j = i + 1
            ki = []
            while j < n:
                nt = texts[j]
                if MP.is_keeper_line(nt):
                    ki.append(j)
                    j += 1
                    while j < n and MP.is_cont_line(texts[ki[-1]], texts[j]):
                        ki.append(j)
                        j += 1
                    continue
                if MP.is_question_start(texts, j):
                    break
                if MP.RE_BOUNDARY.match(nt):
                    break
                j += 1
            if ki:
                entries.append({'num': qnum, 'src': src,
                                'texts': [texts[k] for k in ki],
                                'paras': [paras[k] for k in ki]})
            i = j
            continue
        if MP.is_keeper_line(t):
            ki = [i]
            j = i + 1
            while j < n and MP.is_cont_line(texts[ki[-1]], texts[j]):
                ki.append(j)
                j += 1
            entries.append({'num': '', 'src': '', 'texts': [texts[k] for k in ki],
                            'paras': [paras[k] for k in ki]})
            i = j
            continue
        i += 1
    return entries


# ────────────────────────── 渲染（XML 复制）──────────────────────────

def _new_para(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = MP.LINE_SPACE
    return p


def copy_para_into(dst_p, src_p):
    """把源段落的内容（run + OMML 公式 + 超链接）原样复制到目标段落"""
    for child in src_p._element:
        if child.tag in COPY_TAGS:
            dst_p._element.append(copy.deepcopy(child))


def render_entry(doc, entry):
    ps = entry.get('paras') or []
    if not ps:
        return
    # 首行：题号+来源（加粗） + 源段落内容
    p = _new_para(doc)
    head = (entry['num'] or '') + (entry.get('src') or '')
    if head:
        rn = p.add_run(head + ' ')
        MP.set_font(rn, MP.BODY_FONT, MP.BODY_SIZE, bold=True)
    copy_para_into(p, ps[0])
    # 续段：逐段原样复制
    for sp in ps[1:]:
        copy_para_into(_new_para(doc), sp)


def render_toc_line(doc, idx, title):
    p = _new_para(doc)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    try:
        p.paragraph_format.tab_stops.add_tab_stop(Cm(toc_tab_cm(doc)), WD_TAB_ALIGNMENT.RIGHT)
    except Exception:
        pass
    for txt in (f'{idx:02d}. {title}', '\t', '00'):
        r = p.add_run(txt)
        MP.set_font(r, MP.BODY_FONT, MP.BODY_SIZE)


# ────────────────────────── 主流程 ──────────────────────────

# 小本「练N」的编号：由课时标题派生  第32课时　动量　动量定理 → 练32　动量　动量定理
RE_LESSON = re.compile(r'^第(\d+)课时\s*')


def practice_title(t):
    """大本章节标题 → 小本练标题（非课时篇返回 None）"""
    m = RE_LESSON.match(t)
    return ('练%s　%s' % (m.group(1), t[m.end():].strip())) if m else None


def build_plan():
    items = MP.parse_toc(MP.TOC_FILE)
    plan, unmatched = MP.map_sections(items)
    if unmatched:
        print('[警告] 目录中未匹配到文件的行:')
        for u in unmatched:
            print('   -', u)
    # 每个 docx 只渲染一次（同源重复目录行取首个标题）
    sec_ref, ordered = {}, []
    for ch in plan:
        for sec in ch['sections']:
            if sec['file'] not in sec_ref:
                sec_ref[sec['file']] = sec['title']
                ordered.append((ch['chapter'], sec['title'], sec['file'],
                                practice_title(sec['title'])))
    return plan, ordered, sec_ref


def _toc_block(doc, entries, title):
    """目录：entries = [(显示标题, 用于页码查询的键)]"""
    MP.add_heading(doc, title, Pt(16), Pt(0), Pt(8))
    for i, (disp, _key) in enumerate(entries, 1):
        render_toc_line(doc, i, disp)


def plan_items(ordered):
    """
    返回 (big, small)：
      big   = [(章, 标题, 文件)]            94 章（81 课时 + 8 微点突破 + 5 阶段复习）
      small = [(章, 练标题, 文件)]          81 练（只含带「课时精练」区段的课时）
    两处判据统一在这里，避免 build_book / build_combined / export_titles 各写一份。
    """
    big, small = [], []
    for chap, title, fpath, ptitle in ordered:
        big.append((chap, title, fpath))
        if ptitle is None:
            continue
        if split_boundary([p for p in Document(os.path.join(ROOT, fpath)).paragraphs
                           if p.text.strip()]) is None:
            continue
        small.append((chap, ptitle, fpath))
    return big, small


def build_book(ordered, region, doc_title, toc_title, out_name):
    """region='main' 出大本；'practice' 出小本"""
    big, small = plan_items(ordered)
    items = big if region == 'main' else small

    doc = Document()
    build_page(doc)
    FM.init_cover_section(doc, cols=1, vcenter=True)
    FM.build_cover(doc, '物理')
    FM.new_page_section(doc, cols=1, numbered=False)
    FM.build_colophon(doc, '物理')
    FM.new_page_section(doc, cols=COLS, numbered=False)
    _toc_block(doc, [(t, t) for _c, t, _f in items], toc_title)
    FM.new_page_section(doc, cols=COLS, numbered=True, restart=1)

    n_entries = n_ans = n_ana = 0
    order = ['%s 拼装顺序' % toc_title, '=' * 56]
    cur_chap = None
    for chap, title, fpath in items:
        if chap and chap != cur_chap:
            cur_chap = chap
            MP.add_heading(doc, chap, MP.H1_SIZE, Pt(10), Pt(3))
        MP.add_heading(doc, title, MP.TITLE_SIZE, Pt(8), Pt(3))
        es = extract_entries_rich(os.path.join(ROOT, fpath), region)
        n_entries += len(es)
        for e in es:
            n_ans += sum(1 for x in e['texts'] if MP.RE_ANSWER.match(x))
            n_ana += sum(1 for x in e['texts'] if MP.RE_ANALYSIS.match(x))
            render_entry(doc, e)
        order.append('  %02d. %-44s <- %s' % (len(order), title, os.path.basename(fpath)))

    FM.new_page_section(doc, cols=1, numbered=False)
    FM.build_version_page(doc, '物理', [
        ('大本解析' if region == 'main' else '小本解析',
         '%d %s · %d 条目' % (len(items), '章' if region == 'main' else '练', n_entries))])
    FM.enable_even_odd_headers(doc)
    saved = MP.save_doc(doc, out_name)
    print('%s: %s  章/练 %d  条目 %d  答案 %d  解析 %d'
          % (region, os.path.basename(saved), len(items), n_entries, n_ans, n_ana))
    return items, n_entries, n_ans, n_ana, order


def build_combined(ordered, out_name):
    """合订本：先大本（94 章）后小本（81 练），一本连续页码"""
    big, small = plan_items(ordered)

    doc = Document()
    build_page(doc)
    FM.init_cover_section(doc, cols=1, vcenter=True)
    FM.build_cover(doc, '物理')
    FM.new_page_section(doc, cols=1, numbered=False)
    FM.build_colophon(doc, '物理')
    FM.new_page_section(doc, cols=COLS, numbered=False)
    MP.add_heading(doc, '目　录', Pt(16), Pt(0), Pt(6))
    MP.add_heading(doc, '第一部分　大本解析', MP.TOC_SIZE, Pt(6), Pt(2), center=False)
    i = 0
    for _c, t, _f in big:
        i += 1
        render_toc_line(doc, i, t)
    MP.add_heading(doc, '第二部分　小本解析', MP.TOC_SIZE, Pt(8), Pt(2), center=False)
    for _c, t, _f in small:
        i += 1
        render_toc_line(doc, i, t)
    FM.new_page_section(doc, cols=COLS, numbered=True, restart=1)

    n = 0
    for part, items, region in (('第一部分　大本解析', big, 'main'),
                                ('第二部分　小本解析', small, 'practice')):
        if part.startswith('第二'):
            doc.add_page_break()
        MP.add_heading(doc, part, Pt(16), Pt(0), Pt(8))
        cur = None
        for chap, t, fpath in items:
            if chap and chap != cur:
                cur = chap
                MP.add_heading(doc, chap, MP.H1_SIZE, Pt(10), Pt(3))
            MP.add_heading(doc, t, MP.TITLE_SIZE, Pt(8), Pt(3))
            for e in extract_entries_rich(os.path.join(ROOT, fpath), region):
                render_entry(doc, e)
                n += 1
    FM.new_page_section(doc, cols=1, numbered=False)
    FM.build_version_page(doc, '物理', [
        ('大本解析', '%d 章' % len(big)),
        ('小本解析', '%d 练' % len(small)),
        ('合订本', '%d 章练 · %d 条目' % (len(big) + len(small), n))])
    FM.enable_even_odd_headers(doc)
    saved = MP.save_doc(doc, out_name)
    print('combined: %s  目录 %d 条  条目 %d' % (os.path.basename(saved), i, n))
    return saved


def main():
    _plan, ordered, _ref = build_plan()
    print('版面: %s  %d 栏  左右边距 %.2fcm' % ('A4' if PAPER == 'a4' else 'Letter', COLS, MARGIN_LR))
    print('源文件: %d 个（其中含课时精练的课时 %d 个）'
          % (len(ordered), sum(1 for _c, _t, _f, p in ordered if p)))

    items_b, nb, ab, nb2, order_b = build_book(
        ordered, 'main', '高中物理（步步高 大一轮·广东版）大本解析',
        '目　录', '物理解析_大本.docx')
    items_s, ns, as_, ns2, order_s = build_book(
        ordered, 'practice', '高中物理（步步高 大一轮·广东版）小本解析（81 练）',
        '目　录', '物理解析_小本.docx')
    build_combined(ordered, '物理解析_大本小本合订本.docx')

    with open(MP.ORDER_FILE, 'w', encoding='utf-8') as f:
        f.write('【大本解析】%d 章（%d 条目：答案 %d / 解析 %d）\n' % (len(items_b), nb, ab, nb2))
        f.write('\n'.join(order_b) + '\n\n')
        f.write('【小本解析】%d 练（%d 条目：答案 %d / 解析 %d）\n' % (len(items_s), ns, as_, ns2))
        f.write('\n'.join(order_s) + '\n')
    print('顺序清单:', MP.ORDER_FILE)


if __name__ == '__main__':
    args = sys.argv[1:]

    def _val(flag, default):
        return args[args.index(flag) + 1] if flag in args else default

    configure(cols=int(_val('--cols', '2')),
              paper=_val('--paper', 'a4'),
              mlr=float(_val('--margin-lr', '0.8')),
              mtb=float(_val('--margin-tb', '0.4')))
    main()
