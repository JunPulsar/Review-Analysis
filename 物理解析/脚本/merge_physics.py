#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
merge_physics.py — 按《教师目录.docx》顺序，把物理大一轮全部课时合成一个解析小册子。

借鉴《项目开发报告.md》的生物项目经验：
  1) 抽取内核：白名单(答案/解析/提示) + 续段链 + 全/半角通配 + 防知识小标题误判；
  2) 排版：极小上下边距、宋体 10.5pt、行距 1.15、双面外侧 PAGE 域页码；
  3) 目录：序号 + 章节名 + 页码（00 占位，之后由 WPS 回填真实页码）。
物理差异：
  - 无 (×/√) 判断题；以 例N / N. / 思考 / 拓展 / [变式] 为题目锚点；
  - 题目判定用「前瞻结构法」：数字行之后先出现 答案/解析 才算题，
    先出现另一个编号行/考点/课时精练 则视为知识小标题（避免 1.质点 误判成题号）；
  - 一个 docx 可能对应目录里的两行（如“实验二…　实验三…”合并在一篇），内容只渲染一次。

输出：
  物理解析合集_小册子.docx      —— 唯一合册（目录 + 正文）
  物理解析_顺序清单.txt         —— 实际拼装顺序
"""

import os
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

TOC_FILE = '教师目录.docx'
OUT_FILE = '物理解析合集_小册子.docx'
ORDER_FILE = '物理解析_顺序清单.txt'

# ────────────────────────── 排版（生物项目同款） ──────────────────────────
MARGIN_TOP = Cm(0.4)
MARGIN_BOTTOM = Cm(0.4)
MARGIN_LEFT = Cm(1.5)
MARGIN_RIGHT = Cm(1.5)
BODY_FONT = '宋体'
BODY_SIZE = Pt(10.5)
TITLE_SIZE = Pt(13)
H1_SIZE = Pt(14)
TOC_SIZE = Pt(12)
LINE_SPACE = 1.15
PARA_AFTER = Pt(1)
PAGE_NUM_SIZE = Pt(12)


def set_font(run, name=BODY_FONT, size=None, bold=False):
    run.font.name = name
    run.font.bold = bold
    if size:
        run.font.size = size
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), name)


def add_page_number(paragraph, alignment):
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    run = paragraph.add_run()
    set_font(run, BODY_FONT, PAGE_NUM_SIZE)
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    run._element.append(fld_begin)
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    run._element.append(instr)
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run._element.append(fld_end)


def setup_page_numbers(doc):
    section = doc.sections[0]
    sectPr = section._sectPr
    sectPr.set(qn('w:evenAndOddHeaders'), '1')
    footer_odd = section.footer
    footer_odd.is_linked_to_previous = False
    for p in footer_odd.paragraphs:
        p.clear()
    if footer_odd.paragraphs:
        add_page_number(footer_odd.paragraphs[0], WD_ALIGN_PARAGRAPH.RIGHT)
    else:
        add_page_number(footer_odd.add_paragraph(), WD_ALIGN_PARAGRAPH.RIGHT)
    footer_even = section.even_page_footer
    footer_even.is_linked_to_previous = False
    for p in footer_even.paragraphs:
        p.clear()
    if footer_even.paragraphs:
        add_page_number(footer_even.paragraphs[0], WD_ALIGN_PARAGRAPH.LEFT)
    else:
        add_page_number(footer_even.add_paragraph(), WD_ALIGN_PARAGRAPH.LEFT)


def build_page(doc):
    for section in doc.sections:
        section.top_margin = MARGIN_TOP
        section.bottom_margin = MARGIN_BOTTOM
        section.left_margin = MARGIN_LEFT
        section.right_margin = MARGIN_RIGHT
        section.header_distance = Cm(0)
        section.footer_distance = Cm(0.15)
    style = doc.styles['Normal']
    style.font.name = BODY_FONT
    style.font.size = BODY_SIZE
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = PARA_AFTER
    style.paragraph_format.line_spacing = LINE_SPACE
    setup_page_numbers(doc)


def add_heading(doc, text, size, space_before=Pt(8), space_after=Pt(3), center=True):
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = space_before
    p.paragraph_format.space_after = space_after
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    set_font(run, BODY_FONT, size, bold=True)
    return p


def save_doc(doc, path):
    try:
        doc.save(path)
        return path
    except PermissionError:
        alt = os.path.join(os.path.dirname(path),
                           os.path.splitext(os.path.basename(path))[0] + '_新.docx')
        doc.save(alt)
        print(f'  [占用] 已另存为: {alt}')
        return alt


# ────────────────────────── 抽取规则（物理版） ──────────────────────────
RE_ANSWER = re.compile(r'^答案\s*[　 ]')
RE_ANALYSIS = re.compile(r'^解析\s*[　 ]')
RE_PROMPT = re.compile(r'^提示\s*[　 ]')
RE_BARE = re.compile(r'^(答案|解析|提示)[：:]?$')

RE_Q_EX = re.compile(r'^(例\d+)[　 ]')
RE_Q_THK = re.compile(r'^(思考|拓展)[　 ]')
RE_Q_VAR = re.compile(r'^\[变式\d*\][　 ]?')
RE_Q_NUM = re.compile(r'^\d+[.、．](?!\d+(?:\s*[A-Za-z%]|$))')
RE_Q_SUB = re.compile(r'^([（(]\s*\d+\s*[）)])')

RE_SRC = re.compile(r'[（(]\s*\d{4}[^（(）)]{0,25}[）)]')

# 题目强信号：年份 / 分值括号 / 选择空括号
RE_YEAR = re.compile(r'[（(]\s*\d{4}')
RE_SCORE = re.compile(r'[（(]\s*\d+\s*分\s*[）)]')
RE_EMPTY = re.compile(r'[（(]\s*[　 ]*[）)]')
# 题面弱信号（配合前瞻结构法使用）
RE_QWORDS = re.compile(r'下列|正确的是|错误的是|是否正确|是否符合|为什么|能否|多少|如何|试[：:]?|求[：:]|则[：:]|计算[：:]|回答|描述|说明')

# 边界：新题 / 知识小标题 / 章节标题（用于前瞻判定与收块终止）
RE_BOUNDARY = re.compile(
    r'^(例\d+[　 ]|思考[　 ]|拓展[　 ]|\[变式\]|考点[一二三四五六七八九十]|'
    r'课时精练|微点突破|阶段复习|实验[一二三四五六七八九十]+[　 ]|'
    r'第[一二三四五六七八九十]+章|第\d+课时|目标要求|'
    r'分值|\[\d+~?\d*题|\d+[.、．](?!\d))'
)


def is_keeper_line(t):
    return bool(RE_ANSWER.match(t) or RE_ANALYSIS.match(t) or RE_PROMPT.match(t) or RE_BARE.match(t))


def is_num_item(t):
    return bool(RE_Q_NUM.match(t))


def _look_boundary(t):
    """前瞻使用的边界：新例题/思考/拓展/变式/编号行/(1)子题/独立提问句/章节/课时/考点……”"""
    return bool(RE_Q_EX.match(t) or RE_Q_THK.match(t) or RE_Q_VAR.match(t)
                or RE_Q_NUM.match(t) or re.match(r'^[（(]\s*\d+\s*[）)]', t)
                or t.endswith(('？', '?'))
                or re.match(r'^(考点[一二三四五六七八九十]|课时精练|微点突破|阶段复习|'
                            r'实验[一二三四五六七八九十]+[　 ]|第[一二三四五六七八九十]+章|'
                            r'第\d+课时|目标要求|分值|\[\d+~?\d*题)', t))


def _look_has_answer_before_boundary(texts, i):
    j = i + 1
    limit = min(i + 40, len(texts))
    while j < limit:
        nt = texts[j]
        if is_keeper_line(nt):
            return True
        if _look_boundary(nt):
            # (N)…？ 连续子问（如 (1)…关系怎样？(2)…关系怎样？）不算中止边界
            if re.match(r'^[（(]\s*\d+\s*[）)]', nt) and nt.endswith(('？', '?')):
                j += 1
                continue
            return False
        j += 1
    return False


def is_question_start(texts, i, allow_sub=False):
    """前瞻结构法：数字行只有在其后先出现答案/解析（且未遇到新边界）时才判定为题。"""
    t = texts[i]
    if RE_Q_EX.match(t) or RE_Q_THK.match(t) or RE_Q_VAR.match(t):
        return True
    if is_num_item(t):
        if RE_YEAR.search(t) or RE_SCORE.search(t) or RE_EMPTY.search(t):
            return True
        return _look_has_answer_before_boundary(texts, i)
    # (1)/(2) 子题：块内归属于当前例题；顶层孤立时可作独立题（无归属时按孤立答案处理）
    if RE_Q_SUB.match(t):
        if not allow_sub:
            return False
        return _look_has_answer_before_boundary(texts, i)
    # 独立提问句（如“…应如何操作？”）
    if t.endswith(('？', '?')):
        return _look_has_answer_before_boundary(texts, i)
    return False


def collect_qnum(t):
    m = RE_Q_EX.match(t)
    if m:
        return m.group(1)
    m = RE_Q_THK.match(t)
    if m:
        return m.group(1)
    if RE_Q_VAR.match(t):
        return re.match(r'^\[(变式\d*)\]', t).group(1)
    m = re.match(r'^(\d+)[.、．]', t)
    if m:
        return m.group(1) + '.'
    m = RE_Q_SUB.match(t)
    if m:
        return m.group(1)
    if t.endswith(('？', '?')):
        return '问'
    return None


# ── 续段链（生物项目 is_cont_line 同款，去掉判断题逻辑）──
HEADING_RE = re.compile(
    r'^(考向|考点|限时练|阶段|周末必刷|第\d+讲|高考共鸣点|回扣落实|'
    r'【|】|综合提升|巩固|基础|大概念|本讲|课标|目录|练\s*[··]|'
    r'思维|方法|技巧|易错|警示|归纳|总结|拓展|延伸|微网|构建|'
    r'类型|热点|题型|专题|微专题|变式|注意|提醒|'
    r'目标要求|课时精练|分值|考点[一二三四五六七八九十])'
)
NUM_MARK_RE = re.compile(r'^[（(]\s*\d+\s*[）)]|^[①②③④⑤⑥⑦⑧⑨⑩]')
Q_START_RE = re.compile(r'^\d+[.、．](?!\d+(?:\s*[A-Za-z%]|$))')
CN_NUM = {'①': 1, '②': 2, '③': 3, '④': 4, '⑤': 5,
          '⑥': 6, '⑦': 7, '⑧': 8, '⑨': 9, '⑩': 10}
END_SENT = '。？！?）)】;；'


def leading_num(text):
    m = re.match(r'^(?:答案|解析|提示)\s*[　 ]*[（(]?\s*(\d+)\s*[）)]', text)
    if m:
        return int(m.group(1))
    m = re.match(r'^[（(]\s*(\d+)\s*[）)]', text)
    if m:
        return int(m.group(1))
    if text and text[0] in CN_NUM:
        return CN_NUM[text[0]]
    return None


def is_cont_like(prev):
    if RE_ANSWER.match(prev) or RE_ANALYSIS.match(prev) or RE_PROMPT.match(prev) or RE_BARE.match(prev):
        return True
    if HEADING_RE.match(prev):
        return False
    return prev[-1] not in END_SENT


def is_cont_line(prev, cur):
    if not cur:
        return False
    if is_keeper_line(cur) or RE_BOUNDARY.match(cur) or HEADING_RE.match(cur):
        return False
    if Q_START_RE.match(cur):
        return False
    if NUM_MARK_RE.match(cur):
        pn = leading_num(prev)
        cn = leading_num(cur)
        return pn is not None and cn is not None and cn > pn
    if RE_BARE.match(cur):
        return bool(RE_ANSWER.match(prev) or RE_ANALYSIS.match(prev) or RE_PROMPT.match(prev) or RE_BARE.match(prev))
    if RE_BARE.match(prev):
        return True
    if is_cont_like(prev) and prev[-1] not in END_SENT:
        return True
    return False


# ────────────────────────── 抽取主函数 ──────────────────────────
def extract_entries(path):
    """
    返回 [{'num':..., 'src':..., 'lines':[...]}]
    lines 只含 答案/解析/提示 行及其续段。
    """
    doc = Document(path)
    texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    entries = []
    i = 0
    n = len(texts)
    while i < n:
        t = texts[i]
        if is_question_start(texts, i, allow_sub=True):
            qnum = collect_qnum(t)
            m_src = RE_SRC.search(t)
            src = m_src.group(0) if m_src else ''
            j = i + 1
            keeper = []
            while j < n:
                nt = texts[j]
                if is_keeper_line(nt):
                    keeper.append(nt)
                    j += 1
                    # 多段答案/解析/提示：链式续段
                    while j < n and is_cont_line(keeper[-1], texts[j]):
                        keeper.append(texts[j])
                        j += 1
                    continue
                if is_question_start(texts, j):
                    break
                if RE_BOUNDARY.match(nt):
                    break
                j += 1
            if keeper:
                entries.append({'num': qnum, 'src': src, 'lines': keeper})
            i = j
            continue

        # 孤立答案/解析兜底（题号未被识别时的答案不丢）
        if is_keeper_line(t):
            lines = [t]
            j = i + 1
            while j < n and is_cont_line(lines[-1], texts[j]):
                lines.append(texts[j])
                j += 1
            entries.append({'num': '', 'src': '', 'lines': lines})
            i = j
            continue

        i += 1
    return entries


# ────────────────────────── 目录 → 文件映射 ──────────────────────────
def norm(s):
    return re.sub(r'\s+', '', s)


def parse_toc(path):
    doc = Document(path)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    items = []
    for t in paras:
        if re.match(r'^第[一二三四五六七八九十]+章[　 ]', t):
            items.append(('chapter', t))
        else:
            items.append(('section', t))
    return items


def map_sections(items):
    """
    返回 plan: [{'chapter': ...|None, 'sections': [{'title':..., 'file':..., used_in_body': bool}]}]
    """
    files = []
    for f in sorted(os.listdir('.')):
        if not f.endswith('.docx') or f.startswith('~$') or f == TOC_FILE:
            continue
        stem = os.path.splitext(f)[0]
        # 去掉"第X章　"前缀
        no_chap = re.sub(r'^第[一二三四五六七八九十]+章\s*', '', stem)
        files.append({'file': f, 'key': norm(no_chap), 'used': False})

    plan = []
    cur = {'chapter': None, 'sections': []}
    last_file = None
    unmatched = []
    for kind, text in items:
        if kind == 'chapter':
            if cur['sections']:
                plan.append(cur)
            cur = {'chapter': text, 'sections': []}
            last_file = None
            continue
        skey = norm(text)
        # 先试未用文件，再试上一文件（组合课时：同一 docx 对应目录两行）
        cand = [f for f in files if not f['used'] and skey in f['key']]
        if not cand and last_file and skey in last_file['key']:
            cand = [last_file]
        if not cand:
            unmatched.append(text)
            continue
        f = cand[0]
        f['used'] = True
        last_file = f
        cur['sections'].append({'title': text, 'file': f['file'], 'body_rendered': False})
    if cur['sections']:
        plan.append(cur)
    return plan, unmatched


# ────────────────────────── 渲染 ──────────────────────────
def render_entry(doc, entry):
    lines = entry['lines']
    if not lines:
        return
    first = lines[0]
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = LINE_SPACE
    head = (entry['num'] or '') + (entry.get('src') or '')
    if head:
        rn = p.add_run(head + ' ')
    else:
        rn = p.add_run('')
    set_font(rn, BODY_FONT, BODY_SIZE, bold=True)
    r1 = p.add_run(first)
    set_font(r1, BODY_FONT, BODY_SIZE)
    for extra in lines[1:]:
        ep = doc.add_paragraph(extra)
        ep.paragraph_format.space_before = Pt(0)
        ep.paragraph_format.space_after = Pt(0)
        ep.paragraph_format.line_spacing = LINE_SPACE
        for run in ep.runs:
            set_font(run, BODY_FONT, BODY_SIZE)


def render_toc_line(doc, idx, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = LINE_SPACE
    try:
        p.paragraph_format.tab_stops.add_tab_stop(Cm(18.0), WD_TAB_ALIGNMENT.RIGHT)
    except Exception:
        pass
    r1 = p.add_run(f'{idx:02d}. {title}')
    set_font(r1, BODY_FONT, BODY_SIZE)
    r2 = p.add_run('\t')
    set_font(r2, BODY_FONT, BODY_SIZE)
    r3 = p.add_run('00')
    set_font(r3, BODY_FONT, BODY_SIZE)


def main():
    items = parse_toc(TOC_FILE)
    plan, unmatched = map_sections(items)
    if unmatched:
        print('[警告] 目录中未匹配到文件的行:')
        for u in unmatched:
            print('   -', u)

    doc = Document()
    build_page(doc)
    add_heading(doc, '高中物理（步步高 大一轮·广东版）答案解析合集', Pt(16), Pt(0), Pt(8))

    # 同源映射：同一 docx 的重复目录行 → 该 docx 首个渲染标题
    sec_ref = {}
    for ch in plan:
        for sec in ch['sections']:
            fpath = sec['file']
            if fpath not in sec_ref:
                sec_ref[fpath] = sec['title']

    # ── 目录 ──
    add_heading(doc, '目　录', Pt(15), Pt(0), Pt(8))
    idx = 0
    toc_entries = []
    for ch in plan:
        if ch['chapter']:
            add_heading(doc, ch['chapter'], TOC_SIZE, Pt(4), Pt(1), center=False)
        for sec in ch['sections']:
            idx += 1
            render_toc_line(doc, idx, sec['title'])
            toc_entries.append((sec['title'], sec_ref[sec['file']]))
    doc.add_page_break()

    # ── 目录→正文章节页码定位辅助文件（WPS 回填真实页码用）──
    os.makedirs('_toc_tmp', exist_ok=True)
    with open('_toc_tmp/toc_entries.txt', 'w', encoding='utf-8') as fo:
        for disp, head in toc_entries:
            fo.write(f'{disp}\t{head}\n')
    with open('_toc_tmp/titles_physics.txt', 'w', encoding='utf-8') as fo:
        for head in dict.fromkeys(h for _, h in toc_entries):
            fo.write(head + '\n')

    # ── 正文 ──
    order_lines = ['物理解析合集 — 顺序清单', '=' * 60]
    idx = 0
    total_files = 0
    total_entries = 0
    total_ans = 0
    total_ana = 0
    rendered_files = set()
    for ch in plan:
        if ch['chapter']:
            add_heading(doc, ch['chapter'], H1_SIZE, Pt(12), Pt(4))
            order_lines.append('')
            order_lines.append(ch['chapter'])
        for sec in ch['sections']:
            fpath = sec['file']
            idx += 1
            title = sec['title']
            if fpath in rendered_files:
                # 同源合并（如“实验二…　实验三…”同一 docx），内容只出一次
                note = doc.add_paragraph(f'（与“{sec_ref.get(fpath, "上一条")}”同源，内容见上）')
                note.paragraph_format.space_before = Pt(0)
                note.paragraph_format.space_after = Pt(3)
                note.paragraph_format.line_spacing = LINE_SPACE
                for run in note.runs:
                    set_font(run, BODY_FONT, Pt(9))
                order_lines.append(f'  {idx:02d}. {title}  <-  同源合并，见上（{os.path.basename(fpath)}）')
                continue
            sec_ref[fpath] = title
            rendered_files.add(fpath)
            add_heading(doc, title, TITLE_SIZE, Pt(8), Pt(3))
            try:
                entries = extract_entries(fpath)
            except Exception as e:
                print(f'  [失败] {fpath}: {e}')
                entries = []
            if not entries:
                np = doc.add_paragraph('（无条目）')
                np.paragraph_format.space_before = Pt(0)
                np.paragraph_format.space_after = Pt(0)
                continue
            total_files += 1
            total_entries += len(entries)
            for entry in entries:
                total_ans += sum(1 for ln in entry['lines'] if RE_ANSWER.match(ln))
                total_ana += sum(1 for ln in entry['lines'] if RE_ANALYSIS.match(ln))
                render_entry(doc, entry)
            order_lines.append(f'  {idx:02d}. {title}  <-  {os.path.basename(fpath)}')

    saved = save_doc(doc, OUT_FILE)
    with open(ORDER_FILE, 'w', encoding='utf-8') as f:
        f.write('\n'.join(order_lines) + '\n')

    print('=' * 60)
    print(f'输出: {saved}')
    print(f'目录章节块: {len(plan)}，目录小节: {idx}')
    print(f'实际渲染文件: {total_files}（重复目录行已合并渲染）')
    print(f'抽取条目: {total_entries}')
    print(f'答案行: {total_ans}   解析行: {total_ana}')
    print(f'顺序清单: {ORDER_FILE}')


if __name__ == '__main__':
    main()
