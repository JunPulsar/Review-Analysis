#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
merge_book_order.py — 按新排序整合解析合集（只读原始 docx，生成新文件，不改原文件）。

排序规则（用户确认）：
- 第X讲 1→52 连续；
- 高考共鸣点插入：1→第1讲后，2→第2讲后，3→第5讲后，4→第7讲后，5→第8讲后，
  6→第9讲后，7→第12讲后，8→第14讲后，9→第15讲后，10→第21讲后，11→第29讲后，
  12/13/14→第51讲后；
- 阶段排查放在所在大概念最后；
- 周末必刷不参与。

文档结构（用户要求）：
  1) 最前面给全书目录（大概念分组 + 章节列表）；
  2) 目录后正文：居中章节标题（第1讲/高考共鸣点/阶段排查）→ 该章答案解析整体连续输出
     → 下一章居中标题 → …（每章只在开头标一次标题，不做逐条重复标注）。

输出：
  新整合_全书/解析合集_全书.docx          —— 1 册（目录 + 正文）
  新整合_分册/大概念… .docx              —— 9 册（各自目录 + 正文）
  新整合_全书/顺序清单.txt                —— 实际拼装顺序
"""

import os
import re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

import merge_jiexi as MJ

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_BOOK = os.path.join(ROOT, '新整合_全书')
OUT_PARTS = os.path.join(ROOT, '新整合_分册')

# 高考共鸣点插入映射：讲号 -> [共鸣点号]
EXTRA = {
    1: [1],
    2: [2],
    5: [3],
    7: [4],
    8: [5],
    9: [6],
    12: [7],
    14: [8],
    15: [9],
    21: [10],
    29: [11],
    51: [12, 13, 14],
}

LECT_RE = re.compile(r'^第(\d+)讲')
GXQ_RE = re.compile(r'^高考共鸣点(\d+)')
KATY_RE = re.compile(r'^考点[一二三四五六七八九十\d]+\s*[　 :：]')


def clean_kind(title):
    """返回 (类别, 编号或None)：'lect'/'gxq'/'jc'"""
    m = LECT_RE.match(title)
    if m:
        return 'lect', int(m.group(1))
    m = GXQ_RE.match(title)
    if m:
        return 'gxq', int(m.group(1))
    if title.startswith('阶段排查'):
        return 'jc', None
    return None, None


def collect_daginian():
    """返回 [{'name':..., 'dir':..., 'lect': {num: basename}, 'gxq': {...}, 'jc': basename}]"""
    result = []
    days = sorted(d for d in os.listdir(ROOT) if os.path.isdir(os.path.join(ROOT, d)) and d.startswith('大概念'))
    for d in days:
        dp = os.path.join(ROOT, d)
        lect, gxq, jc = {}, {}, None
        for f in os.listdir(dp):
            if not f.endswith('.docx') or f.startswith('~$') or f.endswith('_解析.docx'):
                continue
            base = f[:-5]
            kind, num = clean_kind(base)
            if kind == 'lect':
                lect[num] = base
            elif kind == 'gxq':
                gxq[num] = base
            elif kind == 'jc':
                jc = base
        result.append({'name': d, 'dir': dp, 'lect': lect, 'gxq': gxq, 'jc': jc})
    # 大概念按最小讲号排序（一1-7、二8-12、三13-15、四16-22、五23-26、六27-29、七30-39、八40-45、九46-52）
    result.sort(key=lambda x: min(x['lect']) if x['lect'] else 999)
    return result


def build_page(doc):
    """页面设置：极小上下边距 + 页脚外侧页码（奇右偶左），减少留白与打印页数"""
    for section in doc.sections:
        section.top_margin = Cm(0.4)
        section.bottom_margin = Cm(0.4)
        section.left_margin = MJ.MARGIN_LEFT
        section.right_margin = MJ.MARGIN_RIGHT
        section.header_distance = Cm(0)
        section.footer_distance = Cm(0.15)
    style = doc.styles['Normal']
    style.font.name = MJ.BODY_FONT
    style.font.size = MJ.BODY_SIZE
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = MJ.PARA_AFTER
    style.paragraph_format.line_spacing = MJ.LINE_SPACE
    # 页脚外侧页码：奇数页右下、偶数页左下（镜像页脚）
    MJ.setup_page_numbers(doc)


def add_heading(doc, text, size, space_before=Pt(8), space_after=Pt(3), center=True):
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = space_before
    p.paragraph_format.space_after = space_after
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    MJ.set_font(run, MJ.BODY_FONT, size, bold=True)
    return p


def render_entry(doc, entry):
    """按解析合集同款排版输出一个条目（考点速览块 / 典型例题）"""
    lines = entry['lines']
    if not lines:
        return
    if entry['num'] == '' and not entry.get('src'):
        for line in lines:
            ep = doc.add_paragraph(line)
            ep.paragraph_format.space_before = Pt(0)
            ep.paragraph_format.space_after = Pt(0)
            ep.paragraph_format.line_spacing = MJ.LINE_SPACE
            for run in ep.runs:
                MJ.set_font(run, MJ.BODY_FONT, MJ.BODY_SIZE)
    else:
        # 典型例题：题号+来源 与首行内容同行（如 6.(2025·浙江6月卷，18) 答案　A）
        first = lines[0]
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = MJ.LINE_SPACE
        head = (entry['num'] or '') + (entry.get('src') or '')
        if head:
            rn = p.add_run(head + ' ')
        else:
            rn = p.add_run('')
        MJ.set_font(rn, MJ.BODY_FONT, MJ.BODY_SIZE, bold=True)
        r1 = p.add_run(first)
        MJ.set_font(r1, MJ.BODY_FONT, MJ.BODY_SIZE)
        for extra in lines[1:]:
            ep = doc.add_paragraph(extra)
            ep.paragraph_format.space_before = Pt(0)
            ep.paragraph_format.space_after = Pt(0)
            ep.paragraph_format.line_spacing = MJ.LINE_SPACE
            for run in ep.runs:
                MJ.set_font(run, MJ.BODY_FONT, MJ.BODY_SIZE)


def extract_entries_katy(input_path):
    """
    与 merge_jiexi.process_one_docx 相同逻辑提取条目，但额外跟踪"考点N　…"标题，
    每个条目记录其所属考点标题（无考点则为 None）。
    返回: [(katy_title|None, {'num':..., 'lines':[...]})]
    """
    doc = Document(input_path)
    texts = MJ.normalize_split_judges([p.text.strip() for p in doc.paragraphs if p.text.strip()])

    entries = []
    cur_katy = None
    i = 0
    n = len(texts)

    while i < n:
        t = texts[i]

        # 考点标题（分段标记）
        if KATY_RE.match(t):
            cur_katy = t
            i += 1
            continue

        # ── 考点速览区域 ──
        if MJ.PTN_TF_Q.match(t) and MJ.PTN_TF_ANS.search(t):
            block = [t]
            i += 1
            while i < n:
                nt = texts[i]
                if KATY_RE.match(nt):
                    break  # 考点标题：块边界
                if MJ.PTN_PROMPT.match(nt) or MJ.PTN_ANSWER.match(nt) or MJ.PTN_ANALYSIS.match(nt):
                    block.append(nt)
                    i += 1
                    while i < n and MJ.is_cont_line(block[-1], texts[i]):
                        block.append(texts[i])
                        i += 1
                else:
                    break
            entries.append((cur_katy, {'num': '', 'lines': block}))
            continue

        # ── 主考题区域（含无年份题、全角年份题、无编号题）──
        qnum = MJ.extract_qnum(t)
        if qnum is not None:
            m_src = MJ.RE_SRC.search(t)
            src = m_src.group(0) if m_src else ''
            i += 1
            keeper_lines = []
            while i < n:
                nt = texts[i]
                if MJ.PTN_ANSWER.match(nt) or MJ.PTN_ANALYSIS.match(nt):
                    keeper_lines.append(nt)
                    i += 1
                    while i < n and MJ.is_cont_line(keeper_lines[-1], texts[i]):
                        keeper_lines.append(texts[i])
                        i += 1
                elif MJ.PTN_BARE_LABEL.match(nt) and keeper_lines:
                    keeper_lines.append(nt)
                    i += 1
                    while i < n and MJ.is_cont_line(keeper_lines[-1], texts[i]):
                        keeper_lines.append(texts[i])
                        i += 1
                elif MJ.PTN_PROMPT.match(nt):
                    i += 1
                elif KATY_RE.match(nt):
                    break  # 考点标题：题目扫描边界
                elif MJ.extract_qnum(nt) is not None or (MJ.PTN_TF_Q.match(nt) and MJ.PTN_TF_ANS.search(nt)):
                    break
                else:
                    i += 1
            if keeper_lines:
                entries.append((cur_katy, {'num': qnum, 'src': src, 'lines': keeper_lines}))
            continue

        # 孤立答案/解析兜底：没有题目编号跟随的答案/解析（如连线题答案）保留为无编号条目
        if MJ.PTN_ANSWER.match(t) or MJ.PTN_ANALYSIS.match(t):
            lns = [t]
            i += 1
            while i < n and MJ.is_cont_line(lns[-1], texts[i]):
                lns.append(texts[i])
                i += 1
            entries.append((cur_katy, {'num': '', 'src': '', 'lines': lns}))
            continue

        i += 1

    return entries


def render_chapter(doc, title, input_path):
    """输出一个章节：居中章节标题一次 + 按考点分段的答案解析"""
    add_heading(doc, title, MJ.TITLE_SIZE)
    try:
        entries = extract_entries_katy(input_path)
    except Exception as e:
        entries = []
        print(f'  读取失败 {title}: {e}')
    if not entries:
        np = doc.add_paragraph('（无条目）')
        np.paragraph_format.space_before = Pt(0)
        np.paragraph_format.space_after = Pt(0)
        return
    last_katy = None
    for katy, entry in entries:
        if katy and katy != last_katy:
            add_heading(doc, katy, Pt(12), Pt(6), Pt(2))
        last_katy = katy
        render_entry(doc, entry)


def render_toc(doc, days, plan, title_of, only_day=None):
    """目录：大概念分组 + 章节列表（序号. 章节名 页码），页码先以 00 占位，稍后由 fill_toc_pages 回填"""
    add_heading(doc, '目　录', Pt(16), Pt(0), Pt(8))
    idx = 0
    for i, (day, kind, num) in enumerate(plan):
        if only_day and day != only_day:
            continue
        if only_day is None:
            # 全书：大概念分组标题
            if i == 0 or plan[i - 1][0] != day:
                add_heading(doc, day, Pt(12), Pt(6), Pt(2), center=False)
        idx += 1
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = MJ.LINE_SPACE
        try:
            p.paragraph_format.tab_stops.add_tab_stop(Cm(18.0), WD_TAB_ALIGNMENT.RIGHT)
        except Exception:
            pass
        r1 = p.add_run(f'{idx:02d}. {title_of(day, kind, num)}')
        MJ.set_font(r1, MJ.BODY_FONT, MJ.BODY_SIZE)
        r2 = p.add_run('\t')
        MJ.set_font(r2, MJ.BODY_FONT, MJ.BODY_SIZE)
        r3 = p.add_run('00')
        MJ.set_font(r3, MJ.BODY_FONT, MJ.BODY_SIZE)


def save_doc(doc, path):
    """保存；文件被占用（如 Word 打开中）时自动另存 _新.docx 并提示"""
    try:
        doc.save(path)
        return path
    except PermissionError:
        alt = os.path.join(os.path.dirname(path),
                           os.path.splitext(os.path.basename(path))[0] + '_新.docx')
        doc.save(alt)
        print(f'  [占用] 已另存为: {alt}（原文件 {os.path.basename(path)} 可能正被打开）')
        return alt


def main():
    os.makedirs(OUT_BOOK, exist_ok=True)
    os.makedirs(OUT_PARTS, exist_ok=True)

    days = collect_daginian()

    # 构建顺序 plan: (daginian_name, kind, num)
    plan = []
    for d in days:
        for n in sorted(d['lect']):
            plan.append((d['name'], 'lect', n))
            for g in EXTRA.get(n, []):
                plan.append((d['name'], 'gxq', g))
        if d['jc']:
            plan.append((d['name'], 'jc', None))

    def day_of(day_name):
        return next(x for x in days if x['name'] == day_name)

    def title_of(day, kind, num):
        d = day_of(day)
        if kind == 'lect':
            return d['lect'][num]
        if kind == 'gxq':
            return d['gxq'][num]
        return d['jc']

    def path_of(day, kind, num):
        d = day_of(day)
        return os.path.join(d['dir'], title_of(day, kind, num) + '.docx')

    # ── 全书：目录 + 正文 ──
    book = Document()
    build_page(book)
    add_heading(book, '高中生物 解析合集（全书）', Pt(16), Pt(0), Pt(8))
    render_toc(book, days, plan, title_of)
    book.add_page_break()
    cur_day = None
    for day, kind, num in plan:
        if day != cur_day:
            cur_day = day
            add_heading(book, day, Pt(14), Pt(12), Pt(4))
        render_chapter(book, title_of(day, kind, num), path_of(day, kind, num))
    save_doc(book, os.path.join(OUT_BOOK, '解析合集_全书.docx'))

    # ── 9 册分册：各自目录 + 正文 ──
    for d in days:
        part = Document()
        build_page(part)
        add_heading(part, f'{d["name"]} —— 解析合集', Pt(14), Pt(0), Pt(6))
        render_toc(part, days, plan, title_of, only_day=d['name'])
        part.add_page_break()
        add_heading(part, f'{d["name"]} —— 解析合集', Pt(14), Pt(8), Pt(4))
        for day, kind, num in plan:
            if day != d['name']:
                continue
            render_chapter(part, title_of(day, kind, num), path_of(day, kind, num))
        save_doc(part, os.path.join(OUT_PARTS, f'{d["name"]}.docx'))

    # 顺序清单
    lines = ['顺序清单（实际拼装顺序）', '=' * 60]
    for i, (day, kind, num) in enumerate(plan, 1):
        lines.append(f'{i:02d}. {title_of(day, kind, num)}')
    with open(os.path.join(OUT_BOOK, '顺序清单.txt'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')

    print(f'章节总数: {len(plan)}（应为 52 讲 + 14 共鸣点 + 9 阶段排查 = 75）')
    print(f'输出: {OUT_BOOK} / {OUT_PARTS}')


if __name__ == '__main__':
    main()
