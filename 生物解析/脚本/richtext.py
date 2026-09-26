# -*- coding: utf-8 -*-
"""
richtext.py — 域感知 + 上下标感知的 docx 读取/渲染模块

解决原链路的两类丢失：
  ① Word 域（EQ）：字符写在 <w:instrText> 里，python-docx 的 paragraph.text 完全读不到
     例：NO + eq \o\al(－,3)  →  原链路读出 "NO"（丢 "3－"）
  ② 上下标格式（w:vertAlign）：字符在 <w:t> 里能读到，但下标/上标格式被丢掉
     例：KNO₃ 的 3 带 subscript  →  原链路读成平坦的 "KNO3"

对外接口：
  read_para(para)      -> [Span]     带 sub/sup 标记的片段序列
  para_plain(para)     -> str        纯文本（含域还原字符），供正则匹配/统计
  add_rich(p, spans)   -> None       把片段渲染进一个 docx 段落（真上下标）

本模块只读原始 docx；渲染写入由调用方决定。
"""

import re

from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


class Span(object):
    """一段文字 + 上下标标记"""
    __slots__ = ('text', 'sub', 'sup')

    def __init__(self, text, sub=False, sup=False):
        self.text = text
        self.sub = sub
        self.sup = sup

    def __repr__(self):
        f = ('sub' if self.sub else '') + ('sup' if self.sup else '')
        return 'Span(%r%s)' % (self.text, (',' + f) if f else '')


# ─────────────────────────── EQ 域解释器 ───────────────────────────

def _split_args(s):
    """按顶层逗号拆分参数（忽略嵌套括号内的逗号）"""
    out, depth, cur = [], 0, []
    for ch in s:
        if ch in '(':
            depth += 1
            cur.append(ch)
        elif ch == ')':
            depth -= 1
            cur.append(ch)
        elif ch == ',' and depth == 0:
            out.append(''.join(cur))
            cur = []
        else:
            cur.append(ch)
    out.append(''.join(cur))
    return [x.strip() for x in out]


def _strip_switches(body):
    """去掉 \\xx 形式的开关，保留括号内的可见内容"""
    return re.sub(r'\\[a-z]+[0-9]*', ' ', body)


def eq_to_spans(instr):
    """
    把 EQ 域指令翻译成 Span 序列。

    已覆盖本项目的全部写法：
      \\o\\al(上标,下标)   化学式角标 → 下标 + 上标（NO 3 － → NO₃⁻）
      \\o(―→,\\s\\up(标注)) 反应箭头   → 箭头 + 上标标注
      \\x(文字)            方框文字   → 文字
      \\f(分子,分母)        分数       → 分子/分母
      \\a\\al(a,b) / \\co1 阵列       → 按行拆分
      \\r(X)               根号       → √X
      \\s\\upN(..)/\\s\\doN(..) 升降  → 上标/下标
    """
    s = instr.strip()
    m = re.match(r'eq\b(.*)$', s, re.I | re.S)
    if not m:
        return []
    body = m.group(1).strip()
    return _parse_eq(body)


def _parse_eq(body):
    spans = []

    # 分数 \f(分子,分母)
    m = re.search(r'\\f\s*\(', body)
    if m:
        inner, _ = _read_paren(body, m.end() - 1)
        args = _split_args(inner)
        if len(args) == 2:
            num, den = args
            num = ''.join(sp.text for sp in _parse_eq(num)).strip()
            den = ''.join(sp.text for sp in _parse_eq(den)).strip()
            if re.search(r'[＋+\-－]', num) and not (num.startswith('(') or num.startswith('（')):
                num = '(' + num + ')'
            return [Span('%s/%s' % (num, den))]
        return [Span(inner)]

    # 阵列 / 方程组 \a\al(...) \a\vs4\al(...) \co1(...)
    m = re.search(r'\\c?o?1?\s*\(', body) if re.search(r'\\co1\s*\(', body) else None
    if re.search(r'\\co1\s*\(', body):
        i = body.index('\\co1')
        inner, _ = _read_paren(body, body.index('(', i))
        items = _split_args(inner)
        parts = [''.join(sp.text for sp in _parse_eq(x)).strip() for x in items]
        return [Span('；'.join(p for p in parts if p))]
    if re.search(r'\\a\\?(vs\d+)?\\?al\s*\(', body):
        i = body.find('\\al')
        if i < 0:
            i = body.find('\\a')
        j = body.find('(', i)
        inner, _ = _read_paren(body, j)
        items = _split_args(inner)
        parts = [''.join(sp.text for sp in _parse_eq(x)).strip() for x in items]
        # \a\al 多为「按显示宽度折行的整句」，直接拼接；\co1 才是并列式子
        return [Span(''.join(p for p in parts if p))]

    # 角标 \o\al(上标,下标)  —— 先下标后上标，符合化学式书写顺序
    m = re.search(r'\\o\s*\\al\s*\(', body)
    if m:
        inner, _ = _read_paren(body, m.end() - 1)
        args = _split_args(inner)
        if len(args) == 2:
            sup_txt, sub_txt = args[0], args[1]
            out = []
            if sub_txt:
                out.append(Span(sub_txt, sub=True))
            if sup_txt:
                out.append(Span(sup_txt, sup=True))
            return out
        return [Span(inner)]

    # 箭头/覆盖 \o(A,\s\upN(B)) → A 然后 B 作上标；\o(A,\s\doN(B)) → B 作下标
    m = re.search(r'\\o\s*\(', body)
    if m:
        inner, _ = _read_paren(body, m.end() - 1)
        args = _split_args(inner)
        out = []
        for a in args:
            mm = re.match(r'\\s\s*\\(up|do)\d*\s*\((.*)\)\s*$', a.strip(), re.S)
            if mm:
                txt = ''.join(sp.text for sp in _parse_eq(mm.group(2)))
                out.append(Span(txt, sub=(mm.group(1) == 'do'), sup=(mm.group(1) == 'up')))
            else:
                txt = ''.join(sp.text for sp in _parse_eq(a))
                if txt.strip():
                    out.append(Span(txt))
        return out

    # 根号 \r(X)
    m = re.search(r'\\r\s*\(', body)
    if m:
        inner, _ = _read_paren(body, m.end() - 1)
        return [Span('√' + ''.join(sp.text for sp in _parse_eq(inner)))]

    # 独立升降 \s\upN(..) / \s\doN(..)  —— 必须在 \o(...) 之后判断
    m = re.match(r'\s*\\s\s*\\(up|do)\d*\s*\(', body)
    if m:
        inner, _ = _read_paren(body, body.index('(', m.end() - 1))
        txt = ''.join(sp.text for sp in _parse_eq(inner))
        return [Span(txt, sub=(m.group(1) == 'do'), sup=(m.group(1) == 'up'))]

    # 方框 \x(文字)
    m = re.search(r'\\x\s*\(', body)
    if m:
        inner, _ = _read_paren(body, m.end() - 1)
        return _parse_eq(inner)

    # 兜底：去掉开关后取可见字符
    bare = _strip_switches(body)
    bare = re.sub(r'[()]', '', bare)
    bare = re.sub(r'\s+', '', bare)
    return [Span(bare)] if bare else []


def _read_paren(s, i):
    """从 s[i]=='(' 开始读到匹配的 ')'，返回 (内容, 结束位置)"""
    depth = 0
    for k in range(i, len(s)):
        if s[k] == '(':
            depth += 1
        elif s[k] == ')':
            depth -= 1
            if depth == 0:
                return s[i + 1:k], k
    return s[i + 1:], len(s)


# ─────────────────────────── 段落读取 ───────────────────────────

def _run_align(run_el):
    """读 run 的上下标标记"""
    rpr = run_el.find(W + 'rPr')
    if rpr is None:
        return False, False
    va = rpr.find(W + 'vertAlign')
    if va is None:
        return False, False
    v = va.get(W + 'val')
    if v == 'subscript':
        return True, False
    if v == 'superscript':
        return False, True
    return False, False


def read_para(para):
    """按文档顺序读取段落 → [Span]（EQ 域展开、上下标标记保留）"""
    spans = []
    in_field = False
    instr = []
    has_result = False
    result_spans = []

    for el in para._element.iter():
        tag = el.tag
        if tag == W + 'fldSimple':
            ins = el.get(W + 'instr') or ''
            if re.match(r'\s*eq\b', ins, re.I):
                spans.extend(eq_to_spans(ins))
            continue
        if tag != W + 'r':
            continue

        fc = el.find(W + 'fldChar')
        if fc is not None:
            typ = fc.get(W + 'fldCharType')
            if typ == 'begin':
                in_field, instr, has_result, result_spans = True, [], False, []
            elif typ == 'separate':
                has_result = True
            elif typ == 'end':
                if in_field:
                    full = ''.join(instr)
                    if re.match(r'\s*eq\b', full, re.I):
                        # 优先用缓存结果（若存在），否则按指令生成
                        cached = [s for s in result_spans if s.text]
                        if cached:
                            spans.extend(cached)
                        else:
                            spans.extend(eq_to_spans(full))
                    else:
                        spans.extend(result_spans)
                in_field, instr, has_result, result_spans = False, [], False, []
            continue

        it = el.find(W + 'instrText')
        if it is not None:
            if in_field and not has_result:
                instr.append(it.text or '')
            continue

        t = el.find(W + 't')
        if t is not None:
            sub, sup = _run_align(el)
            sp = Span(t.text or '', sub, sup)
            if in_field and has_result:
                result_spans.append(sp)
            elif not in_field:
                spans.append(sp)

    return _merge(spans)


def _merge(spans):
    """合并相邻且上下标标记相同的片段"""
    out = []
    for sp in spans:
        if not sp.text:
            continue
        if out and out[-1].sub == sp.sub and out[-1].sup == sp.sup:
            out[-1].text += sp.text
        else:
            out.append(Span(sp.text, sp.sub, sp.sup))
    return out


def para_plain(para):
    """纯文本（含域还原字符），供正则匹配与统计 —— 替代 para.text"""
    return ''.join(sp.text for sp in read_para(para))


def spans_plain(spans):
    return ''.join(sp.text for sp in spans)


# ─────────────────────────── 渲染 ───────────────────────────

def set_font(run, name='宋体', size=None, bold=False):
    run.font.name = name
    run.font.bold = bold
    if size:
        run.font.size = size
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), name)


def add_rich(paragraph, spans, font='宋体', size=Pt(10.5), bold=False):
    """把 Span 序列渲染进段落（真上标/下标）"""
    for sp in spans:
        if not sp.text:
            continue
        run = paragraph.add_run(sp.text)
        set_font(run, font, size, bold)
        if sp.sub:
            run.font.subscript = True
        if sp.sup:
            run.font.superscript = True
    return paragraph


def add_rich_text(paragraph, text, font='宋体', size=Pt(10.5), bold=False):
    """无格式纯文本快捷渲染"""
    return add_rich(paragraph, [Span(text)], font, size, bold)
