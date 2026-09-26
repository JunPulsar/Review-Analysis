# -*- coding: utf-8 -*-
"""
extract_rich.py — 富文本条目抽取内核（区域切分 + 与 merge_jiexi 完全一致的判定）

相对 merge_jiexi 的两点增强：
  ① 全程使用 richtext 读取（域还原 + 上下标标记保留），不再用 para.text
  ② 支持「区域切分」：把一篇讲义切成【主干区】和【练习区段】
     —— 大本只取主干，小本取练习区段

判定谓词全部复用 merge_jiexi（is_cont_line / extract_qnum / PTN_* 等），
保证此前修复过的 5 类 bug（全角年份、无年份题、知识小标题误判、
拆分判断题、多段续行）行为不变。
"""

import re

from docx import Document

import merge_jiexi as MJ
import richtext as RT

PRAC_RE = re.compile(r'^(限时练|强化练)\s*\d+')
KATY_RE = MJ.KATY_RE if hasattr(MJ, 'KATY_RE') else re.compile(r'^考点[一二三四五六七八九十\d]+\s*[　 :：]')


class RPara(object):
    """富文本段落：spans 用于渲染，text 用于正则判定"""
    __slots__ = ('spans', 'text')

    def __init__(self, spans):
        spans = _strip_spans(spans)
        self.spans = spans
        self.text = RT.spans_plain(spans)

    def __repr__(self):
        return 'RPara(%r)' % self.text[:40]


def _strip_spans(spans):
    """去掉首尾空白（与原链路 para.text.strip() 对齐）"""
    out = []
    for sp in spans:
        if not sp.text:
            continue
        out.append(RT.Span(sp.text, sp.sub, sp.sup))
    # 去首
    while out and not out[0].text.strip():
        out.pop(0)
    while out and not out[-1].text.strip():
        out.pop()
    if not out:
        return []
    out[0] = RT.Span(out[0].text.lstrip(), out[0].sub, out[0].sup)
    out[-1] = RT.Span(out[-1].text.rstrip(), out[-1].sub, out[-1].sup)
    return [s for s in out if s.text]


# ─────────────────────────── 载入与切分 ───────────────────────────

def load_rich(path):
    """读取 docx → [RPara]（非空段落）"""
    doc = Document(path)
    out = []
    for p in doc.paragraphs:
        rp = RPara(RT.read_para(p))
        if rp.text:
            out.append(rp)
    return out


def split_practice(rparas):
    """切成 (主干区, 练习区段)。练习标题及其后全部归练习区段。"""
    for i, rp in enumerate(rparas):
        if PRAC_RE.match(rp.text):
            return rparas[:i], rparas[i:]
    return rparas, []


def practice_name(rparas):
    """从练习区段取练习名，如 '限时练3'"""
    for rp in rparas:
        m = PRAC_RE.match(rp.text)
        if m:
            return re.sub(r'\s+', '', m.group(0))
    return None


# ─────────────────────────── 拆分判断题归一化（富文本版）───────────────────────────

RE_TF_START_P = re.compile(r'^[（(]\s*\d+\s*[）)]')
RE_TF_MARK_P = re.compile(r'[×√]')
RE_SRC_LINE = re.compile(r'^[\(（]\d{4}[^（(]{0,30}[\)）]\s*[\(（][×√][\)）]\s*$')


def normalize_split_judges_rich(rparas):
    """把「题面段 + 来源标记段」的拆分判断题合并成一条（spans 拼接）"""
    out = []
    i = 0
    n = len(rparas)
    while i < n:
        rp = rparas[i]
        if (i + 1 < n and RE_TF_START_P.match(rp.text) and not RE_TF_MARK_P.search(rp.text)
                and RE_SRC_LINE.match(rparas[i + 1].text)):
            nxt = rparas[i + 1]
            merged = RPara(list(rp.spans) + list(nxt.spans))
            out.append(merged)
            i += 2
        else:
            out.append(rp)
            i += 1
    return out


# ─────────────────────────── 条目抽取 ───────────────────────────

class Entry(object):
    """一个条目：num=题号，src=来源，lines=[RPara]，katy=所属考点"""
    __slots__ = ('num', 'src', 'lines', 'katy')

    def __init__(self, num='', src='', lines=None, katy=None):
        self.num = num
        self.src = src
        self.lines = lines or []
        self.katy = katy


def extract_entries(rparas):
    """
    与 merge_jiexi.process_one_docx / merge_book_order.extract_entries_katy 同逻辑，
    额外记录每个条目所属的「考点N　…」标题，并携带富文本。
    """
    texts = [rp.text for rp in rparas]
    entries = []
    cur_katy = None
    i = 0
    n = len(rparas)

    while i < n:
        rp = rparas[i]
        t = rp.text

        # 考点标题（分段标记）
        if KATY_RE.match(t):
            cur_katy = t
            i += 1
            continue

        # ── 考点速览区域：(N)…(×/√) 起，收集后续 提示/答案/解析 ──
        if MJ.PTN_TF_Q.match(t) and MJ.PTN_TF_ANS.search(t):
            block = [rp]
            i += 1
            while i < n:
                nt = rparas[i].text
                if KATY_RE.match(nt):
                    break
                if MJ.PTN_PROMPT.match(nt) or MJ.PTN_ANSWER.match(nt) or MJ.PTN_ANALYSIS.match(nt):
                    block.append(rparas[i])
                    i += 1
                    while i < n and MJ.is_cont_line(block[-1].text, rparas[i].text):
                        block.append(rparas[i])
                        i += 1
                else:
                    break
            entries.append(Entry('', '', block, cur_katy))
            continue

        # ── 主考题区域 ──
        qnum = MJ.extract_qnum(t)
        if qnum is not None:
            m_src = MJ.RE_SRC.search(t)
            src = m_src.group(0) if m_src else ''
            i += 1
            keep = []
            while i < n:
                nt = rparas[i].text
                if MJ.PTN_ANSWER.match(nt) or MJ.PTN_ANALYSIS.match(nt):
                    keep.append(rparas[i])
                    i += 1
                    while i < n and MJ.is_cont_line(keep[-1].text, rparas[i].text):
                        keep.append(rparas[i])
                        i += 1
                elif MJ.PTN_BARE_LABEL.match(nt) and keep:
                    keep.append(rparas[i])
                    i += 1
                    while i < n and MJ.is_cont_line(keep[-1].text, rparas[i].text):
                        keep.append(rparas[i])
                        i += 1
                elif MJ.PTN_PROMPT.match(nt):
                    i += 1
                elif KATY_RE.match(nt):
                    break
                elif MJ.extract_qnum(nt) is not None or (MJ.PTN_TF_Q.match(nt) and MJ.PTN_TF_ANS.search(nt)):
                    break
                else:
                    i += 1
            if keep:
                entries.append(Entry(qnum, src, keep, cur_katy))
            continue

        # ── 孤立答案/解析兜底 ──
        if MJ.PTN_ANSWER.match(t) or MJ.PTN_ANALYSIS.match(t):
            lns = [rp]
            i += 1
            while i < n and MJ.is_cont_line(lns[-1].text, rparas[i].text):
                lns.append(rparas[i])
                i += 1
            entries.append(Entry('', '', lns, cur_katy))
            continue

        i += 1

    return entries
