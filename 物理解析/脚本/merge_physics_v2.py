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
import io
import os
import re
import sys

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from PIL import Image
from lxml import etree

import merge_physics as MP          # 复用：全部正则、抽取判定、目录映射
import frontmatter as FM            # 封面 / 说明页 / 版本记录（与生物项目共用）

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
V = '{urn:schemas-microsoft-com:vml}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
COPY_TAGS = {W + 'r', M + 'oMath', M + 'oMathPara', W + 'hyperlink'}
IMG_TAGS = {W + 'drawing', W + 'pict'}
EMU_CM = 360000          # 1 cm = 360000 EMU
MIN_FIG_CM = 1.0         # 小于 1cm 宽的不是插图，是行内符号碎屑（实测源里有大量 0.11cm 的）


def drawing_width_cm(el):
    """取图片宽度（cm）；取不到返回 0"""
    if el.tag == W + 'pict':
        shape = el.find('.//' + V + 'shape')
        if shape is not None:
            try:
                return float(shape.get('style', '').split('width:')[1].split('pt')[0]) / 28.35
            except Exception:
                return 0.0
        return 0.0
    ext = el.find('.//' + WP + 'extent')
    if ext is None:
        return 0.0
    try:
        return int(ext.get('cx') or 0) / float(EMU_CM)
    except (TypeError, ValueError):
        return 0.0


def collect_images(allp, i0, i1, min_cm=MIN_FIG_CM):
    """收集 allp[i0..i1] 范围内的图片元素，按出现顺序，滤掉碎屑

    必须传**全量段落列表**：只有图没有文字的段落，`p.text` 是空串，
    若用「非空段落」列表会把这些段落连带图片一起漏掉（实测漏掉近一半）。
    """
    out = []
    for k in range(i0, min(i1 + 1, len(allp))):
        for el in allp[k]._element.iter():
            if el.tag in IMG_TAGS and drawing_width_cm(el) >= min_cm:
                out.append(el)
    return out


_IMG_FALLBACK = {'emf': 'image/x-emf', 'wmf': 'image/x-wmf', 'svg': 'image/svg+xml',
                 'tif': 'image/tiff', 'tiff': 'image/tiff', 'bmp': 'image/bmp',
                 'gif': 'image/gif', 'png': 'image/png', 'jpg': 'image/jpeg',
                 'jpeg': 'image/jpeg'}


_MAX_FIG_PX = 1400       # 栏宽 8.8cm 在 300dpi 下约 1040px，留点余量取 1400


def _looks_gray(im):
    """抽样判断是否近似灰度（教辅插图多为黑白线条图，转灰度能省一大半）"""
    if im.mode in ('1', 'L', 'LA'):
        return True
    try:
        rgb = im.convert('RGB')
        s = rgb.resize((min(64, rgb.size[0]), min(64, rgb.size[1])))
        px = list(s.getdata())
        step = max(1, len(px) // 400)
        for r, g, b in px[::step]:
            if abs(r - g) > 14 or abs(g - b) > 14 or abs(r - b) > 14:
                return False
        return True
    except Exception:
        return False


def _shrink_image(blob, src_partname):
    """压缩图片体积：超宽图等比缩小 + 统一重压 PNG

    实测源文件里的图**质量参差**：有 1890×784 却 2.2MB 的不压缩 TIFF，
    也有 475×201 却 375KB 的近乎裸存 PNG（3.9 字节/像素）。
    不处理的话合订本 docx 34MB / PDF 73MB，打印店上传常常超限。
    只在压完更小时才采用；解析失败（EMF/WMF）原样返回走回退路径。
    """
    ext = os.path.splitext(str(src_partname))[1].lstrip('.').lower()
    try:
        im = Image.open(io.BytesIO(blob))
        w, h = im.size
    except Exception:
        return blob, ext
    if w > _MAX_FIG_PX:
        im = im.resize((_MAX_FIG_PX, max(1, int(h * _MAX_FIG_PX / float(w)))), Image.LANCZOS)
    buf = io.BytesIO()
    try:
        if _looks_gray(im):
            im.convert('L').save(buf, 'PNG', optimize=True)
        else:
            im.convert('RGB').save(buf, 'PNG', optimize=True)
    except Exception:
        return blob, ext
    out = buf.getvalue()
    return (out, 'png') if len(out) < len(blob) else (blob, ext)


def _image_rid(dst_doc, blob, src_partname):
    """把图片加入目标文档并返回新 rId

    ① 先压缩：TIFF 转 PNG、超宽图缩到 _MAX_FIG_PX（否则合订本 PDF 会到 70MB+）
    ② 优先走 python-docx 标准路径 get_or_add_image：按内容 SHA1 去重、自动分配唯一部件名，
       避免「不同源文档的 imageNN.png 重名」导致打包出现 Duplicate name
    ③ EMF/WMF 等 python-docx 不认识的格式会抛 UnrecognizedImageError，走手动建部件的回退路径
    """
    blob, ext = _shrink_image(blob, src_partname)
    try:
        rId, _ = dst_doc.part.get_or_add_image(io.BytesIO(blob))
        return rId
    except Exception:
        pass
    pkg = dst_doc.part.package
    if not ext:
        ext = 'bin'
    partname = pkg.next_partname('/word/media/image%d.' + ext)
    ct = _IMG_FALLBACK.get(ext, 'application/octet-stream')
    part = Part(partname, ct, blob, pkg)
    return dst_doc.part.relate_to(part, RT.IMAGE)


def fix_image_rids(el, src_doc, dst_doc):
    """复制过来的图片，r:embed / r:id 指向源文档的关系 id，必须换成目标文档的"""
    for blip in el.iter(A + 'blip'):
        rid = blip.get(R + 'embed')
        if rid:
            rel = src_doc.part.rels.get(rid)
            if rel is not None and not rel.is_external:
                tp = rel.target_part
                blip.set(R + 'embed', _image_rid(dst_doc, tp.blob, tp.partname))
        if blip.get(R + 'link') is not None:
            blip.attrib.pop(R + 'link', None)      # 外链图不跟随，去掉链接
    for imd in el.iter(V + 'imagedata'):           # VML 老式图片
        rid = imd.get(R + 'id')
        if rid:
            rel = src_doc.part.rels.get(rid)
            if rel is not None and not rel.is_external:
                tp = rel.target_part
                imd.set(R + 'id', _image_rid(dst_doc, tp.blob, tp.partname))


def anchor_to_inline(drawing):
    """浮动图（wp:anchor）转内嵌图（wp:inline），否则搬运后会压字"""
    anc = None
    for child in drawing:
        if child.tag == WP + 'anchor':
            anc = child
            break
    if anc is None:
        return
    inl = OxmlElement('wp:inline')
    for k in ('distT', 'distB', 'distL', 'distR'):
        v = anc.get(qn('wp:' + k))
        if v:
            inl.set(qn('wp:' + k), v)
    for child in anc:
        if child.tag in (WP + 'extent', WP + 'effectExtent', WP + 'docPr',
                         WP + 'cNvGraphicFramePr', A + 'graphic'):
            inl.append(copy.deepcopy(child))
    drawing.remove(anc)
    drawing.append(inl)


def scale_drawing(drawing, max_cm):
    """宽度超栏宽的图片等比缩到栏宽"""
    max_emu = int(max_cm * EMU_CM)
    ext = drawing.find('.//' + WP + 'extent')
    if ext is None:
        return
    try:
        cx, cy = int(ext.get('cx') or 0), int(ext.get('cy') or 0)
    except ValueError:
        return
    if cx <= 0 or cx <= max_emu:
        return
    r = max_emu / float(cx)
    ncx, ncy = int(cx * r), int(cy * r)
    for node in drawing.iter(WP + 'extent'):
        node.set('cx', str(ncx)); node.set('cy', str(ncy))
    for node in drawing.iter(A + 'ext'):
        node.set('cx', str(ncx)); node.set('cy', str(ncy))

ROOT = os.path.dirname(os.path.abspath(__file__))

# ── 版面参数（默认 = 生物项目实测择优值）──
COLS = 2
PAPER = 'a4'
MARGIN_LR = 0.8
MARGIN_TB = 0.4
# 图片最大宽度：A4 双栏 0.8cm 边距时栏宽 9.4cm，留点余量
COL_W_CM = 8.8
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
    allp = list(doc.paragraphs)                       # 全量段落（含“纯图无字”的段）
    pos = [i for i, p in enumerate(allp) if p.text.strip()]   # 非空段在全量里的下标
    paras = [allp[i] for i in pos]
    if region != 'all':
        cut = split_boundary(paras)
        if cut is None:
            if region == 'practice':
                return []
            # 微点突破 / 阶段复习 无精练 → 整篇归大本
        elif region == 'main':
            paras, pos = paras[:cut], pos[:cut]
        else:
            paras, pos = paras[cut:], pos[cut:]
    texts = [p.text.strip() for p in paras]
    entries = []
    i, n = 0, len(texts)
    while i < n:
        t = texts[i]
        if MP.is_question_start(texts, i, allow_sub=True):
            qs = i
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
                # 只带「解析自己的图」，范围为【首个 答案/解析 段 → 链末(含其后连续纯图段)】。
                # 图有两种摆法，都要覆盖：
                #   ① 夹在解析中间：[解析] 对C受力分析如图，→ [★图] → [续段] 由相似三角形可知…
                #   ② 摆在解析下面：[解析] …如图甲…如图乙… → [★图甲][★图乙] → 下一段
                # 范围到链末后，只继续吞「纯图片段」（无文字），一遇到有文字的段立刻停 ——
                # 否则会把后面「方法总结框」「下一节」的图也算进来。
                # 题图在「题首之后、答案之前」，不在范围内，自然排除。
                end_all = pos[ki[-1]]
                k2 = end_all + 1
                while k2 < len(allp) and not allp[k2].text.strip():
                    end_all = k2
                    k2 += 1
                entries.append({'num': qnum, 'src': src,
                                'texts': [texts[k] for k in ki],
                                'paras': [paras[k] for k in ki],
                                'imgs': collect_images(allp, pos[ki[0]], end_all),
                                'srcdoc': doc})
            i = j
            continue
        if MP.is_keeper_line(t):
            ki = [i]
            j = i + 1
            while j < n and MP.is_cont_line(texts[ki[-1]], texts[j]):
                ki.append(j)
                j += 1
            entries.append({'num': '', 'src': '', 'texts': [texts[k] for k in ki],
                            'paras': [paras[k] for k in ki],
                            'imgs': collect_images(allp, pos[ki[0]], pos[ki[-1]]),
                            'srcdoc': doc})
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


def copy_para_into(dst_p, src_p, src_doc=None, dst_doc=None):
    """把源段落的内容（run + OMML 公式 + 超链接 + 内嵌图）原样复制到目标段落"""
    for child in src_p._element:
        if child.tag in COPY_TAGS:
            new = copy.deepcopy(child)
            if src_doc is not None and dst_doc is not None:
                for d in new.iter(W + 'drawing'):
                    anchor_to_inline(d)
                fix_image_rids(new, src_doc, dst_doc)
            dst_p._element.append(new)


def render_entry(doc, entry):
    ps = entry.get('paras') or []
    if not ps:
        return
    sd = entry.get('srcdoc')
    # 首行：题号+来源（加粗） + 源段落内容
    p = _new_para(doc)
    head = (entry['num'] or '') + (entry.get('src') or '')
    if head:
        rn = p.add_run(head + ' ')
        MP.set_font(rn, MP.BODY_FONT, MP.BODY_SIZE, bold=True)
    copy_para_into(p, ps[0], sd, doc)
    # 题图 / 选项图：插在“答案”之后、“解析”之前，正是原书里图的位置
    for d in entry.get('imgs') or []:
        ip = _new_para(doc)
        ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
        nd = copy.deepcopy(d)
        anchor_to_inline(nd)
        if sd is not None:
            fix_image_rids(nd, sd, doc)
        scale_drawing(nd, COL_W_CM)
        ip._element.append(nd)
    # 续段：逐段原样复制
    for sp in ps[1:]:
        copy_para_into(_new_para(doc), sp, sd, doc)


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

    n_entries = n_ans = n_ana = n_fig = 0
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
            n_fig += len(e.get('imgs') or [])
            render_entry(doc, e)
        order.append('  %02d. %-44s <- %s' % (len(order), title, os.path.basename(fpath)))

    FM.new_page_section(doc, cols=1, numbered=True)
    FM.build_version_page(doc, '物理', [
        ('大本解析' if region == 'main' else '小本解析',
         '%d %s · %d 条目' % (len(items), '章' if region == 'main' else '练', n_entries))])
    FM.enable_even_odd_headers(doc)
    saved = MP.save_doc(doc, out_name)
    print('%s: %s  章/练 %d  条目 %d  答案 %d  解析 %d  插图 %d'
          % (region, os.path.basename(saved), len(items), n_entries, n_ans, n_ana, n_fig))
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
    FM.new_page_section(doc, cols=1, numbered=True)
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
