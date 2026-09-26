#!/usr/bin/env python3
"""
去除 docx 文档中的全部知识点讲解，仅保留解析相关内容。

保留规则：
- 文档主标题（如 "第1讲　走近细胞"、"周末必刷1"）
- 考点速览判断题整行: (数字)...(×/√)
- 提示行: 提示　...
- 答案行: 答案　X
- 解析行: 解析　...
其余全部删除（体系构建、新增内容、课标要求、考情分析、考点讲解、考向标题等）。
"""

import re
import os
from docx import Document
from docx.oxml.ns import qn

# ── 正则模式 ──

# 考点速览判断题  "(1)...(×)" 或 "(1)...(√)"
PATTERN_TF_LINE = re.compile(r'^\(\d+\)')
TF_MARKER = re.compile(r'[\(（][×√][\)）]')

# 保留标记行
KEEPER_PATTERNS = [
    re.compile(r'^答案\s*[　 ]'),
    re.compile(r'^解析\s*[　 ]'),
    re.compile(r'^提示\s*[　 ]'),
]

# 文档主标题特征（不同大概念文件夹的文件名前缀）
TITLE_PATTERNS = [
    re.compile(r'^第\d+讲'),
    re.compile(r'^周末必刷\d+'),
    re.compile(r'^阶段排查'),
    re.compile(r'^高考共鸣点\d+'),
]

# ── 续段判定（多段答案/解析/提示的截断修复）──

# 答案/解析/提示 + 空格（与 KEEPER_PATTERNS 同义，供续段判断复用）
RE_KEEP_STRICT = re.compile(r'^(答案|解析|提示)\s*[　 ]')

# 常见小标题/框题特征 → 不会是答案续段
HEADING_RE = re.compile(
    r'^(考向|考点|限时练|阶段|周末必刷|第\d+讲|高考共鸣点|回扣落实|'
    r'【|】|综合提升|巩固|基础|大概念|本讲|课标|目录|练\s*[··]|'
    r'思维|方法|技巧|易错|警示|归纳|总结|拓展|延伸|微网|构建|'
    r'类型|热点|题型|专题|微专题|变式|注意|提醒)'
)
# 编号续行："(2)…" / "（3）…" / "②…"
NUM_MARK_RE = re.compile(r'^[（(]\s*\d+\s*[）)]|^[①②③④⑤⑥⑦⑧⑨⑩]')
# 题号起始："1." / "1、" / "1．"（新题 / 列表项，不算续段）
Q_START_RE = re.compile(r'^\d+[.、．]')
# 孤立标签行：单独成段的 "答案" / "解析" / "提示" 或 "答案："
BARE_LABEL_RE = re.compile(r'^(答案|解析|提示)[：:]?$')
# 拆分判断题：题面段 "(4)…" 无标记 + 下一段纯来源标记 "(2024·甘肃卷，1C)(×)"
RE_TF_START = re.compile(r'^[（(]\s*\d+\s*[）)]')
RE_TF_MARK = re.compile(r'[×√]')
RE_SRC_LINE = re.compile(r'^[\(（]\d{4}[^（(]{0,30}[\)）]\s*[\(（][×√][\)）]\s*$')
# 答案/解析/提示 + 编号开头："答案　(1)…" / "提示　①…"
PREF_NUM_RE = re.compile(r'^(?:答案|解析|提示)\s*[　 ]*[（(]?\s*(\d+)\s*[）)]')
CN_NUM = {'①': 1, '②': 2, '③': 3, '④': 4, '⑤': 5,
          '⑥': 6, '⑦': 7, '⑧': 8, '⑨': 9, '⑩': 10}
END_SENT = '。？！？)）】'


def leading_num(text):
    """取行首编号：'答案　(1)…' -> 1；'(2)…' -> 2；'②…' -> 2；无编号 -> None"""
    m = PREF_NUM_RE.match(text)
    if m:
        return int(m.group(1))
    m = re.match(r'^[（(]\s*(\d+)\s*[）)]', text)
    if m:
        return int(m.group(1))
    if text and text[0] in CN_NUM:
        return CN_NUM[text[0]]
    return None


def is_cont_like(prev):
    """prev 是否为“答案/解析/提示 行或已接受的续段”（标题/判断题不算）"""
    if RE_KEEP_STRICT.match(prev) or BARE_LABEL_RE.match(prev):
        return True
    if is_title(prev):
        return False
    # 无前缀续段：只要不以句末标点结尾（如“…卵细胞”），仍属于解析流
    return prev[-1] not in END_SENT


def is_cont_line(prev, cur):
    """
    判断 cur 是否为 prev（答案/解析/提示 行或裸标签）的续段：
    - 编号续行：prev 与 cur 都以编号开头（如 答案(1) → (2)、提示(1) → ②）
    - 裸标签续段：prev 是单独成段的 答案/解析/提示 → cur 为真实内容
    - 纯文字续段：本段未以句末标点结尾（句子被拆成多段，含无前缀续段的链式延续）
    """
    if not cur:
        return False
    if is_keep_line(cur) or is_title(cur) or is_image_ref(cur) or is_table_line(cur):
        return False
    if Q_START_RE.match(cur) or HEADING_RE.match(cur):
        return False
    if NUM_MARK_RE.match(cur):
        pn = leading_num(prev)
        cn = leading_num(cur)
        return pn is not None and cn is not None and cn > pn
    if BARE_LABEL_RE.match(cur):
        return bool(RE_KEEP_STRICT.match(prev) or BARE_LABEL_RE.match(prev))
    if BARE_LABEL_RE.match(prev):
        return True
    if is_cont_like(prev) and prev[-1] not in END_SENT:
        return True
    return False


def is_title(text):
    """判断是否为文档主标题（第N讲/周末必刷N/阶段排查/高考共鸣点N）"""
    return any(p.match(text) for p in TITLE_PATTERNS)


def is_keep_line(text):
    """判断是否是需要保留的行"""
    if not text:
        return False
    # 考点速览判断题整行: (数字)...(×/√)
    if PATTERN_TF_LINE.match(text) and TF_MARKER.search(text):
        return True
    # 答案/解析/提示行
    for pat in KEEPER_PATTERNS:
        if pat.match(text):
            return True
    # 文档主标题
    for pat in TITLE_PATTERNS:
        if pat.match(text):
            return True
    return False


def is_image_ref(text):
    """判断是否为图片引用行 [F:...]"""
    return text.startswith('[F:') or text.startswith('[f:')


def is_table_line(text):
    """判断是否为表格装饰行  +:+..."""
    return text.startswith('+:') and text.endswith(':+')


def process_docx(input_path, output_path):
    """处理单个 docx 文件"""
    doc = Document(input_path)
    paras = doc.paragraphs
    body = doc.element.body
    
    # ── 第一阶段：确定要保留的段落（含多段答案/解析/提示的续段）──
    keep_indices = set()
    last_kept = None  # 链式判断：最近一个已保留/续段行

    for idx, para in enumerate(paras):
        text = para.text.strip()

        if not text:
            continue

        # 跳过图片引用和表格装饰线（不打断续段链）
        if is_image_ref(text) or is_table_line(text):
            continue

        if is_keep_line(text):
            keep_indices.add(idx)
            last_kept = text
            continue

        # 拆分判断题：题面 "(4)…"（本段无×/√），下一段为纯来源标记 "(2024·甘肃卷，1C)(×)"
        if (RE_TF_START.match(text) and not RE_TF_MARK.search(text)
                and idx + 1 < len(paras)
                and RE_SRC_LINE.match(paras[idx + 1].text.strip())):
            keep_indices.add(idx)
            keep_indices.add(idx + 1)
            last_kept = paras[idx + 1].text.strip()
            continue

        # 多段答案/解析/提示：紧跟在保留行后的续段（(2)/(3)/②…或句子被拆段的续文）
        if last_kept and is_cont_line(last_kept, text):
            keep_indices.add(idx)
            last_kept = text
            continue

        last_kept = None
    
    # 如果没有任何保留行被命中，尝试保存第一个有意义的段落作为标题
    if not keep_indices:
        for idx, para in enumerate(paras):
            text = para.text.strip()
            if text and not is_image_ref(text) and not is_table_line(text):
                keep_indices.add(idx)
                break
    
    # ── 第二阶段：删除所有非保留段落和所有表格 ──
    
    # 删除段落（从后往前避免索引变化）
    for idx in range(len(paras) - 1, -1, -1):
        if idx not in keep_indices:
            p_element = paras[idx]._element
            try:
                body.remove(p_element)
            except (ValueError, AttributeError):
                pass
    
    # 删除所有表格（表格属于知识点内容）
    for table in doc.tables:
        tbl_element = table._element
        try:
            body.remove(tbl_element)
        except (ValueError, AttributeError):
            pass
    
    # ── 第三阶段：清理孤立空白段落 ──
    # 连续的空段落只保留一个
    
    # 保存
    doc.save(output_path)
    removed_count = len(paras) - len(keep_indices)
    return removed_count


def get_all_roots():
    """自动发现所有大概念文件夹"""
    roots = []
    for entry in os.listdir('.'):
        if os.path.isdir(entry) and entry.startswith('大概念'):
            roots.append(entry)
    roots.sort()
    return roots


def main():
    roots = get_all_roots()
    if not roots:
        print("未找到大概念文件夹！")
        return
    
    print(f"发现 {len(roots)} 个大概念文件夹: {roots}\n")
    
    total_files = 0
    total_removed = 0
    
    for root in roots:
        if not os.path.isdir(root):
            print(f"跳过（不是目录）: {root}")
            continue
        
        for fname in os.listdir(root):
            if not fname.endswith('.docx') or fname.startswith('~$'):
                continue
            if fname.endswith('_解析.docx'):
                continue  # 跳过已处理文件
            
            input_path = os.path.join(root, fname)
            name_only, ext = os.path.splitext(fname)
            output_path = os.path.join(root, f"{name_only}_解析.docx")
            
            try:
                removed = process_docx(input_path, output_path)
                print(f"  OK {fname} -> {name_only}_解析.docx (removed {removed} paras+tables)")
                total_files += 1
                total_removed += removed
            except Exception as e:
                print(f"  FAIL {fname}: {e}")
    
    print(f"\nDone! {total_files} files processed, {total_removed} paragraphs+tables removed.")
    print("_解析.docx files created. Originals unchanged.")


if __name__ == '__main__':
    main()
