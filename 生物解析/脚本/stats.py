import os
from docx import Document

orig_chars = 0
orig_files = 0
for d in sorted(os.listdir('.')):
    if os.path.isdir(d) and d.startswith('大概念'):
        for f in os.listdir(d):
            if not f.endswith('.docx') or f.startswith('~$'):
                continue
            fp = os.path.join(d, f)
            doc = Document(fp)
            for p in doc.paragraphs:
                orig_chars += len(p.text)
            orig_files += 1

j_chars = 0
j_files = 0
for root, dirs, files in os.walk('解析合集'):
    for f in files:
        if not f.endswith('.docx'):
            continue
        fp = os.path.join(root, f)
        doc = Document(fp)
        for p in doc.paragraphs:
            j_chars += len(p.text)
        j_files += 1

# 默认A4排版约1400字/页(含图), 紧凑排版约2200字/页
# 原始文档图表占位修正1.67x
orig_pages = orig_chars / 1400 * 1.67
j_pages = j_chars / 2200

print(f'原始文件: {orig_files} 个, 总字符 {orig_chars:,}')
print(f'解析合集: {j_files} 个, 总字符 {j_chars:,}')
print(f'文本保留比例: {j_chars/orig_chars*100:.1f}%')
print()
print(f'=== 页数估算 ===')
print(f'原始全部打印: ~{orig_pages:.0f} 页')
print(f'解析合集打印: ~{j_pages:.0f} 页')
print(f'节省纸张: ~{orig_pages - j_pages:.0f} 页')
print(f'节省比例: {(1 - j_pages/orig_pages)*100:.0f}%')
