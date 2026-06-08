# -*- coding: utf-8 -*-
"""
AI 散户投资课程 · 完整交付 Word 文档生成脚本

功能：
- 读取 Markdown 源文件（大纲 v2 / 卖课话术 / 3 套授课设计 / README）
- 生成正式 Word 文档（封面 / 目录 / 6 大主体 / 页眉页脚）
- 字体：中文宋体（SimSun），英文 Times New Roman
- 标题层级 H1/H2/H3
- 页眉：「AI 散户投资课程 · 课程大纲 + 面试授课设计」+ 日期
- 页脚：第 X 页 / 共 Y 页

运行：python scripts/build_deliverable_docx.py
"""
import os
import sys
import hashlib
import re
from datetime import date

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# === 路径配置 ===
ROOT = r'D:\Learning\AI\面试\AI散户投资课程'
OUT_FILE = os.path.join(ROOT, 'AI散户投资课程_完整交付.docx')
DATE_STR = '2026-06-05'

# === 字体配置 ===
CN_FONT = '宋体'   # SimSun - 中文
EN_FONT = 'Times New Roman'  # 英文
MONO_FONT = 'Consolas'  # 等宽（用于 prompt 块）

# === 颜色 ===
COLOR_H1 = RGBColor(0x1F, 0x3A, 0x5F)      # 深蓝
COLOR_H2 = RGBColor(0x2C, 0x5F, 0x8C)      # 中蓝
COLOR_H3 = RGBColor(0x44, 0x72, 0xC4)      # 浅蓝
COLOR_QUOTE = RGBColor(0x70, 0x70, 0x70)   # 灰
COLOR_CODE_BG = 'F2F2F2'                    # 代码块底色（灰）
COLOR_ACCENT = RGBColor(0xC0, 0x39, 0x2B)  # 红（封面强调）


# ==============================================================
# XML 工具
# ==============================================================
def set_run_fonts(run, cn=CN_FONT, en=EN_FONT, mono=False):
    """同时设置 CJK + ASCII 字体（含 code 字体覆盖）"""
    if mono:
        run.font.name = MONO_FONT
    else:
        run.font.name = en
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn)
    rFonts.set(qn('w:ascii'), en)
    rFonts.set(qn('w:hAnsi'), en)
    rFonts.set(qn('w:cs'), en)


def set_paragraph_shading(paragraph, color_hex):
    """设置段落底色"""
    pPr = paragraph._element.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    pPr.append(shd)


def set_paragraph_borders(paragraph, top=False, bottom=False, left=False, right=False, color='999999', size='6'):
    """设置段落边框"""
    pPr = paragraph._element.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    for edge, on in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        bdr = pBdr.find(qn(f'w:{edge}'))
        if bdr is None:
            bdr = OxmlElement(f'w:{edge}')
            pBdr.append(bdr)
        if on:
            bdr.set(qn('w:val'), 'single')
            bdr.set(qn('w:sz'), size)
            bdr.set(qn('w:space'), '4')
            bdr.set(qn('w:color'), color)
        else:
            bdr.set(qn('w:val'), 'nil')


def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


# ==============================================================
# 文档骨架
# ==============================================================
def init_document():
    """初始化文档：设置默认样式 + 页眉页脚"""
    doc = Document()

    # 全局默认字体
    style = doc.styles['Normal']
    style.font.name = EN_FONT
    style.font.size = Pt(10.5)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), CN_FONT)
    rFonts.set(qn('w:ascii'), EN_FONT)
    rFonts.set(qn('w:hAnsi'), EN_FONT)

    # 页面边距
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.header_distance = Cm(1.2)
        section.footer_distance = Cm(1.2)

        # 页眉
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hr = hp.add_run(f'AI 散户投资课程 · 课程大纲 + 面试授课设计    {DATE_STR}')
        set_run_fonts(hr, cn=CN_FONT, en=EN_FONT)
        hr.font.size = Pt(9)
        hr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        # 页眉下划线
        set_paragraph_borders(hp, bottom=True, color='B0B0B0', size='4')

        # 页脚：第 X 页 / 共 Y 页
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # 用 FIELD 指令插入 PAGE / NUMPAGES
        run_pre = fp.add_run('第 ')
        set_run_fonts(run_pre, cn=CN_FONT, en=EN_FONT)
        run_pre.font.size = Pt(9)
        run_pre.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        _add_field(fp, 'PAGE')
        run_mid = fp.add_run(' 页 / 共 ')
        set_run_fonts(run_mid, cn=CN_FONT, en=EN_FONT)
        run_mid.font.size = Pt(9)
        run_mid.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        _add_field(fp, 'NUMPAGES')
        run_end = fp.add_run(' 页')
        set_run_fonts(run_end, cn=CN_FONT, en=EN_FONT)
        run_end.font.size = Pt(9)
        run_end.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    return doc


def _add_field(paragraph, field_code):
    """添加 Word 字段（如 PAGE / NUMPAGES）"""
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = f' {field_code} '
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar1)
    run._element.append(instrText)
    run._element.append(fldChar2)
    set_run_fonts(run)
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)


# ==============================================================
# 标题/正文/列表/代码/引用/表格
# ==============================================================
def add_h1(doc, text, bookmark=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run_fonts(r)
    r.font.size = Pt(20)
    r.bold = True
    r.font.color.rgb = COLOR_H1
    # 大标题加下边线
    set_paragraph_borders(p, bottom=True, color='1F3A5F', size='12')
    if bookmark:
        _add_bookmark(p, bookmark)
    return p


def add_h2(doc, text, bookmark=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run_fonts(r)
    r.font.size = Pt(15)
    r.bold = True
    r.font.color.rgb = COLOR_H2
    if bookmark:
        _add_bookmark(p, bookmark)
    return p


def add_h3(doc, text, bookmark=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run_fonts(r)
    r.font.size = Pt(12.5)
    r.bold = True
    r.font.color.rgb = COLOR_H3
    if bookmark:
        _add_bookmark(p, bookmark)
    return p


def _add_bookmark(paragraph, name):
    """在段落开头添加书签（用于目录跳转）"""
    bookmark_id = str(abs(hash(name)) % 1000000)
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), bookmark_id)
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), bookmark_id)
    paragraph._element.insert(0, start)
    paragraph._element.append(end)


def add_hyperlink_to_bookmark(paragraph, text, anchor):
    """在段落中加一个指向书签的超链接（用于目录）"""
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), anchor)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), CN_FONT)
    rFonts.set(qn('w:ascii'), EN_FONT)
    rFonts.set(qn('w:hAnsi'), EN_FONT)
    rPr.append(rFonts)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '2C5F8C')
    rPr.append(color)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '21')
    rPr.append(sz)
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._element.append(hyperlink)


def add_para(doc, text, size=10.5, bold=False, italic=False, color=None, align=None,
             indent_first=False, space_after=4, line_spacing=1.4):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_after = Pt(space_after)
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == 'justify':
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent_first:
        p.paragraph_format.first_line_indent = Cm(0.75)
    r = p.add_run(text)
    set_run_fonts(r)
    r.font.size = Pt(size)
    if bold:
        r.bold = True
    if italic:
        r.italic = True
    if color:
        r.font.color.rgb = color
    return p


def add_inline_para(doc, segments, indent_first=False, space_after=4, line_spacing=1.4, align=None):
    """
    添加内联混排段落。
    segments: list of (text, dict) where dict 可包含 bold/italic/color/size/mono/code
    """
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_after = Pt(space_after)
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'justify':
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent_first:
        p.paragraph_format.first_line_indent = Cm(0.75)
    for text, attrs in segments:
        r = p.add_run(text)
        is_mono = attrs.get('mono', False) or attrs.get('code', False)
        set_run_fonts(r, mono=is_mono)
        r.font.size = Pt(attrs.get('size', 10.5))
        if attrs.get('bold'):
            r.bold = True
        if attrs.get('italic'):
            r.italic = True
        if attrs.get('color'):
            r.font.color.rgb = attrs['color']
    return p


def add_quote(doc, text, size=10.5, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.4
    # 左侧色条
    set_paragraph_borders(p, left=True, color='4472C4', size='18')
    r = p.add_run(text)
    set_run_fonts(r)
    r.font.size = Pt(size)
    r.font.color.rgb = color or COLOR_QUOTE
    r.italic = True
    return p


def add_bullet(doc, text, level=0, size=10.5, bold=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6 + 0.6 * level)
    p.paragraph_format.first_line_indent = Cm(-0.6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.35
    bullet_char = '•' if level == 0 else ('–' if level == 1 else '·')
    r0 = p.add_run(f'{bullet_char}  ')
    set_run_fonts(r0)
    r0.font.size = Pt(size)
    r0.font.color.rgb = color or COLOR_H3
    r = p.add_run(text)
    set_run_fonts(r)
    r.font.size = Pt(size)
    if bold:
        r.bold = True
    if color:
        r.font.color.rgb = color
    return p


def add_code_block(doc, code_text):
    """代码块（浅灰底色 + 等宽字体）"""
    # 拆行
    lines = code_text.split('\n')
    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.4)
        p.paragraph_format.right_indent = Cm(0.4)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.2
        set_paragraph_shading(p, COLOR_CODE_BG)
        r = p.add_run(line if line else ' ')
        set_run_fonts(r, mono=True)
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    # 代码块后加一个空段作分隔
    sep = doc.add_paragraph()
    sep.paragraph_format.space_after = Pt(4)


def add_table_from_md(doc, md_lines, header_row=True):
    """Markdown 表格 → docx 表格"""
    if not md_lines:
        return
    # 解析
    rows = []
    for line in md_lines:
        if '|' not in line:
            continue
        # 去除首尾 |
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        rows.append(cells)
    # 过滤对齐行（--- | ---）
    rows = [r for r in rows if not all(set(c) <= set('-: ') and c for c in r)]
    if not rows:
        return
    n_cols = len(rows[0])
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, cell_text in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ''
            p = cell.paragraphs[0]
            r = p.add_run(cell_text)
            set_run_fonts(r)
            r.font.size = Pt(9.5)
            if i == 0 and header_row:
                r.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                # 单元格底色
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), '2C5F8C')
                tcPr.append(shd)
    # 表格后空行
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(6)


# ==============================================================
# Markdown 渲染（手写解析：够用即可）
# ==============================================================
INLINE_BOLD = re.compile(r'\*\*(.+?)\*\*')
INLINE_CODE = re.compile(r'`([^`]+?)`')


def render_inline(text):
    """把 **bold** 和 `code` 转成 segments"""
    segments = []
    pos = 0
    pattern = re.compile(r'\*\*(.+?)\*\*|`([^`]+?)`')
    last = 0
    for m in pattern.finditer(text):
        if m.start() > last:
            segments.append((text[last:m.start()], {}))
        if m.group(1) is not None:
            segments.append((m.group(1), {'bold': True}))
        else:
            segments.append((m.group(2), {'code': True, 'mono': True, 'color': RGBColor(0xC0, 0x39, 0x2B)}))
        last = m.end()
    if last < len(text):
        segments.append((text[last:], {}))
    if not segments:
        segments = [(text, {})]
    return segments


def render_markdown(doc, md_text, h1_bookmark=None, h2_bookmark_map=None):
    """
    极简 Markdown 渲染：H1 / H2 / H3 / 引用 / 列表 / 代码块 / 表格 / 段落
    """
    lines = md_text.split('\n')
    i = 0
    n = len(lines)
    in_code = False
    code_buf = []
    in_table = False
    table_buf = []

    def flush_table():
        nonlocal table_buf
        if table_buf:
            add_table_from_md(doc, table_buf)
            table_buf = []

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # 代码块
        if stripped.startswith('```'):
            if not in_code:
                in_code = True
                code_buf = []
            else:
                # 结束
                in_code = False
                add_code_block(doc, '\n'.join(code_buf))
                code_buf = []
            i += 1
            continue
        if in_code:
            code_buf.append(line)
            i += 1
            continue

        # 空行
        if not stripped:
            i += 1
            continue

        # 分隔
        if stripped == '---':
            # 用空段 + 底边线
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(6)
            set_paragraph_borders(p, bottom=True, color='B0B0B0', size='4')
            i += 1
            continue

        # 标题
        if stripped.startswith('#### '):
            add_h3(doc, stripped[5:].strip())
            i += 1
            continue
        if stripped.startswith('### '):
            txt = stripped[4:].strip()
            bm = None
            if h2_bookmark_map and txt in h2_bookmark_map:
                bm = h2_bookmark_map[txt]
            add_h3(doc, txt, bookmark=bm)
            i += 1
            continue
        if stripped.startswith('## '):
            txt = stripped[3:].strip()
            bm = None
            if h2_bookmark_map and txt in h2_bookmark_map:
                bm = h2_bookmark_map[txt]
            add_h2(doc, txt, bookmark=bm)
            i += 1
            continue
        if stripped.startswith('# '):
            txt = stripped[2:].strip()
            # H1 不在中间使用，封面和章节标题手写
            add_h1(doc, txt, bookmark=h1_bookmark)
            i += 1
            continue

        # 引用
        if stripped.startswith('> '):
            add_quote(doc, stripped[2:].strip())
            i += 1
            continue

        # 表格（连续 | 开头）
        if '|' in stripped and (i + 1 < n) and re.match(r'^\s*\|?[\s:|-]+\|', lines[i + 1]):
            table_buf.append(line)
            in_table = True
            i += 1
            continue
        elif in_table and '|' in stripped:
            table_buf.append(line)
            i += 1
            continue
        elif in_table:
            flush_table()
            in_table = False

        # 列表（- 或 *）
        if re.match(r'^[\-\*]\s+', stripped):
            text = re.sub(r'^[\-\*]\s+', '', stripped)
            add_bullet(doc, text, level=0)
            i += 1
            continue
        # 嵌套列表（  - 或    -）
        m = re.match(r'^( {2,})[\-\*]\s+(.*)$', line)
        if m:
            text = m.group(2)
            level = min(2, len(m.group(1)) // 2)
            add_bullet(doc, text, level=level)
            i += 1
            continue
        # 数字列表
        m = re.match(r'^(\d+)\.\s+(.*)$', stripped)
        if m:
            text = m.group(2)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.6)
            p.paragraph_format.first_line_indent = Cm(-0.6)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.35
            r0 = p.add_run(f'{m.group(1)}.  ')
            set_run_fonts(r0)
            r0.font.size = Pt(10.5)
            r0.bold = True
            r0.font.color.rgb = COLOR_H3
            segs = render_inline(text)
            for s, a in segs:
                rr = p.add_run(s)
                is_mono = a.get('mono') or a.get('code')
                set_run_fonts(rr, mono=is_mono)
                rr.font.size = Pt(a.get('size', 10.5))
                if a.get('bold'):
                    rr.bold = True
                if a.get('color'):
                    rr.font.color.rgb = a['color']
            i += 1
            continue

        # 普通段落（含 **bold** / `code`）
        # 合并连续非空行
        buf = [line]
        j = i + 1
        while j < n and lines[j].strip() and not (
            lines[j].strip().startswith('#') or
            lines[j].strip().startswith('> ') or
            lines[j].strip() == '---' or
            lines[j].strip().startswith('```') or
            re.match(r'^[\-\*]\s+', lines[j].strip()) or
            re.match(r'^(\d+)\.\s+', lines[j].strip()) or
            '|' in lines[j].strip()
        ):
            buf.append(lines[j])
            j += 1
        full = ' '.join(b.strip() for b in buf)
        segs = render_inline(full)
        add_inline_para(doc, segs, line_spacing=1.45, space_after=4)
        i = j
    flush_table()


# ==============================================================
# 封面 / 目录 / 章节
# ==============================================================
def add_cover(doc):
    """封面页"""
    # 上方留白
    for _ in range(3):
        doc.add_paragraph()

    # 主标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run('AI 散户投资课程')
    set_run_fonts(r)
    r.font.size = Pt(36)
    r.bold = True
    r.font.color.rgb = COLOR_H1

    # 副标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run('完整交付包')
    set_run_fonts(r)
    r.font.size = Pt(20)
    r.font.color.rgb = COLOR_H2

    # 副副标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(36)
    r = p.add_run('课程大纲 v2 + 3 套面试授课设计 + 卖课话术')
    set_run_fonts(r)
    r.font.size = Pt(14)
    r.font.color.rgb = COLOR_QUOTE
    r.italic = True

    # 装饰横线
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    set_paragraph_borders(p, bottom=True, color='1F3A5F', size='18')

    # 课程定位
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run('课程定位')
    set_run_fonts(r)
    r.font.size = Pt(12)
    r.bold = True
    r.font.color.rgb = COLOR_H3

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(24)
    r = p.add_run('让一个完全没金融背景的普通散户\n用 AI 工具把"凭感觉亏钱"变成"有纪律地少亏多赚"')
    set_run_fonts(r)
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # 6 大模块速览
    items = [
        '模块 0｜0 基础开箱（2 子节）',
        '模块 1｜别再凭感觉炒股（3 小节）',
        '模块 2｜你的 AI 投资武器库（3 小节）',
        '模块 3｜用 AI 找到能买的标的（3 小节）',
        '模块 4｜把零散玩法变成投资系统（3 小节）',
        '模块 5｜风控 + 心理（3 小节）',
        '模块 6｜复盘——让经验变成真本事（3 小节）',
    ]
    for it in items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run('·  ' + it)
        set_run_fonts(r)
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # 装饰横线
    for _ in range(2):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(8)
    set_paragraph_borders(p, bottom=True, color='1F3A5F', size='18')

    # 作者
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run('作者：方逸之')
    set_run_fonts(r)
    r.font.size = Pt(14)
    r.bold = True

    # 日期
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(DATE_STR)
    set_run_fonts(r)
    r.font.size = Pt(12)
    r.font.color.rgb = COLOR_QUOTE

    # 性质说明
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run('本课程定位为投资教育内容，不构成投资建议')
    set_run_fonts(r)
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = COLOR_QUOTE

    add_page_break(doc)


def add_toc(doc):
    """目录页：用超链接锚到各章"""
    add_h1(doc, '目  录', bookmark='_toc_root')

    # 章节
    sections = [
        ('一、课程定位与卖课话术', '_sec_positioning'),
        ('二、课程大纲 v2（6 模块 × 18 小节 + 模块 0 + 2 附录）', '_sec_outline'),
        ('三、面试授课设计 · 3 套候选', '_sec_lessons'),
        ('    3.1  A 套｜决策沙盘', '_sec_lesson_A'),
        ('    3.2  C 套｜投资人格', '_sec_lesson_C'),
        ('    3.3  D 套｜主力逆向', '_sec_lesson_D'),
        ('四、附：交付物清单', '_sec_assets'),
    ]
    for text, anchor in sections:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Cm(0.5)
        add_hyperlink_to_bookmark(p, text, anchor)

    # 大纲 v2 速览锚
    add_h2(doc, '大纲 v2 模块速览（点击跳转）', bookmark='_sec_outline_overview')
    modules = [
        ('模块 0｜0 基础开箱（2 子节）', '_bm_m0'),
        ('模块 1｜别再凭感觉炒股（3 小节）', '_bm_m1'),
        ('模块 2｜你的 AI 投资武器库（3 小节）', '_bm_m2'),
        ('模块 3｜用 AI 找到能买的标的（3 小节）', '_bm_m3'),
        ('模块 4｜把零散玩法变成投资系统（3 小节）', '_bm_m4'),
        ('模块 5｜风控 + 心理（3 小节）', '_bm_m5'),
        ('模块 6｜复盘——让经验变成真本事（3 小节）', '_bm_m6'),
        ('附录 A｜研究底料速查卡（18 张）', '_bm_appA'),
        ('附录 B｜散户必装的 AI 工具组合包', '_bm_appB'),
    ]
    for text, anchor in modules:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_after = Pt(2)
        add_hyperlink_to_bookmark(p, '· ' + text, anchor)

    add_page_break(doc)


def section_positioning(doc):
    """一、课程定位与卖课话术"""
    add_h1(doc, '一、课程定位与卖课话术', bookmark='_sec_positioning')

    add_h2(doc, '1.1  一句话总结')
    add_quote(doc, '让一个完全没金融背景的普通散户，用 AI 工具把"凭感觉亏钱"变成"有纪律地少亏多赚"——6 个模块 + 18 个小节 + 课前 0 基础开箱 + 2 份附录（研究底料速查卡 + 工具组合包），从扫盲到工具、从选股到复盘，全程实操可落地，学完当天就能用。')

    add_h2(doc, '1.2  50-100 字卖课话术')
    add_quote(doc, '面向"会用 AI 但不会用 AI 投资"的散户，把市面 90% 的"AI 选股噱头课"换成"AI 防亏 + 提效"工具箱。6 模块 18 小节 + 课前 0 基础开箱：教 5 大坑避雷、3 款 AI 工具上手（i 问财 + Kimi + 通义）、5 维度个股梳理、回测 5 大坑、4 类情绪陷阱、4 步复盘官。全程实操 + 国产平替 + 4 套 prompt 直接 copy——学完当天能用，让散户"少亏多赚"。')

    add_h2(doc, '1.3  关键卖点（CEO 内部使用）')
    add_h3(doc, '一句话广告版')
    add_quote(doc, '让 AI 当你的"研究员 + 速读员 + 复盘员"——你只做"决策员"。')

    add_h3(doc, '3 个付费动机')
    add_bullet(doc, '韭菜动机（199 元心理价位）：避免"被 AI 忽悠"+"被群消息骗"——5 大坑 + 信息源识别 + 幻觉识别 4 招建立"防亏护城河"。')
    add_bullet(doc, '老司机动机（299-599 元心理价位）：省时间——AI 速读让 3 小时年报压成 5 分钟，回测 + 复盘让 1 个月 1 次变成 1 周 1 次。')
    add_bullet(doc, '共同动机（⭐⭐⭐⭐⭐ 小节）：2.2 / 2.3 / 3.1 三节是"双视角都买单"的核心——课程的 4 大爆点。')

    add_h2(doc, '1.4  和市面 90% AI 投资课的差异化')
    diff_rows = [
        ['维度', '市面常见', '本课程'],
        ['定位', '"AI 帮你选股赚钱"（红海 + 合规危险）', '"AI 帮你识别风险 + 复盘自己"（蓝海 + 合规）'],
        ['工具栈', '教 ChatGPT 为主（需科学上网）', 'i 问财 + Kimi + 通义（国产平替，全程手机搞定）'],
        ['小节密度', '18 节里 7-8 节"小白入门"', '模块 0 开箱 5 分钟搞定，18 节全是"高密度干货"'],
        ['收益话术', '强调"年化 30%+"', '强调"AI 是加速器不是信号源" + 6 条红线提醒'],
        ['差异化玩法', '用"通用案例"教学', '用"AI 复盘官"+"基金经理对齐"+"信息源识别"——3 个市面 90% 课没做的差异化玩法'],
    ]
    add_table_from_md(doc, ['| ' + ' | '.join(r) + ' |' for r in diff_rows])

    add_h2(doc, '1.5  风险红线（CEO 内部提醒）')
    add_bullet(doc, '课程定位为教育用途，不构成投资建议。')
    add_bullet(doc, '课程坚决不涉及融资融券、场外配资、期货合约、加密货币合约。')
    add_bullet(doc, '课程不教具体股票 / 基金推荐——所有案例使用脱敏样本或公开指数。')
    add_bullet(doc, '课程每节课固定脚注 3 行："本节为投资教育内容，不构成投资建议。AI 输出仅供研究参考。投资有风险，决策需谨慎。"')


def section_outline(doc):
    """二、课程大纲 v2"""
    add_h1(doc, '二、课程大纲 v2', bookmark='_sec_outline')

    # 速览
    add_h2(doc, '2.0  大纲 v2 速览（6 模块 × 18 小节 + 模块 0 + 2 附录）', bookmark='_sec_outline_overview')
    add_quote(doc,
              '本版保留 v1 的 6 模块 × 18 小节骨架，整合 research-team / trader / buy-side-council / retail-panel / ai-pm 五方共 18 处"必改"+ 19 处"建议改"+ 10+ 项"加分项"。'
              '结构性变化仅 3 处：① 新增"模块 0｜0 基础开箱"（不计正式模块）；② 模块 1 1.1 散户坑从 3 扩到 5；③ 附录新增"研究底料速查卡 + 工具组合包"。')

    # 设计方法
    add_h2(doc, '2.1  设计方法：逆向设计 + 闭环')
    add_bullet(doc, '设计方法：逆向设计（先定学员产出，再倒推模块）+ 五方反馈整合。')
    add_bullet(doc, '模块顺序：认知重塑 → AI 工具盘点 → 标的筛选 → 投资系统 → 风控与心理 → 实战复盘（闭环）。')
    add_bullet(doc, '小节标题口语化，无 Alpha / Beta / Sharpe / Kelly 等机构黑话。')
    add_bullet(doc, '每个模块开头都有"学习产出"——明确"学完能多做什么"。')
    add_bullet(doc, '每个模块引言都加"三视角提示（宏观/行业/量化）"+"警示（研究 ≠ 预测）"。')

    # 覆盖说明
    add_h2(doc, '2.2  课程覆盖与性质')
    add_bullet(doc, '覆盖市场：A 股 + 港美股 + 基金。')
    add_bullet(doc, '难度：低门槛、强实操。')
    add_bullet(doc, '性质：仅作教育用途，不构成投资建议。')

    # 总体修订说明
    add_h2(doc, '2.3  v1 → v2 总体修订说明')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 总体修订说明', end_marker='## 模块 0'))

    # 模块 0
    add_h2(doc, '模块 0｜0 基础开箱（不计正式模块）', bookmark='_bm_m0')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 模块 0', end_marker='## 模块 1'))

    # 模块 1
    add_h2(doc, '模块 1｜别再凭感觉炒股——先搞懂游戏规则（认知重塑）', bookmark='_bm_m1')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 模块 1', end_marker='## 模块 2'))

    # 模块 2
    add_h2(doc, '模块 2｜你的 AI 投资武器库——主流工具上手', bookmark='_bm_m2')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 模块 2', end_marker='## 模块 3'))

    # 模块 3
    add_h2(doc, '模块 3｜用 AI 找到能买的标的（A 股 + 港美股 + 基金）', bookmark='_bm_m3')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 模块 3', end_marker='## 模块 4'))

    # 模块 4
    add_h2(doc, '模块 4｜把零散玩法变成"你自己的投资系统"', bookmark='_bm_m4')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 模块 4', end_marker='## 模块 5'))

    # 模块 5
    add_h2(doc, '模块 5｜风控 + 心理——散户最容易翻车的两块', bookmark='_bm_m5')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 模块 5', end_marker='## 模块 6'))

    # 模块 6
    add_h2(doc, '模块 6｜复盘——让经验变成真本事', bookmark='_bm_m6')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 模块 6', end_marker='## 附录 A'))

    # 附录 A
    add_h2(doc, '附录 A｜研究底料速查卡（每小节 1 张）', bookmark='_bm_appA')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 附录 A', end_marker='## 附录 B'))

    # 附录 B
    add_h2(doc, '附录 B｜散户必装的 AI 工具组合包（按预算分档）', bookmark='_bm_appB')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## 附录 B', end_marker='## v1 → v2'))

    # v1 → v2 修订总览
    add_h2(doc, '2.4  v1 → v2 修订总览（修改表）')
    render_markdown(doc, _read('大纲_v2.md', start_marker='## v1 → v2', end_marker='## 教学交付说明'))


def _read(name, start_marker=None, end_marker=None):
    """读取源 .md 文本，可按 marker 切片"""
    p = os.path.join(ROOT, name)
    with open(p, 'r', encoding='utf-8') as f:
        text = f.read()
    if start_marker:
        idx = text.find(start_marker)
        if idx < 0:
            return ''
        text = text[idx:]
    if end_marker:
        idx = text.find(end_marker)
        if idx > 0:
            text = text[:idx]
    return text


def section_lessons(doc):
    """三、面试授课设计 · 3 套候选"""
    add_h1(doc, '三、面试授课设计 · 3 套候选', bookmark='_sec_lessons')
    add_quote(doc,
              'v1 授课设计（选 v2 模块 2.2 速读财报）已被 CEO 标为"太普通"。本次重做 3 套"炸点版"备选，'
              '每套独立文件、4 项齐全、面试现场不依赖 PPT。3 套差异化定位，覆盖不同面试场景。')

    add_h2(doc, '3.0  3 套对比（CEO 挑选）')
    rows = [
        ['套件', '主题', '炸点', '适合面试场景', '合规定位'],
        ['A 套', '10 分钟让 AI 给你 5 维红绿灯——"我想买 XX 票"决策推演沙盘', '让 AI 把"凭感觉下单"变成"5 维推演 + 红绿灯 + 3 种学员画像适配"；可让面试官"指定一只票重跑"',
         '产品运营 / Agent / 数据科学岗', 'AI 是"决策陪练"不是"决策者"，红绿灯 ≠ 买卖信号'],
        ['C 套', '5 分钟测出你的"投资人格"——AI 帮你匹配专属工具和策略', '让学员从"AI 课"变成"我的 AI 课"；可让面试官"答 5 道题"，整个面试里最易被记住',
         '产品运营 / 增长 / 用户画像 / 推荐算法岗', '人格评估是教育用途，不是心理测评'],
        ['D 套', '从一只票回溯 3 年——AI 带你做"主力建仓痕迹"逆向追踪', '市面 99% 散户课讲"怎么买"，几乎没人讲"主力怎么买"——机构投研思路平移到散户',
         'Data Scientist / 量化研究 / 风控 / 投研类岗', '逆向追踪是"看见过去"不是"预测未来"'],
    ]
    add_table_from_md(doc, ['| ' + ' | '.join(r) + ' |' for r in rows])

    # 3.1 A 套
    add_h2(doc, '3.1  A 套｜决策沙盘', bookmark='_sec_lesson_A')
    add_quote(doc,
              '课程主题：《10 分钟让 AI 给你 5 维红绿灯——"我想买 XX 票"决策推演沙盘》\n'
              '选取理由：每个散户最大痛点不是"没工具"，是"我为什么总在冲动时刻下手"——给一套"决策推演 SOP"，把"凭感觉下单"变成"5 维推演 + 红绿灯 + 综合判断"。\n'
              '适配选自 v2 模块 4.1（写 SOP）+ 模块 3.1（多维度梳理）+ 模块 1.1（5 大坑）三节混合。')
    render_markdown(doc, _read('授课设计_决策沙盘.md', start_marker='## 1. 课程主题', end_marker='## 4. AI 演示 prompt 套件'))

    # A 套 prompt 套件
    add_h3(doc, 'A 套 · AI 演示 prompt 套件（4 套）')
    render_markdown(doc, _read('授课设计_决策沙盘.md', start_marker='## 4. AI 演示 prompt 套件', end_marker='## 5. 卖点'))

    # 3.2 C 套
    add_h2(doc, '3.2  C 套｜投资人格', bookmark='_sec_lesson_C')
    add_quote(doc,
              '课程主题：《5 分钟测出你的"投资人格"——AI 帮你匹配专属工具和策略》\n'
              '选取理由：韭菜和老司机的最大分歧不是"用不用 AI"，是"用什么 AI / 怎么用 AI"——给学员一个自我评估 + 工具匹配 + 策略推荐的完整动作。\n'
              '适配选自 v2 模块 0（0 基础开箱）+ 模块 2（AI 工具盘点）+ 模块 1.1（散户 5 大坑）混合。')
    render_markdown(doc, _read('授课设计_投资人格.md', start_marker='## 1. 课程主题', end_marker='## 4. AI 演示 prompt 套件'))

    # C 套 prompt 套件
    add_h3(doc, 'C 套 · AI 演示 prompt 套件（4 套）')
    render_markdown(doc, _read('授课设计_投资人格.md', start_marker='## 4. AI 演示 prompt 套件', end_marker='## 5. 卖点'))

    # 3.3 D 套
    add_h2(doc, '3.3  D 套｜主力逆向', bookmark='_sec_lesson_D')
    add_quote(doc,
              '课程主题：《从一只票回溯 3 年——AI 带你做"主力建仓痕迹"逆向追踪》\n'
              '选取理由：市面 99% 散户课都讲"怎么买"，几乎没人讲"主力是怎么买的"——机构投研思路平移到散户的硬核玩法，差异化极强。\n'
              '适配选自 v2 模块 3.1（多维度梳理）+ 模块 6.1（AI 复盘官）+ 模块 2.1/2.2（工具栈）混合。')
    render_markdown(doc, _read('授课设计_主力逆向.md', start_marker='## 1. 课程主题', end_marker='## 4. AI 演示 prompt 套件'))

    # D 套 prompt 套件
    add_h3(doc, 'D 套 · AI 演示 prompt 套件（4 套）')
    render_markdown(doc, _read('授课设计_主力逆向.md', start_marker='## 4. AI 演示 prompt 套件', end_marker='## 5. 卖点'))


def section_assets(doc):
    """四、附：交付物清单"""
    add_h1(doc, '四、附：交付物清单', bookmark='_sec_assets')

    add_h2(doc, '4.1  完整 Markdown 交付物（v2 全量）')
    rows = [
        ['文件', '说明', '状态'],
        ['大纲_v1.md', '第一版课程大纲：6 模块 × 3 小节', 'v1 已交付'],
        ['大纲_v2.md', '最终版课程大纲：6 模块 × 18 小节 + 模块 0 + 2 附录', 'v2 已交付'],
        ['授课设计.md', 'v1 10-15 分钟面试授课设计：选 v2 模块 2.2 速读财报', 'v1 已交付'],
        ['授课设计_决策沙盘.md', 'A 套炸点版：5 维红绿灯决策推演沙盘', '已交付'],
        ['授课设计_投资人格.md', 'C 套炸点版：5 分钟测出投资人格 + AI 工具匹配', '已交付'],
        ['授课设计_主力逆向.md', 'D 套炸点版：3 年主力建仓逆向追踪', '已交付'],
        ['卖课话术.md', '50-100 字一句话卖课话术 + 内部 3 卖点拆解 + 风险红线', '已交付'],
        ['README.md', '项目说明 + 一句话总结', '已更新'],
        ['AI散户投资课程_完整交付.docx', '本 Word 文档（封面 / 目录 / 6 大主体 / 页眉页脚）', '已生成'],
    ]
    add_table_from_md(doc, ['| ' + ' | '.join(r) + ' |' for r in rows])

    add_h2(doc, '4.2  研究 / 审稿输入物（v2 反馈来源）')
    rows2 = [
        ['文件', '说明'],
        ['研究底料_手册.md', '研究员团队喂料：6 模块 × 3 视角 × 4 维度'],
        ['研究_审稿意见.md', '研究员对 v1 的审稿（4 必改 + 5 建议 + 3 加分）'],
        ['交易员_散户避坑.md', '5 大散户坑 × AI 救生圈'],
        ['交易员_模块案例.md', '13 个真实案例覆盖 6 模块'],
        ['交易员_审稿意见.md', '交易员对 v1 的审稿（7 必改 + 7 建议 + 6 加分）'],
        ['买方_机构对照.md', '机构 vs 散户 AI 用法 5 大对照'],
        ['买方_合规清单.md', '合规审稿（6 必改 + 12 建议 + 8 加分）'],
        ['买方_差异化建议.md', '3 条"市面少见"差异化建议'],
        ['散户_韭菜审稿.md', '韭菜视角审稿（必改 / 建议 / 加分）'],
        ['散户_优秀散户审稿.md', '优秀散户视角审稿'],
        ['散户_愿不愿买单.md', '散户代表团"30 秒买单决策"'],
        ['AI工具_评测矩阵.md', '12 款 AI 投资工具 × 6 列评测'],
        ['AI工具_推荐组合.md', '3 档预算工具组合'],
        ['AI工具_避坑指南.md', '5 条真实陷阱 + 公开案例'],
        ['AI工具_模块映射建议.md', '6 模块 × 工具强度矩阵'],
    ]
    add_table_from_md(doc, ['| ' + ' | '.join(r) + ' |' for r in rows2])

    add_h2(doc, '4.3  协作分工')
    rows3 = [
        ['角色', '负责'],
        ['research-team', '喂研究底料 + 研究入门内容（财报阅读、量化基础、行为金融等）'],
        ['trader', '补实战案例 + 散户 5 大坑 + 模拟盘 → 小仓位渐进式训练'],
        ['buy-side-council', '专业买方视角审稿（机构对照 / 合规 / 差异化 3 份）'],
        ['retail-panel', '散户代表团审稿（韭菜 + 优秀散户 + 愿不愿买单 3 份）'],
        ['ai-pm', '主流 AI 投资工具的评测与选型（评测 / 推荐 / 避坑 / 映射 4 份）'],
        ['course-designer', '统筹大纲 v1 → v2 → 授课设计 → 卖课话术 → Word 交付'],
    ]
    add_table_from_md(doc, ['| ' + ' | '.join(r) + ' |' for r in rows3])

    add_h2(doc, '4.4  6 条红线提醒（贯穿全课程）')
    add_bullet(doc, 'AI 是加速器不是信号源——重复 3 次以上。')
    add_bullet(doc, '数字必须查原始来源——重复 3 次以上。')
    add_bullet(doc, '回测不等于实盘——重复 3 次以上。')
    add_bullet(doc, '境内散户用富途/老虎 = 高风险（2023 年 5 月境内 APP 已下架）。')
    add_bullet(doc, '任何"年化 30%+ 稳赚"都是营销。')
    add_bullet(doc, '凯利公式不适用于散户（散户没有"无限次重复"前提）。')

    add_h2(doc, '4.5  文档元信息')
    add_bullet(doc, f'生成时间：{DATE_STR}')
    add_bullet(doc, '作者：方逸之')
    add_bullet(doc, '本课程定位为投资教育内容，不构成投资建议。')
    add_bullet(doc, 'AI 输出仅供研究参考，投资有风险，决策需谨慎。')


# ==============================================================
# Main
# ==============================================================
def main():
    print('=' * 60)
    print('AI 散户投资课程 · Word 完整交付生成')
    print('=' * 60)

    doc = init_document()
    print('[1/6] 文档骨架 + 页眉页脚 ✓')

    add_cover(doc)
    print('[2/6] 封面页 ✓')

    add_toc(doc)
    print('[3/6] 目录页 ✓')

    section_positioning(doc)
    add_page_break(doc)
    print('[4/6] 一、课程定位与卖课话术 ✓')

    section_outline(doc)
    add_page_break(doc)
    print('[5/6] 二、课程大纲 v2 ✓')

    section_lessons(doc)
    add_page_break(doc)
    print('[6/6] 三、面试授课设计 · 3 套候选 ✓')

    section_assets(doc)
    print('[+] 四、附：交付物清单 ✓')

    doc.save(OUT_FILE)
    print(f'\n输出：{OUT_FILE}')

    # 文件大小 + hash
    size = os.path.getsize(OUT_FILE)
    h = hashlib.sha256(open(OUT_FILE, 'rb').read()).hexdigest()
    print(f'大小：{size:,} bytes')
    print(f'SHA-256：{h}')

    return OUT_FILE, size, h


if __name__ == '__main__':
    main()
