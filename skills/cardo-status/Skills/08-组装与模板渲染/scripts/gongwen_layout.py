# -*- coding: utf-8 -*-
"""gongwen_layout.py -- 把 GB/T 9704 版式注入任意 docx（后处理，不依赖 Office）

用法：
  python gongwen_layout.py <in.docx> <out.docx> [--bodypages 4]

做的事：
1. sectPr：A4 / 页边距 / docGrid 22 行 × 28 字 / 奇偶页脚引用
2. styles.xml：Normal / Heading1-3 / Title 的字体、字号、固定行距、首行缩进
   （源文件若用数字 styleId，先按"样式名→公文样式"映射改段落引用，见 remap_pstyles）
3. 正文段落去掉直接字体字号与行距（表格内保留，避免表内文字被放大）
4. 表格：全部行加 cantSplit，首行加 tblHeader
5. settings.xml：evenAndOddHeaders + updateFields
6. footer_gw1/2（正文节，单页右 / 双页左）：PAGE + SECTIONPAGES 域；
   footer_gw3/4（前置节，封面+目录）：空页脚

口径：版心 156×225mm；正文仿宋 16pt / 固定 29pt（580 twips，22 行/页）；
      每行 28 字 → docGrid charSpace=-4（156mm/28=15.793pt，16-15.793=0.207pt，
      charSpace 以 1/20 pt 计）。
      "共 M 页"按 D-01 修订用 SECTIONPAGES（正文节页数），不用 NUMPAGES（全册），
      避免末页出现"页码 4 / 共 5 页"。
"""
from __future__ import annotations

import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

PGW, PGH = 11906, 16838
MAR_TOP, MAR_BOT, MAR_LEFT, MAR_RIGHT = 2098, 1984, 1588, 1474
HEADER, FOOTER = 851, 1587
LINE_PITCH = 580      # 29pt
CHAR_SPACE = -4       # 15.793pt 字距
BODY_SZ = 32          # 16pt
TITLE_SZ = 44         # 22pt
IND_FIRST = 640       # 2 字符 = 32pt


def footer_xml(align: str, body_pages: int) -> str:
    ind = ('<w:ind w:right="320" w:left="0"/>' if align == 'right'
           else '<w:ind w:left="320" w:right="0"/>')
    rpr = ('<w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/>'
           '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>')
    big = ('<w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/>'
           '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>')

    def fld(name, cached, rp):
        return ('<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r>%s<w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
                '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r>'
                '<w:r>%s<w:t>%s</w:t></w:r>'
                '<w:r>%s<w:fldChar w:fldCharType="end"/></w:r>'
                % (rp, rp, name, rp, rp, cached, rp))

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:ftr xmlns:w="%s" xmlns:r="%s"><w:p><w:pPr>'
            '<w:pStyle w:val="Footer"/>%s<w:jc w:val="%s"/>%s</w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">— </w:t></w:r>%s'
            '<w:r>%s<w:t xml:space="preserve"> —　共 </w:t></w:r>%s'
            '<w:r>%s<w:t xml:space="preserve"> 页</w:t></w:r>'
            '</w:p></w:ftr>'
            % (W_NS, R_NS, ind, align, rpr, rpr, fld('PAGE', '1', rpr),
               rpr, fld('SECTIONPAGES', str(body_pages), big), rpr))


def empty_footer_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:ftr xmlns:w="{W_NS}" xmlns:r="{R_NS}">'
            '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p></w:ftr>')


def sectpr(ids: dict, kind: str = 'body') -> str:
    """kind='body' 正文节（PAGE/NUMPAGES、页码从 1 起）；'front' 前置节（封面+目录，无页码）。"""
    if kind == 'body':
        foot = (f'<w:footerReference w:type="default" r:id="{ids["body_odd"]}"/>'
                f'<w:footerReference w:type="even" r:id="{ids["body_even"]}"/>')
        head = ''
        pgnum = '<w:pgNumType w:start="1"/>'
    else:
        foot = (f'<w:footerReference w:type="default" r:id="{ids["front_odd"]}"/>'
                f'<w:footerReference w:type="even" r:id="{ids["front_even"]}"/>')
        head = '<w:type w:val="nextPage"/>'
        pgnum = ''
    return ('<w:sectPr>' + foot + head +
            f'<w:pgSz w:w="{PGW}" w:h="{PGH}"/>'
            f'<w:pgMar w:top="{MAR_TOP}" w:right="{MAR_RIGHT}" w:bottom="{MAR_BOT}" '
            f'w:left="{MAR_LEFT}" w:header="{HEADER}" w:footer="{FOOTER}" w:gutter="0"/>'
            + pgnum +
            '<w:cols w:space="720"/>'
            f'<w:docGrid w:type="linesAndChars" w:linePitch="{LINE_PITCH}" '
            f'w:charSpace="{CHAR_SPACE}"/></w:sectPr>')


def insert_front_section(xml: str, ids: dict) -> str:
    """把封面+目录划为前置节（不编页码），正文另起一节。

    节属性只能落在段落的 pPr 里；若正文首段之前紧邻的是表格（封面常为表），
    则按 Word 自身做法插一个 1 pt 空段承载节属性，避免把节属性写进单元格。
    """
    # 首个一级标题即正文起点；但"目录"若也套了 Heading1（指标说明即如此），
    # 它属于前置节，不能当正文起点，否则目录页会被编上页码。
    start = -1
    for pm in re.finditer(r'<w:p\b[^>]*>.*?</w:p>', xml, re.S):
        body = pm.group(0)
        if '<w:pStyle w:val="Heading1"/>' not in body:
            continue
        txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body))
        if txt.replace('\u3000', '').replace(' ', '') in ('目录', '目　录'):
            continue
        start = pm.start()
        break
    if start < 0:
        return xml
    sec = sectpr(ids, 'front')
    prev_end = xml.rfind('</w:p>', 0, start)
    if prev_end >= 0:
        pstart = max(xml.rfind('<w:p ', 0, prev_end), xml.rfind('<w:p>', 0, prev_end))
        head = xml[:pstart] if pstart >= 0 else ''
        if pstart >= 0 and head.count('<w:tbl>') == head.count('</w:tbl>'):
            pp = xml.find('<w:pPr>', pstart, prev_end)
            close = xml.find('</w:pPr>', pp) if pp >= 0 else -1
            if pp >= 0 and 0 <= close <= prev_end:
                return xml[:close] + sec + xml[close:]
    blank = ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" '
             'w:lineRule="exact"/>' + sec + '</w:pPr></w:p>')
    return xml[:start] + blank + xml[start:]


def strip_direct(chunk: str) -> str:
    chunk = re.sub(r'<w:rFonts\b[^>]*/>', '', chunk)
    chunk = re.sub(r'<w:sz\b[^>]*/>', '', chunk)
    chunk = re.sub(r'<w:szCs\b[^>]*/>', '', chunk)
    chunk = re.sub(r'<w:spacing\b[^>]*/>', '', chunk)
    chunk = re.sub(r'<w:ind\b[^>]*/>', '', chunk)
    return chunk


def process_document(xml: str, ids: dict) -> str:
    # 封面区（首个 Heading1 之前）保持原样；正文区清掉直接字体字号行距
    head = re.search(r'<w:pStyle w:val="Heading1"', xml)
    cut = (max(xml.rfind('<w:p ', 0, head.start()),
               xml.rfind('<w:p>', 0, head.start())) if head else 0)
    text, pos = [], 0
    for m in re.finditer(r'<w:tbl>.*?</w:tbl>', xml, re.S):
        seg = xml[pos:m.start()]
        if m.start() < cut:
            text.append(seg)
        else:
            text.append(strip_direct(seg))
        text.append(m.group(0))
        pos = m.end()
    tail = xml[pos:]
    text.append(tail if pos < cut else strip_direct(tail))
    xml = ''.join(text)

    # 表格行：cantSplit + 首行 tblHeader
    def fix_tbl(m):
        tbl = m.group(0)
        first = [True]

        def fix_tr(tm):
            tr = tm.group(0)
            add = []
            if '<w:cantSplit/>' not in tr:
                add.append('<w:cantSplit/>')
            if first[0] and '<w:tblHeader/>' not in tr:
                add.append('<w:tblHeader/>')
            first[0] = False
            add = ''.join(add)
            if not add:
                return tr
            if '<w:trPr>' in tr:
                tr = tr.replace('<w:trPr>', '<w:trPr>' + add, 1)
            elif '<w:trPr/>' in tr:
                tr = tr.replace('<w:trPr/>', '<w:trPr>' + add + '</w:trPr>', 1)
            else:
                tr = re.sub(r'(<w:tr\b[^>]*>)', r'\1<w:trPr>' + add + '</w:trPr>',
                            tr, count=1)
            return tr

        return re.sub(r'<w:tr\b.*?</w:tr>', fix_tr, tbl, flags=re.S)

    xml = re.sub(r'<w:tbl>.*?</w:tbl>', fix_tbl, xml, flags=re.S)

    # 替换最后一个 sectPr 为正文节（页码从 1 起）
    idx = list(re.finditer(r'<w:sectPr\b[^>]*>.*?</w:sectPr>', xml, re.S))
    if idx:
        m = idx[-1]
        xml = xml[:m.start()] + sectpr(ids, 'body') + xml[m.end():]
    else:
        xml = xml.replace('</w:body>', sectpr(ids, 'body') + '</w:body>', 1)
    return insert_front_section(xml, ids)


def style_el(sid, name, east, sz, line, ppr_extra, bold=False, keep=False):
    ascii_font = 'Times New Roman'
    rpr = (f'<w:rPr><w:rFonts w:ascii="{ascii_font}" w:eastAsia="{east}" '
           f'w:hAnsi="{ascii_font}" w:cs="{ascii_font}"/>'
           + ('<w:b/><w:bCs/>' if bold else '')
           + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>')
    keep_xml = '<w:keepNext/>' if keep else ''
    based = '' if sid == 'Normal' else '<w:basedOn w:val="Normal"/>'
    return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>'
            f'{based}<w:qFormat/>'
            f'<w:pPr>{keep_xml}<w:spacing w:before="0" w:after="0" '
            f'w:line="{line}" w:lineRule="exact"/>{ppr_extra}</w:pPr>{rpr}</w:style>')


HEAD_NAME = re.compile(r'^(?:heading|标题)\s*([1-4])$', re.I)
BODY_NAMES = {'Normal', 'Body Text', '正文', '正文文本', '首行缩进正文'}


def build_style_map(styles_xml: str) -> dict:
    """按样式名（而非 styleId）建 id → 公文样式 的映射。

    有的包（如现状分析报告 v1）段落引用的是数字 styleId（2/3/6/13），
    直接 upsert Normal/Heading1-3/Title 会打空——段落根本不指向这些 id。
    """
    smap = {}
    for m in re.finditer(r'<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>',
                         styles_xml, re.S):
        sid, body = m.group(1), m.group(2)
        nm = re.search(r'<w:name w:val="([^"]*)"', body)
        if not nm:
            continue
        name = nm.group(1).strip()
        hm = HEAD_NAME.match(name)
        if hm:
            smap[sid] = 'Heading' + hm.group(1)
        elif name == 'Title' or name == '标题':
            smap[sid] = 'Title'
        elif name in BODY_NAMES:
            smap[sid] = 'Normal'
    return {k: v for k, v in smap.items() if k != v}


def remap_pstyles(xml: str, smap: dict) -> str:
    """把段落里的 w:pStyle 引用改到公文样式 id（只改引用，不改样式定义）。"""
    if not smap:
        return xml

    def rep(m):
        val = m.group(1)
        return '<w:pStyle w:val="%s"/>' % smap.get(val, val)

    return re.sub(r'<w:pStyle w:val="([^"]+)"/>', rep, xml)


def toc_entries_xml(entries, max_level: int = 3) -> str:
    """TOC 域：begin + `TOC \\o "1-N" \\h \\z \\u` + separate + 回填条目 + end。

    条目为 (级别, 标题, 页码)，页码取同构 PDF 实测；`updateFields` 保持 false，
    打开即显示缓存结果，业主按 F9 可自行重算。N 取条目最大级别（建议书源件为
    1-2，若强行写 1-3 会多收 H3 条目、目录可能涨页）。
    """
    out = []
    for i, (lvl, text, page) in enumerate(entries):
        runs = ''
        if i == 0:
            runs += ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                     '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-%d" '
                     '\\h \\z \\u </w:instrText></w:r>'
                     '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' % max_level)
        esc = (text.replace('&', '&amp;').replace('<', '&lt;')
               .replace('>', '&gt;'))
        runs += ('<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
                 '<w:r><w:tab/></w:r><w:r><w:t>%d</w:t></w:r>' % (esc, page))
        if i == len(entries) - 1:
            runs += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        style = 'TOC1' if lvl == 1 else 'TOC2'
        out.append('<w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>%s</w:p>'
                   % (style, runs))
    return ''.join(out)


def insert_toc(xml: str, entries) -> str:
    """把"目　　录"标题后的空段/旧 TOC 缓存段换成新的 TOC 域；不动承载 sectPr 的空段。"""
    if not entries:
        return xml
    target = None
    for m in re.finditer(r'<w:p\b[^>]*>.*?</w:p>', xml, re.S):
        t = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', m.group(0)))
        if t.replace('\u3000', '').replace(' ', '') == '目录':
            target = m
            break
    if target is None:
        return xml
    pos = target.end()
    entry_keys = {re.sub(r'\s+', '', t) for _, t, _ in entries}
    drop = []
    kept_note = 0
    for m in re.finditer(r'<w:p\b[^>]*>.*?</w:p>', xml[pos:], re.S):
        body = m.group(0)
        txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body))
        norm = re.sub(r'\s+', '', txt)
        # 静态目录条目是"标题文字 + 页码"挤在一段里（如 '0 导读3'），
        # 页码与本次实测值未必相同，故先剔末位页码再比标题文字。
        key = re.sub(r'\d{1,3}$', '', norm)
        is_toc = re.search(r'<w:pStyle w:val="TOC[1-3]"/>', body) is not None
        is_static = (key in entry_keys and re.search(r'\d{1,3}$', norm) is not None)
        if (is_toc or is_static or txt.strip() == '') and '<w:sectPr>' not in body:
            drop.append((pos + m.start(), pos + m.end()))
            continue
        # 目录标题下可能有一句说明（如"目录页码以 PDF 为准"），原样保留，继续向后找旧条目
        if not drop and kept_note < 2 and txt.strip():
            kept_note += 1
            continue
        break
    max_level = max(lvl for lvl, _, _ in entries)
    toc = toc_entries_xml(entries, max_level)
    cut = [toc if i == 0 else '' for i in range(len(drop))]
    if not drop:
        return xml[:pos] + toc + xml[pos:]
    parts, last = [], 0
    for (s, e), rep in zip(drop, cut):
        parts.append(xml[last:s])
        parts.append(rep)
        last = e
    parts.append(xml[last:])
    return ''.join(parts)


def process_styles(xml: str) -> str:
    def upsert(style_id, name, east, sz, line, extra_ppr='', bold=False, keep=False):
        nonlocal xml
        new = style_el(style_id, name, east, sz, line, extra_ppr, bold, keep)
        pat = re.compile(r'<w:style\b[^>]*w:styleId="%s"[^>]*>.*?</w:style>' % style_id,
                         re.S)
        if pat.search(xml):
            xml = pat.sub(new, xml, count=1)
        else:
            xml = xml.replace('</w:styles>', new + '</w:styles>', 1)

    upsert('Normal', 'Normal', '仿宋', BODY_SZ, LINE_PITCH,
           f'<w:ind w:firstLine="{IND_FIRST}"/><w:jc w:val="both"/>')
    upsert('Heading1', 'heading 1', '黑体', BODY_SZ, LINE_PITCH,
           f'<w:ind w:firstLine="{IND_FIRST}"/><w:jc w:val="left"/>'
           '<w:outlineLvl w:val="0"/>', keep=True)
    upsert('Heading2', 'heading 2', '楷体', BODY_SZ, LINE_PITCH,
           f'<w:ind w:firstLine="{IND_FIRST}"/><w:jc w:val="left"/>'
           '<w:outlineLvl w:val="1"/>', keep=True)
    upsert('Heading3', 'heading 3', '仿宋', BODY_SZ, LINE_PITCH,
           f'<w:ind w:firstLine="{IND_FIRST}"/><w:jc w:val="left"/>'
           '<w:outlineLvl w:val="2"/>', bold=True, keep=True)
    upsert('Heading4', 'heading 4', '仿宋', BODY_SZ, LINE_PITCH,
           f'<w:ind w:firstLine="{IND_FIRST}"/><w:jc w:val="left"/>'
           '<w:outlineLvl w:val="3"/>', keep=True)
    upsert('Title', 'Title', '华文中宋', TITLE_SZ, 660,
           '<w:jc w:val="center"/><w:ind w:firstLine="0"/>')
    # 目录条目：三号仿宋、固定 29 磅、右制表位点线（8844 twips = 版心 156mm）
    for sid, name, left in (('TOC1', 'toc 1', 0), ('TOC2', 'toc 2', 640),
                            ('TOC3', 'toc 3', 1280)):
        new = (f'<w:style w:type="paragraph" w:styleId="{sid}">'
               f'<w:name w:val="{name}"/><w:basedOn w:val="Normal"/><w:qFormat/>'
               '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8844"/>'
               '</w:tabs><w:spacing w:before="0" w:after="0" w:line="580" '
               'w:lineRule="exact"/>'
               f'<w:ind w:left="{left}" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>'
               '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋" '
               'w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
               '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>')
        pat = re.compile(r'<w:style\b[^>]*w:styleId="%s"[^>]*>.*?</w:style>' % sid,
                         re.S)
        if pat.search(xml):
            xml = pat.sub(new, xml, count=1)
        else:
            xml = xml.replace('</w:styles>', new + '</w:styles>', 1)
    return xml


def process_settings(xml: str, update_fields: bool = True) -> str:
    val = 'true' if update_fields else 'false'
    if '<w:updateFields' in xml:
        xml = re.sub(r'<w:updateFields\b[^>]*/>',
                     f'<w:updateFields w:val="{val}"/>', xml, count=1)
    else:
        xml = re.sub(r'(<w:settings\b[^>]*>)',
                     r'\1<w:updateFields w:val="%s"/>' % val, xml, count=1)
    if '<w:evenAndOddHeaders' not in xml:
        xml = re.sub(r'(<w:settings\b[^>]*>)', r'\1<w:evenAndOddHeaders/>', xml, count=1)
    return xml


PARTS = {'body_odd': 'footer_gw1.xml', 'body_even': 'footer_gw2.xml',
         'front_odd': 'footer_gw3.xml', 'front_even': 'footer_gw4.xml'}


def _alloc_ids(rels_txt: str) -> dict:
    used = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_txt)]
    n = (max(used) if used else 0) + 1
    ids = {}
    for k in ('body_odd', 'body_even', 'front_odd', 'front_even'):
        ids[k] = f'rId{n}'
        n += 1
    return ids


def apply(src: Path, dst: Path, body_pages: int = 4, toc=None,
          update_fields: bool = True):
    """toc = [(级别, 标题, 页码)]；给了就插 TOC 域。update_fields 默认 True，
    插 TOC 域时由调用方传 False（打开不弹更新、直接显示缓存结果）。"""
    rels_name = 'word/_rels/document.xml.rels'
    tmp = Path(tempfile.mkdtemp()) / 'out.docx'
    with zipfile.ZipFile(src) as zin:
        ids = _alloc_ids(zin.read(rels_name).decode('utf-8'))
        smap = build_style_map(zin.read('word/styles.xml').decode('utf-8'))
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                xml = remap_pstyles(data.decode('utf-8'), smap)
                # 先插目录再分节：若反过来，插分节时会把 sectPr 挂到目录区最后一个
                # 旧条目上，insert_toc 为避免破坏节属性只能保留它，导致旧条目残留。
                if toc:
                    xml = insert_toc(xml, toc)
                xml = process_document(xml, ids)
                data = xml.encode('utf-8')
            elif item.filename == 'word/styles.xml':
                data = process_styles(data.decode('utf-8')).encode('utf-8')
            elif item.filename == 'word/settings.xml':
                data = process_settings(
                    data.decode('utf-8'),
                    update_fields=update_fields and not toc).encode('utf-8')
            elif re.match(r'word/footer\d*\.xml$', item.filename):
                data = empty_footer_xml().encode('utf-8')
            elif re.match(r'word/header\d*\.xml$', item.filename):
                data = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        f'<w:hdr xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:p/></w:hdr>'
                        ).encode('utf-8')
            elif item.filename == rels_name:
                txt = data.decode('utf-8')
                add = ''.join(
                    f'<Relationship Id="{ids[k]}" Type="{R_NS}/footer" '
                    f'Target="{PARTS[k]}"/>'
                    for k in ('body_odd', 'body_even', 'front_odd', 'front_even'))
                data = txt.replace('</Relationships>', add + '</Relationships>'
                                   ).encode('utf-8')
            elif item.filename == '[Content_Types].xml':
                txt = data.decode('utf-8')
                add = ''.join(
                    f'<Override PartName="/word/{PARTS[k]}" ContentType='
                    '"application/vnd.openxmlformats-officedocument.'
                    'wordprocessingml.footer+xml"/>'
                    for k in ('body_odd', 'body_even', 'front_odd', 'front_even'))
                data = txt.replace('</Types>', add + '</Types>').encode('utf-8')
            zout.writestr(item, data)
        zout.writestr('word/' + PARTS['body_odd'], footer_xml('right', body_pages))
        zout.writestr('word/' + PARTS['body_even'], footer_xml('left', body_pages))
        zout.writestr('word/' + PARTS['front_odd'], empty_footer_xml())
        zout.writestr('word/' + PARTS['front_even'], empty_footer_xml())
    shutil.move(str(tmp), str(dst))
    return dst


if __name__ == '__main__':
    src, dst = Path(sys.argv[1]), Path(sys.argv[2])
    body_pages = 4
    toc = None
    if '--toc' in sys.argv:
        import json
        with open(sys.argv[sys.argv.index('--toc') + 1], encoding='utf-8') as f:
            data = json.load(f)
        entries = data.get('entries') if isinstance(data, dict) else data
        toc = [(int(x['level']), x['text'], int(x['page']))
               for x in entries or [] if str(x.get('page', '')).strip() != '']
        if isinstance(data, dict) and data.get('bodyPages'):
            body_pages = int(data['bodyPages'])
    for flag in ('--bodypages', '--total'):
        if flag in sys.argv:
            body_pages = int(sys.argv[sys.argv.index(flag) + 1])
    update_fields = '--noupdate' not in sys.argv
    apply(src, dst, body_pages, toc, update_fields)
    print('written', dst, 'toc entries', len(toc) if toc else 0,
          'bodyPages', body_pages,
          'updateFields', update_fields and not toc)
