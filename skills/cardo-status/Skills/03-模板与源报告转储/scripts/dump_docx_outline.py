#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dump_docx_outline.py -- 把 docx 结构转储为文本 / JSON（UTF-8 无 BOM）。

用法:
    python dump_docx_outline.py <docx> [--out <dir>] [--tables]

输出（默认写入 <docx 所在目录>/<stem>_dump/）:
    <stem>.blocks.txt    文档顺序的块清单：段落 "[样式] 文本"、表格 "[TABLE 行x列] header=[...]"
    <stem>.outline.json  标题 / 表格 / 图片块的索引 + level + text
    <stem>.styles.json   文档实际用到的样式 + 各节页面设置
    <stem>.tables.txt    每张表的单元格全文（仅 --tables）

设计要点:
    * 块索引 i 为 body 子元素的 0 基序号（空段落也占号），可稳定地跨工具引用。
    * 标题层级优先取样式名（Heading N / 标题 N），其次取直接 outlineLvl，
      再次沿 basedOn 链取样式里的 outlineLvl（0 基 -> 返回 1 基）。
    * 页面设置缺失的属性用 Word 默认值参与计算，但 JSON 里如实标注 inherited，
      不臆造模板里不存在的页边距。
退出码: 0 成功; 1 运行/写盘失败; 2 参数或输入错误。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

__version__ = "1.0.0"

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 通用候选：英文内置样式名 + 中文版内置样式名（不含任何项目/模板专有名词）
HEADING_STYLE_RE = re.compile(r"^(?:heading|标题)\s*([1-9])$", re.IGNORECASE)

# Word 内置默认值（twips）。仅用于「缺失时如何参与计算」，不会写回文档。
DEFAULT_PG_W = 11906
DEFAULT_PG_H = 16838
DEFAULT_MARGIN = 1440
TWIPS_PER_CM = 566.9291338582677


# --------------------------------------------------------------------------- #
# 基础工具
# --------------------------------------------------------------------------- #
def _silence_console() -> None:
    """避免在 cp936 控制台打印生僻字符时崩掉，不影响文件内容的 UTF-8。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass


def iter_blocks(doc):
    """按文档顺序产出 body 下的 (kind, obj)，kind in {paragraph, table}。"""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield "paragraph", Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield "table", Table(child, doc)


def para_style_name(p: Paragraph) -> str:
    try:
        name = p.style.name
    except Exception:
        name = ""
    return (name or "").strip()


def _outline_lvl_from_pPr(pPr) -> int | None:
    if pPr is None:
        return None
    el = pPr.find(qn("w:outlineLvl"))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val is None:
        return None
    try:
        iv = int(val)
    except (TypeError, ValueError):
        return None
    if iv >= 9:
        return None
    return iv + 1  # outlineLvl 0 基 -> 1 基


def _style_outline_level(style) -> int | None:
    """沿 basedOn 链找样式里继承来的 outlineLvl。"""
    seen: set[int] = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        try:
            pPr = style.element.find(qn("w:pPr"))
        except Exception:
            pPr = None
        lvl = _outline_lvl_from_pPr(pPr)
        if lvl is not None:
            return lvl
        try:
            style = style.base_style
        except Exception:
            return None
    return None


def heading_level(p: Paragraph) -> int:
    """返回 1 基标题层级；非标题返回 0。"""
    m = HEADING_STYLE_RE.match(para_style_name(p))
    if m:
        return int(m.group(1))
    try:
        pPr = p._p.pPr
    except Exception:
        pPr = None
    lvl = _outline_lvl_from_pPr(pPr)
    if lvl is not None:
        return lvl
    try:
        return _style_outline_level(p.style) or 0
    except Exception:
        return 0


def paragraph_images(p: Paragraph) -> list[dict]:
    """段落里的图片（blip / imagedata），解析成媒体文件名。"""
    rids: list[str] = []
    for el in p._p.iter():
        try:
            embed = el.get(qn("r:embed")) or el.get(qn("r:id"))
        except Exception:
            embed = None
        if not embed:
            continue
        tag = el.tag.rsplit("}", 1)[-1].lower()
        if tag in ("blip", "imagedata", "hlinkclick"):
            if tag == "hlinkclick":
                continue
            if embed not in rids:
                rids.append(embed)
    images = []
    for rid in rids:
        filename = None
        try:
            part = p.part.related_parts[rid]
            filename = str(part.partname)
        except Exception:
            filename = None
        images.append({"rid": rid, "part": filename})
    if not images:
        # 图形/图表/文本框等没有 blip 的 drawing，也记录数量
        drawings = sum(1 for el in p._p.iter() if el.tag == qn("w:drawing"))
        if drawings:
            images.append({"rid": None, "part": None, "drawings": drawings})
    return images


def _row_cells_text(row, dedupe: bool = False) -> list[str]:
    cells = []
    seen = set()
    for c in row.cells:
        key = id(c._tc)
        if dedupe:
            if key in seen:
                continue
            seen.add(key)
        cells.append(c.text.replace("\r", " ").replace("\n", " ").replace("\t", " ").strip())
    return cells


def table_grid_cols(t: Table) -> int:
    try:
        return len(t.columns)
    except Exception:
        try:
            return len(t.rows[0].cells) if len(t.rows) else 0
        except Exception:
            return 0


def table_header(t: Table) -> list[str]:
    if not len(t.rows):
        return []
    return [c[:40] for c in _row_cells_text(t.rows[0], dedupe=True)]


def table_style_name(t: Table) -> str:
    try:
        st = t.style
        return (st.name or "").strip() if st is not None else ""
    except Exception:
        return ""


def _get_twips(sectPr, child: str, attr: str, default: int):
    if sectPr is None:
        return default, True
    el = sectPr.find(qn(child))
    if el is None:
        return default, True
    raw = el.get(qn(attr))
    if raw is None:
        return default, True
    try:
        return int(raw), False
    except (TypeError, ValueError):
        return default, True


def section_info(doc) -> dict:
    """单节页面设置。缺失的 page size / margin 用内置默认值参与计算并标注 inherited。"""
    sectPr = None
    try:
        sectPr = doc.sections[0]._sectPr if doc.sections else None
    except Exception:
        sectPr = None
    if sectPr is None:
        sectPr = doc.element.body.find(qn("w:sectPr"))
    pw, pw_inherited = _get_twips(sectPr, "w:pgSz", "w:w", DEFAULT_PG_W)
    ph, ph_inherited = _get_twips(sectPr, "w:pgSz", "w:h", DEFAULT_PG_H)
    ml, ml_inh = _get_twips(sectPr, "w:pgMar", "w:left", DEFAULT_MARGIN)
    mr, mr_inh = _get_twips(sectPr, "w:pgMar", "w:right", DEFAULT_MARGIN)
    mt, mt_inh = _get_twips(sectPr, "w:pgMar", "w:top", DEFAULT_MARGIN)
    mb, mb_inh = _get_twips(sectPr, "w:pgMar", "w:bottom", DEFAULT_MARGIN)
    hdr, hdr_inh = _get_twips(sectPr, "w:pgMar", "w:header", 851)
    ftr, ftr_inh = _get_twips(sectPr, "w:pgMar", "w:footer", 992)
    gut, gut_inh = _get_twips(sectPr, "w:pgMar", "w:gutter", 0)
    usable_w = max(pw - ml - mr, 0)
    usable_h = max(ph - mt - mb, 0)
    orient = "LANDSCAPE" if pw > ph else "PORTRAIT"
    return {
        "page_w_cm": round(pw / TWIPS_PER_CM, 2),
        "page_h_cm": round(ph / TWIPS_PER_CM, 2),
        "margin_l_cm": round(ml / TWIPS_PER_CM, 2),
        "margin_r_cm": round(mr / TWIPS_PER_CM, 2),
        "margin_t_cm": round(mt / TWIPS_PER_CM, 2),
        "margin_b_cm": round(mb / TWIPS_PER_CM, 2),
        "usable_w_cm": round(usable_w / TWIPS_PER_CM, 2),
        "usable_h_cm": round(usable_h / TWIPS_PER_CM, 2),
        "orient": orient,
        # twips 原值：供渲染器精确复用
        "pg_w_twips": pw,
        "pg_h_twips": ph,
        "margin_l_twips": ml,
        "margin_r_twips": mr,
        "margin_t_twips": mt,
        "margin_b_twips": mb,
        "header_twips": hdr,
        "footer_twips": ftr,
        "gutter_twips": gut,
        "inherited": {
            "page_size": bool(pw_inherited or ph_inherited),
            "margins": bool(ml_inh or mr_inh or mt_inh or mb_inh),
            "header_footer": bool(hdr_inh or ftr_inh),
        },
    }


def all_sections(doc) -> list[dict]:
    out = []
    try:
        sections = list(doc.sections)
    except Exception:
        sections = []
    for idx, s in enumerate(sections):
        # 用当前节自己的 sectPr 计算
        try:
            sectPr = s._sectPr
            info = _section_info_from_sectPr(sectPr)
            info["header_refs"] = len(sectPr.findall(qn("w:headerReference")))
            info["footer_refs"] = len(sectPr.findall(qn("w:footerReference")))
        except Exception:
            info = section_info(doc)
        info["index"] = idx
        info["start_type"] = str(s.start_type).split(" (")[0] if s.start_type is not None else None
        out.append(info)
    if not out:
        out.append({**section_info(doc), "index": 0, "header_refs": 0, "footer_refs": 0})
    return out


def _section_info_from_sectPr(sectPr) -> dict:
    pw, pw_inherited = _get_twips(sectPr, "w:pgSz", "w:w", DEFAULT_PG_W)
    ph, ph_inherited = _get_twips(sectPr, "w:pgSz", "w:h", DEFAULT_PG_H)
    ml, ml_inh = _get_twips(sectPr, "w:pgMar", "w:left", DEFAULT_MARGIN)
    mr, mr_inh = _get_twips(sectPr, "w:pgMar", "w:right", DEFAULT_MARGIN)
    mt, mt_inh = _get_twips(sectPr, "w:pgMar", "w:top", DEFAULT_MARGIN)
    mb, mb_inh = _get_twips(sectPr, "w:pgMar", "w:bottom", DEFAULT_MARGIN)
    hdr, hdr_inh = _get_twips(sectPr, "w:pgMar", "w:header", 851)
    ftr, ftr_inh = _get_twips(sectPr, "w:pgMar", "w:footer", 992)
    gut, gut_inh = _get_twips(sectPr, "w:pgMar", "w:gutter", 0)
    return {
        "page_w_cm": round(pw / TWIPS_PER_CM, 2),
        "page_h_cm": round(ph / TWIPS_PER_CM, 2),
        "margin_l_cm": round(ml / TWIPS_PER_CM, 2),
        "margin_r_cm": round(mr / TWIPS_PER_CM, 2),
        "margin_t_cm": round(mt / TWIPS_PER_CM, 2),
        "margin_b_cm": round(mb / TWIPS_PER_CM, 2),
        "usable_w_cm": round(max(pw - ml - mr, 0) / TWIPS_PER_CM, 2),
        "usable_h_cm": round(max(ph - mt - mb, 0) / TWIPS_PER_CM, 2),
        "orient": "LANDSCAPE" if pw > ph else "PORTRAIT",
        "pg_w_twips": pw,
        "pg_h_twips": ph,
        "margin_l_twips": ml,
        "margin_r_twips": mr,
        "margin_t_twips": mt,
        "margin_b_twips": mb,
        "header_twips": hdr,
        "footer_twips": ftr,
        "gutter_twips": gut,
        "inherited": {
            "page_size": bool(pw_inherited or ph_inherited),
            "margins": bool(ml_inh or mr_inh or mt_inh or mb_inh),
            "header_footer": bool(hdr_inh or ftr_inh),
        },
    }


def doc_default_fonts(doc) -> dict:
    """docDefaults 里的字体（正文 eastAsia 的兜底线索之一）。"""
    out = {"ascii": None, "east_asia": None, "east_asia_theme": None}
    try:
        styles_el = doc.styles.element
    except Exception:
        return out
    rpr_default = styles_el.find(qn("w:docDefaults"))
    if rpr_default is None:
        return out
    rPr = rpr_default.find(qn("w:rPrDefault"))
    if rPr is None:
        return out
    rPr = rPr.find(qn("w:rPr"))
    if rPr is None:
        return out
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return out
    out["ascii"] = rFonts.get(qn("w:ascii"))
    out["east_asia"] = rFonts.get(qn("w:eastAsia"))
    out["east_asia_theme"] = rFonts.get(qn("w:eastAsiaTheme"))
    return out


def scan(doc) -> dict:
    """扫描正文，返回 blocks / outline / used styles / tables 明细。"""
    blocks: list[tuple[str, object]] = []  # (kind, obj) 文档顺序，含空段落
    used_para_styles: dict[str, int] = {}
    used_table_styles: dict[str, int] = {}
    outline: list[dict] = []
    idx = 0
    skipped = 0
    for kind, obj in iter_blocks(doc):
        if kind == "paragraph":
            p: Paragraph = obj  # type: ignore[assignment]
            text = p.text.strip()
            st = para_style_name(p)
            used_para_styles[st] = used_para_styles.get(st, 0) + 1
            imgs = paragraph_images(p)
            lvl = heading_level(p)
            if lvl:
                outline.append(
                    {"i": idx, "kind": "heading", "level": lvl, "text": text, "style": st}
                )
            if imgs:
                outline.append(
                    {
                        "i": idx,
                        "kind": "image",
                        "level": 0,
                        "text": text,
                        "images": imgs,
                        "count": sum(im.get("drawings", 1) for im in imgs),
                    }
                )
        else:
            t: Table = obj  # type: ignore[assignment]
            st = table_style_name(t)
            used_table_styles[st] = used_table_styles.get(st, 0) + 1
            outline.append(
                {
                    "i": idx,
                    "kind": "table",
                    "level": 0,
                    "text": "",
                    "rows": len(t.rows),
                    "cols": table_grid_cols(t),
                    "header": table_header(t),
                    "style": st,
                }
            )
        blocks.append((kind, obj))
        idx += 1
    # 其它 body 子元素（sdt 等）计数，便于发现漏读
    body_tags = {qn("w:p"), qn("w:tbl"), qn("w:sectPr")}
    for child in doc.element.body.iterchildren():
        if child.tag not in body_tags:
            skipped += 1
    return {
        "blocks": blocks,
        "outline": outline,
        "used_para_styles": used_para_styles,
        "used_table_styles": used_table_styles,
        "skipped_body_elements": skipped,
    }


def render_blocks_txt(info: dict) -> str:
    lines: list[str] = []
    for kind, obj in info["blocks"]:
        if kind == "paragraph":
            p: Paragraph = obj  # type: ignore[assignment]
            text = p.text.strip()
            st = para_style_name(p)
            imgs = paragraph_images(p)
            img_names = [im.get("part") or f"drawing x{im.get('drawings')}" for im in imgs]
            img_suffix = ""
            if img_names:
                shown = ", ".join(os.path.basename(n) if n else "" for n in img_names if n)
                img_suffix = "  <IMG" + (f" {shown}" if shown else "") + ">"
            lvl = heading_level(p)
            if lvl:
                lines.append(f"[{st}] {text}{img_suffix}")
            elif not text and not imgs:
                continue
            elif not text:
                lines.append(f"[IMAGE]{img_suffix}")
            else:
                lines.append(f"[{st}] {text}{img_suffix}")
        else:
            t: Table = obj  # type: ignore[assignment]
            rows = len(t.rows)
            cols = table_grid_cols(t)
            header = json.dumps(table_header(t), ensure_ascii=False)
            lines.append(f"[TABLE {rows}x{cols}] header={header}")
    return "\n".join(lines) + ("\n" if lines else "")


def render_tables_txt(info: dict) -> str:
    out: list[str] = []
    tno = 0
    for kind, obj in info["blocks"]:
        if kind != "table":
            continue
        t: Table = obj  # type: ignore[assignment]
        tno += 1
        st = table_style_name(t)
        out.append(
            f"===== TABLE {tno} ({len(t.rows)}x{table_grid_cols(t)}) style={st or '-'} ====="
        )
        for r_i, row in enumerate(t.rows):
            cells = _row_cells_text(row, dedupe=True)
            out.append(f"  R{r_i + 1}: " + " | ".join(cells))
        out.append("")
    return "\n".join(out)


def _write_text(path: str, content: str) -> None:
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)


def _write_json(path: str, obj) -> None:
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
        f.write("\n")


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="dump_docx_outline.py",
        description="转储 docx 的正文块 / 大纲 / 样式与页面设置（UTF-8 无 BOM）。",
    )
    ap.add_argument("docx", help="待转储的 .docx 文件")
    ap.add_argument("--out", default=None, help="输出目录（默认 <docx目录>/<stem>_dump）")
    ap.add_argument("--tables", action="store_true", help="额外输出 <stem>.tables.txt（单元格全文）")
    ap.add_argument("--version", action="version", version="%(prog)s " + __version__)
    return ap


def main(argv: list[str] | None = None) -> int:
    _silence_console()
    args = build_parser().parse_args(argv)
    src = os.path.abspath(args.docx)
    if not os.path.isfile(src):
        print(f"ERROR: 找不到 docx 文件: {args.docx}", file=sys.stderr)
        return 2
    if not src.lower().endswith(".docx"):
        print(f"ERROR: 只支持 .docx（不支持 .doc）: {args.docx}", file=sys.stderr)
        return 2

    outdir = args.out or os.path.join(os.path.dirname(src), os.path.splitext(os.path.basename(src))[0] + "_dump")
    os.makedirs(outdir, exist_ok=True)

    try:
        doc = Document(src)
        info = scan(doc)
        sections = all_sections(doc)
        defaults = doc_default_fonts(doc)

        base = os.path.splitext(os.path.basename(src))[0]
        blocks_path = os.path.join(outdir, base + ".blocks.txt")
        outline_path = os.path.join(outdir, base + ".outline.json")
        styles_path = os.path.join(outdir, base + ".styles.json")
        tables_path = os.path.join(outdir, base + ".tables.txt")

        _write_text(blocks_path, render_blocks_txt(info))
        _write_json(outline_path, info["outline"])
        _write_json(
            styles_path,
            {
                "source": args.docx,
                "used": info["used_para_styles"],
                "table_styles": info["used_table_styles"],
                "sections": sections,
                "defined": [s.name for s in doc.styles],
                "defined_table_styles": [
                    s.name for s in doc.styles if _is_table_style(s)
                ],
                "doc_defaults": defaults,
                "counts": {
                    "blocks": len(info["blocks"]),
                    "headings": sum(1 for o in info["outline"] if o["kind"] == "heading"),
                    "tables": sum(1 for o in info["outline"] if o["kind"] == "table"),
                    "images": sum(1 for o in info["outline"] if o["kind"] == "image"),
                    "skipped_body_elements": info["skipped_body_elements"],
                },
            },
        )
        if args.tables:
            _write_text(tables_path, render_tables_txt(info))
    except Exception as exc:  # noqa: BLE001 - 明确区分运行期失败与参数错误
        print(f"ERROR: 转储失败: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    counts = {
        "blocks": len(info["blocks"]),
        "headings": sum(1 for o in info["outline"] if o["kind"] == "heading"),
        "tables": sum(1 for o in info["outline"] if o["kind"] == "table"),
        "images": sum(1 for o in info["outline"] if o["kind"] == "image"),
    }
    print(f"OK {args.docx}")
    print(f"  blocks={counts['blocks']} headings={counts['headings']} "
          f"tables={counts['tables']} images={counts['images']}")
    print(f"  outdir={outdir}")
    print(f"  sections={len(sections)} used_styles={json.dumps(info['used_para_styles'], ensure_ascii=False)}")
    if info["used_table_styles"]:
        print(f"  table_styles={json.dumps(info['used_table_styles'], ensure_ascii=False)}")
    if info["skipped_body_elements"]:
        print(f"  WARN: 跳过 {info['skipped_body_elements']} 个非 p/tbl 的 body 元素（如 sdt）")
    print(f"  files: blocks.txt outline.json styles.json" + (" tables.txt" if args.tables else ""))
    return 0


def _is_table_style(style) -> bool:
    try:
        return int(style.type) == 3
    except Exception:
        return False


if __name__ == "__main__":
    sys.exit(main())