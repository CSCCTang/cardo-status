#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_report_docx.py -- 以模板 docx 为基底，按 markdown 重建正文（通用、零硬编码）。

用法:
    python build_report_docx.py --template <tpl.docx> --content <content.md>
                                --out <out.docx> [--style-map <yaml>] [--title <str>]
                                [--force]

行为:
    * 打开模板 docx，只保留 body 末尾的 sectPr（纸张 / 页边距 / 页眉页脚引用），
      删除正文里的段落与表格，再按 markdown 重建；模板样式表原样保留。
    * markdown 约定: "# " -> h1, "## " -> h2, "### " -> h3（更深的标题并入 h3），
      普通行 -> 正文样式；|---|---| 分隔行可省略；--- 按配置忽略或分页；
      以 （数据来源：…） 开头的行按「来源」段落处理（默认同正文样式）。
    * 样式名不硬编码: 默认按候选表从模板探测（Heading N / 标题 N、Body Text / 正文、
      Title 等），--style-map 可覆盖；缺失 h2/h3 自动回退到上一级，正文回退到模板默认段落样式。
    * 中文字体: 正文 / 来源 / 表格单元格 run 显式写 w:eastAsia（默认从模板同样式继承，
      取不到则用「宋体」）；标题不覆盖模板字体。
    * 表格: 沿用模板表格样式（无则生成 Table Grid 兜底），列宽按单元格内容显示宽度
      按比例分配（固定布局，单列占比限制在 8%~55%）、首行跨页重复；可选首行加粗 / 底纹。
    * 自检: 生成后用 python-docx 重新打开回读，比对样式分布 / 表格形状 / 页面设置 /
      页眉页脚引用 / eastAsia 字体，打印 [verify] ok|FAIL 明细；FAIL 时退出码 1。

退出码: 0 成功; 1 运行失败或自检 FAIL; 2 参数 / 输入 / 配置错误。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import unicodedata

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Pt, RGBColor, Twips
from docx.table import Table
from docx.text.paragraph import Paragraph

__version__ = "1.0.0"

# --------------------------------------------------------------------------- #
# 通用候选表（英文内置样式名 + 中文版内置样式名，不含任何项目专有名词）
# --------------------------------------------------------------------------- #
STYLE_CANDIDATES: dict[str, list[str]] = {
    "title": ["Title", "标题", "文档标题", "Title Text"],
    "h1": ["Heading 1", "标题 1", "Heading1"],
    "h2": ["Heading 2", "标题 2", "Heading2"],
    "h3": ["Heading 3", "标题 3", "Heading3"],
    "body": ["Body Text", "正文", "正文文本", "Normal", "普通"],
    "source": ["Source", "数据来源", "Bibliography", "引用"],
}
STYLE_ROLES = ("title", "h1", "h2", "h3", "body", "source")

CONFIG_DEFAULTS: dict[str, str] = {
    "page_break_on_hr": "false",
    "table_header_bold": "true",
    "table_header_shading": "",
    "table_header_repeat": "true",
    "source_color": "",
    "source_size_delta": "0",
    "east_asia_font": "",
    "source_regex": "",
}
CONFIG_KEYS = tuple(CONFIG_DEFAULTS)
ALLOWED_MAP_KEYS = STYLE_ROLES + ("table",) + CONFIG_KEYS

DEFAULT_EAST_ASIA = "宋体"
DEFAULT_SOURCE_RE = (
    r"^\s*[（(]\s*数据\s*(?:来源|出处|说明|引用)\s*[:：].*[)）]\s*$"
)
DEFAULT_TABLE_STYLE_NAMES = (
    "normal table",
    "table normal",
    "普通表格",
    "normal table 1",
    "table normal1",
)

TWIPS_PER_CM = 566.9291338582677
CJK_TWIPS = 210.0  # 约 10.5pt 全角字宽（twips），用于把"显示宽度"折算成列宽下限
DEFAULT_PG_W = 11906
DEFAULT_PG_H = 16838
DEFAULT_MARGIN = 1440

HEADING_RE = re.compile(r"^(#{1,6})[ \t]+(.*)$")
IMAGE_RE = re.compile(r"^!\[([^\]]*)\]\(\s*([^)\s]+)(?:\s+\"[^\"]*\")?\s*\)$")
HR_RE = re.compile(r"^([-*_])(?:[ \t]*\1){2,}$")
SEP_CELL_RE = re.compile(r"^:?-{3,}:?$")
HEX_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


class BuildError(Exception):
    """参数 / 输入 / 配置错误（退出码 2）。"""


# --------------------------------------------------------------------------- #
# 控制台与配置
# --------------------------------------------------------------------------- #
def _silence_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass


def _norm(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip()).lower()


def _strip_inline_comment(value: str) -> str:
    quote = None
    for i, ch in enumerate(value):
        if quote:
            if ch == quote:
                quote = None
        elif ch in ("'", '"'):
            quote = ch
        elif ch == "#" and (i == 0 or value[i - 1] in " \t"):
            return value[:i].rstrip()
    return value.rstrip()


def load_style_map(path: str | None) -> dict[str, str]:
    """读取扁平 key: value 的 YAML（无 pyyaml 依赖）；.json 文件也可。"""
    if not path:
        return {}
    if not os.path.isfile(path):
        raise BuildError(f"--style-map 文件不存在: {path}")
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            text = f.read()
    except OSError as exc:
        raise BuildError(f"读取 --style-map 失败: {exc}") from exc

    if path.lower().endswith(".json"):
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise BuildError(f"--style-map JSON 不合法: {exc}") from exc
        if not isinstance(data, dict):
            raise BuildError("--style-map JSON 顶层必须是对象")
        return {str(k): "" if v is None else str(v) for k, v in data.items()}

    result: dict[str, str] = {}
    for lineno, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("-"):
            raise BuildError(
                f"--style-map 第 {lineno} 行: 只支持扁平 'key: value'，不支持列表"
            )
        m = re.match(r"^([A-Za-z_][\w.-]*)\s*:\s*(.*)$", line)
        if not m:
            raise BuildError(
                f"--style-map 第 {lineno} 行不是 'key: value' 形式: {raw.strip()!r}"
            )
        key, val = m.group(1), _strip_inline_comment(m.group(2).strip())
        if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
            val = val[1:-1]
        result[key] = val
    return result


def parse_bool(value: str, key: str) -> bool:
    v = (value or "").strip().lower()
    if v in ("true", "yes", "y", "1", "on"):
        return True
    if v in ("false", "no", "n", "0", "off", ""):
        return False
    raise BuildError(f"配置 {key} 需要布尔值（true/false），收到: {value!r}")


def resolve_config(raw: dict[str, str]) -> dict:
    """把 style-map 里的字符串配置解析成结构化配置。"""
    cfg: dict = {}
    cfg["page_break_on_hr"] = parse_bool(
        raw.get("page_break_on_hr", CONFIG_DEFAULTS["page_break_on_hr"]), "page_break_on_hr"
    )
    cfg["table_header_bold"] = parse_bool(
        raw.get("table_header_bold", CONFIG_DEFAULTS["table_header_bold"]), "table_header_bold"
    )
    cfg["table_header_repeat"] = parse_bool(
        raw.get("table_header_repeat", CONFIG_DEFAULTS["table_header_repeat"]),
        "table_header_repeat",
    )
    shading = raw.get("table_header_shading", "").strip().lstrip("#")
    if shading and not HEX_RE.match(shading):
        raise BuildError(f"配置 table_header_shading 需要 6 位十六进制颜色，收到: {raw.get('table_header_shading')!r}")
    cfg["table_header_shading"] = shading.upper()

    color = raw.get("source_color", "").strip().lstrip("#")
    if color and not HEX_RE.match(color):
        raise BuildError(f"配置 source_color 需要 6 位十六进制颜色，收到: {raw.get('source_color')!r}")
    cfg["source_color"] = color.upper()

    delta_raw = raw.get("source_size_delta", "0").strip() or "0"
    try:
        cfg["source_size_delta"] = float(delta_raw)
    except ValueError as exc:
        raise BuildError(f"配置 source_size_delta 需要数字（磅），收到: {delta_raw!r}") from exc

    cfg["east_asia_font"] = raw.get("east_asia_font", "").strip()
    regex = raw.get("source_regex", "").strip()
    if regex:
        try:
            cfg["source_re"] = re.compile(regex)
        except re.error as exc:
            raise BuildError(f"配置 source_regex 不是合法正则: {exc}") from exc
    else:
        cfg["source_re"] = re.compile(DEFAULT_SOURCE_RE)
    return cfg


# --------------------------------------------------------------------------- #
# docx 读取 / 结构工具
# --------------------------------------------------------------------------- #
def iter_blocks(doc):
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


def style_index(doc) -> dict[str, object]:
    idx: dict[str, object] = {}
    for st in doc.styles:
        try:
            nm = st.name
        except Exception:
            continue
        if nm:
            idx.setdefault(_norm(nm), st)
    return idx


def style_type(style) -> int:
    try:
        return int(style.type)
    except Exception:
        return -1


def _font_from_style(style) -> str | None:
    """沿 basedOn 链取样式里显式声明的 w:eastAsia。"""
    seen: set[int] = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        try:
            rPr = style.element.find(qn("w:rPr"))
        except Exception:
            rPr = None
        if rPr is not None:
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is not None:
                ea = rFonts.get(qn("w:eastAsia"))
                if ea:
                    return ea
        try:
            style = style.base_style
        except Exception:
            return None
    return None


def _size_from_style(style) -> float | None:
    """样式链上的字号（磅）。"""
    seen: set[int] = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        try:
            rPr = style.element.find(qn("w:rPr"))
        except Exception:
            rPr = None
        if rPr is not None:
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                try:
                    return int(sz.get(qn("w:val"))) / 2.0
                except (TypeError, ValueError):
                    pass
        try:
            style = style.base_style
        except Exception:
            return None
    return None


def doc_default_east_asia(doc) -> str | None:
    try:
        styles_el = doc.styles.element
    except Exception:
        return None
    defaults = styles_el.find(qn("w:docDefaults"))
    if defaults is None:
        return None
    rPr = defaults.find(qn("w:rPrDefault"))
    if rPr is None:
        return None
    rPr = rPr.find(qn("w:rPr"))
    if rPr is None:
        return None
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return None
    return rFonts.get(qn("w:eastAsia"))


def default_para_style(doc):
    try:
        st = doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
        if st is not None:
            return st
    except Exception:
        pass
    idx = style_index(doc)
    return idx.get(_norm("Normal"))


# --------------------------------------------------------------------------- #
# 样式解析
# --------------------------------------------------------------------------- #
def _pick_style(doc, role: str, override: str | None, idx: dict[str, object]):
    """返回 (style|None, 来源说明)。override 为空串表示显式禁用该角色。"""
    want = WD_STYLE_TYPE.PARAGRAPH
    if override is not None:
        if not override.strip():
            return None, "disabled"
        key = _norm(override)
        st = idx.get(key)
        if st is None:
            avail = sorted(
                s.name for s in doc.styles if style_type(s) == want and getattr(s, "name", None)
            )
            raise BuildError(
                f"--style-map {role}: 模板中不存在段落样式 {override!r}。"
                f"可用段落样式: {', '.join(avail) or '(无)'}"
            )
        if style_type(st) != want:
            raise BuildError(f"--style-map {role}: 样式 {override!r} 不是段落样式")
        return st, "mapped"
    for cand in STYLE_CANDIDATES.get(role, []):
        st = idx.get(_norm(cand))
        if st is not None and style_type(st) == want:
            return st, "template"
    return None, "missing"


def _is_default_table_style(name: str) -> bool:
    return _norm(name) in {_norm(n) for n in DEFAULT_TABLE_STYLE_NAMES}


def ensure_table_grid(doc):
    idx = style_index(doc)
    st = idx.get(_norm("Table Grid"))
    if st is not None and style_type(st) == WD_STYLE_TYPE.TABLE:
        return st
    try:
        st = doc.styles.add_style("Table Grid", WD_STYLE_TYPE.TABLE)
    except Exception as exc:  # noqa: BLE001
        raise BuildError(f"无法创建兜底表格样式 Table Grid: {exc}") from exc
    normal_tbl = idx.get(_norm("Normal Table")) or idx.get(_norm("Table Normal"))
    if normal_tbl is not None and style_type(normal_tbl) == WD_STYLE_TYPE.TABLE:
        try:
            st.base_style = normal_tbl
        except Exception:
            pass
    el = st.element
    tblPr = el.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        el.append(tblPr)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    tblPr.append(borders)
    return st


def resolve_table_style(doc, override: str | None, idx: dict[str, object]):
    want = WD_STYLE_TYPE.TABLE
    if override is not None:
        if not override.strip():
            return None, "disabled"
        st = idx.get(_norm(override))
        if st is None:
            # 通用兜底表样式名允许「模板没有则现场生成」，避免常见映射直接失败
            if _norm(override) in {"table grid", "tablegrid", "网格", "网格型"}:
                return ensure_table_grid(doc), "mapped:generated"
            avail = sorted(
                s.name for s in doc.styles if style_type(s) == want and getattr(s, "name", None)
            )
            raise BuildError(
                f"--style-map table: 模板中不存在表格样式 {override!r}。"
                f"可用表格样式: {', '.join(avail) or '(无)'}"
            )
        if style_type(st) != want:
            raise BuildError(f"--style-map table: 样式 {override!r} 不是表格样式")
        return st, "mapped"
    for st in doc.styles:
        try:
            nm = st.name
        except Exception:
            continue
        if not nm or style_type(st) != want:
            continue
        if _is_default_table_style(nm):
            continue
        return st, "template"
    return ensure_table_grid(doc), "generated"


def resolve_styles(doc, style_map: dict[str, str]):
    idx = style_index(doc)
    resolved: dict[str, object] = {}
    origins: dict[str, str] = {}
    for role in STYLE_ROLES:
        st, how = _pick_style(doc, role, style_map.get(role), idx)
        resolved[role] = st
        origins[role] = how

    fallback_notes: list[str] = []
    if resolved["body"] is None:
        resolved["body"] = default_para_style(doc)
        origins["body"] = "default"
        if resolved["body"] is None:
            raise BuildError("模板中没有可用作正文的段落样式（Body Text / 正文 / Normal 均缺失）")
    if resolved["h1"] is None:
        resolved["h1"] = resolved["body"]
        origins["h1"] = "fallback:body"
        fallback_notes.append("h1 缺失，回退正文样式")
    if resolved["h2"] is None:
        resolved["h2"] = resolved["h1"]
        origins["h2"] = "fallback:h1"
        fallback_notes.append("h2 缺失，回退 h1 样式")
    if resolved["h3"] is None:
        resolved["h3"] = resolved["h2"]
        origins["h3"] = "fallback:h2"
        fallback_notes.append("h3 缺失，回退 h2 样式")
    if resolved["source"] is None:
        resolved["source"] = resolved["body"]
        origins["source"] = "fallback:body"
    if resolved["title"] is None:
        resolved["title"] = resolved["h1"]
        origins["title"] = "fallback:h1"
    table_style, table_origin = resolve_table_style(doc, style_map.get("table"), idx)

    ear = style_map.get("east_asia_font", "").strip()
    if not ear:
        ear = _font_from_style(resolved["body"]) or doc_default_east_asia(doc) or DEFAULT_EAST_ASIA

    # 来源段落字号基准：来源样式链 > 正文字号 > 12pt
    base_pt = _size_from_style(resolved["source"]) or _size_from_style(resolved["body"]) or 12.0

    return {
        "styles": resolved,
        "origins": origins,
        "table_style": table_style,
        "table_origin": table_origin,
        "east_asia": ear,
        "source_base_pt": base_pt,
        "notes": fallback_notes,
    }


# --------------------------------------------------------------------------- #
# markdown 解析
# --------------------------------------------------------------------------- #
def split_table_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    cells = re.split(r"(?<!\\)\|", s)
    return [c.strip().replace("\\|", "|") for c in cells]


def is_separator_row(cells: list[str]) -> bool:
    if not cells:
        return False
    return all(SEP_CELL_RE.match(c.strip()) for c in cells if c.strip()) and any(
        c.strip() for c in cells
    )


def _clean_heading_text(text: str) -> str:
    text = re.sub(r"\s+#+\s*$", "", text.strip())
    return text.replace("**", "").strip()


def parse_markdown(text: str, cfg: dict) -> tuple[list[dict], list[str]]:
    blocks: list[dict] = []
    warnings: list[str] = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    i = 0
    n = len(lines)
    while i < n:
        raw = lines[i]
        line = raw.strip()
        if not line:
            i += 1
            continue

        # HTML 注释（含切章脚本写入的图片占位）整段跳过
        if line.startswith("<!--"):
            while i < n and "-->" not in lines[i]:
                i += 1
            i += 1
            continue

        m = HEADING_RE.match(line)
        if m:
            level = len(m.group(1))
            text_ = _clean_heading_text(m.group(2))
            if not text_:
                warnings.append(f"第 {i + 1} 行是空标题，已忽略")
            else:
                blocks.append({"kind": "heading", "level": level, "text": text_})
            i += 1
            continue

        m = IMAGE_RE.match(line)
        if m:
            # 独立成行的图片：![图注](相对路径)。路径相对 --content 所在目录解析，
            # 图片块本身不占正文段落，图注仍由紧随其后的来源行承担。
            blocks.append({"kind": "image", "alt": m.group(1).strip(), "path": m.group(2).strip()})
            i += 1
            continue

        if HR_RE.match(line):
            if cfg["page_break_on_hr"]:
                blocks.append({"kind": "pagebreak"})
            i += 1
            continue

        if line.startswith("|"):
            rows: list[list[str]] = []
            while i < n and lines[i].strip().startswith("|"):
                cells = split_table_row(lines[i])
                if not is_separator_row(cells):
                    rows.append(cells)
                i += 1
            if not rows:
                warnings.append(f"第 {i} 行附近有空的管道表，已忽略")
                continue
            ncols = max(len(r) for r in rows)
            for r in rows:
                if len(r) < ncols:
                    r.extend([""] * (ncols - len(r)))
            blocks.append(
                {
                    "kind": "table",
                    "header": rows[0],
                    "rows": rows[1:],
                    "cols": ncols,
                }
            )
            continue

        if cfg["source_re"].match(line):
            blocks.append({"kind": "source", "text": line})
        else:
            blocks.append({"kind": "para", "text": line})
        i += 1
    return blocks, warnings


# --------------------------------------------------------------------------- #
# 构建
# --------------------------------------------------------------------------- #
def clear_body(doc) -> int:
    body = doc.element.body
    removed = 0
    for child in list(body.iterchildren()):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)
        removed += 1
    return removed


def _sectPr(doc):
    try:
        if len(doc.sections):
            return doc.sections[-1]._sectPr
    except Exception:
        pass
    return doc.element.body.find(qn("w:sectPr"))


def _twips(sectPr, child: str, attr: str, default: int) -> int:
    if sectPr is None:
        return default
    el = sectPr.find(qn(child))
    if el is None:
        return default
    raw = el.get(qn(attr))
    if raw is None:
        return default
    try:
        return int(raw)
    except (TypeError, ValueError):
        return default


def usable_width_twips(doc) -> int:
    sectPr = _sectPr(doc)
    pw = _twips(sectPr, "w:pgSz", "w:w", DEFAULT_PG_W)
    ml = _twips(sectPr, "w:pgMar", "w:left", DEFAULT_MARGIN)
    mr = _twips(sectPr, "w:pgMar", "w:right", DEFAULT_MARGIN)
    return max(pw - ml - mr, 1440)


def sect_snapshot(doc) -> dict:
    sectPr = _sectPr(doc)
    snap: dict = {
        "pgSz_present": False,
        "pgSz_attrs": {},
        "pgMar_present": False,
        "pgMar_attrs": {},
        "header_refs": 0,
        "footer_refs": 0,
    }
    if sectPr is None:
        return snap
    pgSz = sectPr.find(qn("w:pgSz"))
    if pgSz is not None:
        snap["pgSz_present"] = True
        snap["pgSz_attrs"] = {k.split("}")[-1]: v for k, v in pgSz.attrib.items()}
    pgMar = sectPr.find(qn("w:pgMar"))
    if pgMar is not None:
        snap["pgMar_present"] = True
        snap["pgMar_attrs"] = {k.split("}")[-1]: v for k, v in pgMar.attrib.items()}
    snap["header_refs"] = len(sectPr.findall(qn("w:headerReference")))
    snap["footer_refs"] = len(sectPr.findall(qn("w:footerReference")))
    return snap


def set_east_asia(run, font: str) -> None:
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rStyle = rPr.find(qn("w:rStyle"))
        rPr.insert(1 if rStyle is not None else 0, rFonts)
    rFonts.set(qn("w:eastAsia"), font)


def add_text_runs(par, text: str, *, east_asia: str | None = None, bold_all: bool = False,
                  color: str | None = None, size_pt: float | None = None) -> None:
    """支持 **粗体** 的最小行内解析；表格单元格里用 \\n 表示换行。"""
    segments = str(text).split("\n")
    for s_i, segment in enumerate(segments):
        for tok in re.split(r"(\*\*.+?\*\*)", segment):
            if not tok:
                continue
            bold = False
            body = tok
            if tok.startswith("**") and tok.endswith("**") and len(tok) >= 4:
                bold, body = True, tok[2:-2]
            if not body:
                continue
            run = par.add_run(body)
            if bold or bold_all:
                run.bold = True
            if east_asia:
                set_east_asia(run, east_asia)
            if color:
                run.font.color.rgb = RGBColor.from_string(color)
            if size_pt is not None:
                run.font.size = Pt(size_pt)
        if s_i < len(segments) - 1:
            tail = par.add_run()
            tail.add_break(WD_BREAK.LINE)


def set_table_width(table, width_twips: int) -> None:
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        successors = (
            "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
            "w:tblLayout", "w:tblCellMar", "w:tblLook",
        )
        anchor = None
        for tag in successors:
            anchor = tblPr.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(tblW)
        else:
            tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(width_twips)))


def shade_cell(cell, fill: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        successors = (
            "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark",
        )
        anchor = None
        for tag in successors:
            anchor = tcPr.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(shd)
        else:
            tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def mark_header_row(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    trPr.append(el)


def add_paragraph_block(doc, block: dict, resolved: dict, cfg: dict) -> Paragraph:
    role = "body"
    if block["kind"] == "heading":
        role = {1: "h1", 2: "h2"}.get(block["level"], "h3")
    elif block["kind"] == "source":
        role = "source"
    style = resolved["styles"][role]
    p = doc.add_paragraph(style=style)
    text = block["text"]
    if block["kind"] in ("heading", "title"):
        # 标题不覆盖模板字体：不加 eastAsia 直接格式
        p.add_run(text)
    elif block["kind"] == "source":
        color = cfg["source_color"] or None
        size = None
        delta = cfg["source_size_delta"]
        if delta:
            size = max(resolved["source_base_pt"] + delta, 1.0)
        add_text_runs(p, text, east_asia=resolved["east_asia"], color=color, size_pt=size)
    else:
        add_text_runs(p, text, east_asia=resolved["east_asia"])
    return p


# 表头固定列宽下限（twips）：1cm ≈ 567twips。
# 序号列 1.2cm、问题类别列 3.5cm，避免表头被挤成逐字竖排或断成两行。
HEADER_MIN_TW = {"序号": 680, "编号": 680, "问题类别": 1984}


def _disp_width(text: str) -> int:
    total = 0
    for ch in text or "":
        total += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return total


def compute_table_widths(usable: int, header: list, rows: list, ncols: int) -> list:
    """按内容显示宽度分配列宽（固定比例，非均分）。

    做法分两步，避免"长文本列被挤成窄条"和"短列被拉得过宽"两种极端：
      1) 给每列一个**下限**：取该列 90 分位内容宽度折算成 twips（全角字约 210twips），
         限制在可用宽度的 6% ~ 20%（列少时 25%）；
      2) 余量按 85 分位权重的比例分配；最后单列上限 55%，总额等于可用宽度。

    中文表格里"文件名称""缺项内容"这类列需要宽度，而"序号""类型"只占几个字；
    均分会让前者反复折行、后者留白，所以按内容分配。
    """
    if ncols <= 0:
        return []
    p85s: list = []
    needs: list = []
    for c in range(ncols):
        lens = [_disp_width(header[c] if c < len(header) else "")]
        for r in rows:
            lens.append(_disp_width(r[c] if c < len(r) else ""))
        lens.sort()
        n = len(lens)
        p85s.append(float(max(lens[int(round(0.85 * (n - 1)))], 4)))
        p90 = lens[int(round(0.90 * (n - 1)))]
        head = _disp_width(header[c] if c < len(header) else "")
        needs.append(float(max(p90, head, 3)))

    cap_ratio = 0.25 if ncols <= 4 else 0.20
    floors = [min(max(w * CJK_TWIPS, usable * 0.06), usable * cap_ratio) for w in needs]
    total_floor = sum(floors)
    if total_floor >= usable:
        widths = [usable * w / (sum(p85s) or 1.0) for w in p85s]
    else:
        remainder = usable - total_floor
        tot_w = sum(p85s) or 1.0
        widths = [f + remainder * w / tot_w for f, w in zip(floors, p85s)]
    hi = usable * 0.55
    widths = [min(w, hi) for w in widths]
    s = sum(widths)
    if s > 0 and abs(s - usable) > 0.5:
        widths = [w * usable / s for w in widths]
    ints = [int(round(w)) for w in widths]
    # 最小可读宽度兜底：表头/单元格里的最长不可断 ASCII 串与两个字宽，
    # 避免「序号」「置信度」这类表头被挤成逐字竖排（与确认单、缺口清单同一条规则）。
    pad, cjk = 300, 9.0 * 20
    mins = []
    for c in range(ncols):
        texts = [str(header[c]) if c < len(header) else ""]
        texts += [str(r[c]) for r in rows if c < len(r)]
        tokens = [tok for txt in texts for tok in re.findall(r"[A-Za-z0-9_.%\-/\\]+", txt)]
        ascii_w = max([len(t) for t in tokens] or [1]) * 9.0 * 10
        mins.append(int(max(ascii_w, 2 * cjk) + pad))
    # 固定列宽下限：序号列 1.2cm、问题类别列 3.5cm（A4 版心约 8300twips）。
    # 均分或纯按内容分配会把「序号」挤成竖排「序/号」、「问题类别」断成「问题类/别」。
    for c in range(ncols):
        head_txt = str(header[c]).strip() if c < len(header) else ""
        floor_tw = HEADER_MIN_TW.get(head_txt)
        if floor_tw:
            mins[c] = max(mins[c], floor_tw)
    ints = [max(a, m) for a, m in zip(ints, mins)]
    overflow = sum(ints) - usable
    while overflow > 0:
        cand = [(ints[i] - mins[i], i) for i in range(ncols) if ints[i] > mins[i]]
        if not cand:
            break
        _, idx = max(cand)
        take = min(overflow, ints[idx] - mins[idx])
        ints[idx] -= take
        overflow -= take
    ints[-1] += usable - sum(ints)
    return [max(i, 1) for i in ints]


def add_table_block(doc, block: dict, resolved: dict, cfg: dict) -> Table:
    ncols = block["cols"]
    header = block["header"]
    rows = block["rows"]
    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    style = resolved["table_style"]
    if style is not None:
        try:
            table.style = style
        except Exception as exc:  # noqa: BLE001
            raise BuildError(f"套用表格样式失败: {exc}") from exc
    table.autofit = False
    usable = usable_width_twips(doc)
    widths = compute_table_widths(usable, header, rows, ncols)
    set_table_width(table, usable)
    for c in range(ncols):
        table.columns[c].width = Twips(widths[c])

    for r_idx, row_data in enumerate([header] + rows):
        row = table.rows[r_idx]
        is_header = r_idx == 0
        for c_idx in range(ncols):
            cell = row.cells[c_idx]
            try:
                cell.width = Twips(widths[c_idx])
            except Exception:
                pass
            text = row_data[c_idx] if c_idx < len(row_data) else ""
            para = cell.paragraphs[0]
            add_text_runs(
                para,
                text,
                east_asia=resolved["east_asia"],
                bold_all=bool(is_header and cfg["table_header_bold"]),
            )
            if is_header and cfg["table_header_shading"]:
                shade_cell(cell, cfg["table_header_shading"])
        if is_header and cfg["table_header_repeat"]:
            mark_header_row(row)
    return table


def add_image_block(doc, block: dict, resolved: dict, cfg: dict) -> None:
    """图片块：按版心宽度等比缩放，居中，不随正文首行缩进。"""
    path = block["path"]
    if not os.path.isabs(path):
        path = os.path.join(cfg.get("image_base") or "", path)
    if not os.path.isfile(path):
        raise BuildError(f"图片不存在: {block['path']}（解析为 {path}）")
    p = doc.add_paragraph(style=resolved["styles"]["body"])
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Twips(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    width_tw = usable_width_twips(doc)
    run.add_picture(path, width=Emu(int(width_tw * 635)))


def _add_field(paragraph, instr: str) -> None:
    """在段落里插一个简单域（PAGE / NUMPAGES / TOC），Word 打开后按域计算。"""
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r = OxmlElement("w:r")
    r.append(OxmlElement("w:t"))
    fld.append(r)
    paragraph._p.append(fld)


def setup_footer(doc) -> None:
    """页脚写「第 PAGE 页 / 共 NUMPAGES 页」真域：docx 单独交付时也有页码。"""
    sec = doc.sections[-1]
    footer = sec.footer
    paras = list(footer.paragraphs)
    for extra in paras[1:]:
        extra._element.getparent().remove(extra._element)
    p = paras[0] if paras else footer.add_paragraph()
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for s in ("第 ",):
        run = p.add_run(s)
        run.font.size = Pt(9)
    _add_field(p, " PAGE ")
    run = p.add_run(" 页 / 共 ")
    run.font.size = Pt(9)
    _add_field(p, " NUMPAGES ")
    run = p.add_run(" 页")
    run.font.size = Pt(9)


def enable_title_page(doc) -> None:
    """封面不印页码：只置「首页不同」，不新建页脚部件（避免多出 footer2.xml）。"""
    for sec in doc.sections:
        sec.different_first_page_header_footer = True


def enable_update_fields(doc) -> None:
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")


def add_cover_and_toc(doc, resolved: dict, title: str, cover_note: str) -> None:
    """封面（标题 + 编制时间）+ 目录域。正文样式行数由 verify 计入，不靠空段充数。"""
    if cover_note:
        p = doc.add_paragraph(style=resolved["styles"]["body"])
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Twips(0)
        run = p.add_run(cover_note)
        set_east_asia(run, DEFAULT_EAST_ASIA)
    brk = doc.add_paragraph()
    brk.add_run().add_break(WD_BREAK.PAGE)
    head = doc.add_paragraph(style=resolved["styles"]["body"])
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    head.paragraph_format.first_line_indent = Twips(0)
    run = head.add_run("目　　录")
    run.bold = True
    set_east_asia(run, DEFAULT_EAST_ASIA)
    toc = doc.add_paragraph()
    toc.paragraph_format.first_line_indent = Twips(0)
    # TOC 域本身不带缓存文本：Word 打开即按 updateFields 重算，
    # 不在 docx 里写死一份会过期的静态目录。
    _add_field(toc, ' TOC \\o "1-2" \\h \\z \\u ')
    brk2 = doc.add_paragraph()
    brk2.add_run().add_break(WD_BREAK.PAGE)


def build_document(template_doc, blocks: list[dict], resolved: dict, cfg: dict,
                   title: str | None, cover_note: str = "") -> tuple[object, dict]:
    """在已打开并完成样式解析的模板文档上重建正文（样式对象必须来自同一个 doc）。"""
    doc = template_doc
    if len(doc.sections) > 1:
        print(
            f"WARN: 模板含 {len(doc.sections)} 节，重建后只保留最后一节页面设置（简单化处理）",
            file=sys.stderr,
        )
    removed = clear_body(doc)
    stats = {"removed_template_blocks": removed, "paragraphs": 0, "headings": 0,
             "tables": 0, "sources": 0, "pagebreaks": 0, "title": bool(title)}
    pending_break = False
    last_kind = None

    if title:
        p = doc.add_paragraph(style=resolved["styles"]["title"])
        run = p.add_run(title)
        # 封面标题按可用宽度缩字号：长项目名在默认字号下会把「报告」二字挤到下一行，
        # 看起来像折行事故。等比缩到一行放得下（下限 16pt），不拆词、不换行。
        try:
            sec = doc.sections[-1]
            # 模板可能不写页边距（继承默认值）：None 时按 Word 默认 2.54cm 计，
            # 否则整段异常被吞、字号不缩，「报告」二字照旧折行。
            lm = sec.left_margin if sec.left_margin is not None else Twips(DEFAULT_MARGIN)
            rm = sec.right_margin if sec.right_margin is not None else Twips(DEFAULT_MARGIN)
            usable_pt = (sec.page_width - lm - rm) / 12700.0
            size = run.font.size.pt if run.font.size else 26.0
            unit = sum(1.0 if unicodedata.east_asian_width(ch) in ("W", "F") else 0.55
                       for ch in title)
            # 留 12% 余量：按 1em 估算会贴边，实际字距/字体宽度会把最后几个字挤下去。
            # 模板页边距可能是 3.17cm（可用 415pt），按 0.88 缩到 16.5pt 仍不折行。
            if unit * size > usable_pt and unit > 0:
                run.font.size = Pt(max(14.0, usable_pt * 0.88 / unit))
        except Exception:
            pass
        stats["paragraphs"] += 1
        add_cover_and_toc(doc, resolved, title, cover_note)
        stats["cover"] = True
        stats["tocField"] = True

    for block in blocks:
        kind = block["kind"]
        if kind == "pagebreak":
            pending_break = True
            stats["pagebreaks"] += 1
            continue
        if kind == "table":
            if pending_break:
                bridge = doc.add_paragraph()
                bridge.paragraph_format.page_break_before = True
                stats["paragraphs"] += 1
                pending_break = False
            add_table_block(doc, block, resolved, cfg)
            stats["tables"] += 1
            last_kind = "table"
            continue
        if kind == "image":
            if pending_break:
                bridge = doc.add_paragraph()
                bridge.paragraph_format.page_break_before = True
                stats["paragraphs"] += 1
                pending_break = False
            add_image_block(doc, block, resolved, cfg)
            stats["images"] = stats.get("images", 0) + 1
            last_kind = "image"
            continue
        p = add_paragraph_block(doc, block, resolved, cfg)
        if pending_break:
            p.paragraph_format.page_break_before = True
            pending_break = False
        if kind == "heading":
            # 标题跟随下段：避免标题落在页末成为孤行（正文从下一页才开始）。
            p.paragraph_format.keep_with_next = True
            stats["headings"] += 1
        elif kind == "source":
            stats["sources"] += 1
        else:
            stats["paragraphs"] += 1
        last_kind = kind

    if last_kind == "table":
        # Word 惯例：表格后补一个空段落（自检时忽略空段落，不影响计数）。
        # 但默认字号的空段在表格刚好压到版心底部时会溢出到下一页，产出一整页空白
        # （报告末页曾出现 0 填充的 p15）。这里压成 1pt 固定行高、零段距，
        # 既保留「正文以段落收尾」，又不会把末页撑出来。
        tail = doc.add_paragraph()
        tail.paragraph_format.space_before = Pt(0)
        tail.paragraph_format.space_after = Pt(0)
        tail.paragraph_format.line_spacing = Pt(1)
    return doc, stats


def save_docx(doc, out_path: str, force: bool) -> None:
    if os.path.exists(out_path) and not force:
        raise BuildError(f"输出已存在: {out_path}（如需覆盖请加 --force）")
    out_dir = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(out_dir, exist_ok=True)
    tmp = out_path + f".tmp-{os.getpid()}"
    try:
        doc.save(tmp)
        os.replace(tmp, out_path)
    except Exception as exc:  # noqa: BLE001
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        raise BuildError(f"保存 docx 失败: {exc}") from exc


# --------------------------------------------------------------------------- #
# 自检
# --------------------------------------------------------------------------- #
def _expected_style_counts(blocks: list[dict], resolved: dict, cfg: dict) -> dict[str, int]:
    names = {role: resolved["styles"][role].name for role in STYLE_ROLES}
    counts: dict[str, int] = {}
    for block in blocks:
        kind = block["kind"]
        if kind in ("pagebreak", "table", "image"):
            continue
        if kind == "heading":
            level = block["level"]
            role = {1: "h1", 2: "h2"}.get(level, "h3")
        elif kind == "source":
            role = "source"
        else:
            role = "body"
        nm = names[role]
        counts[nm] = counts.get(nm, 0) + 1
    return counts


def verify(out_path: str, template: str, blocks: list[dict], resolved: dict, cfg: dict,
           title: str | None, expected_table_shapes: list[tuple[int, int]],
           cover_note: str = "") -> tuple[bool, list[str]]:
    lines: list[str] = []
    fails: list[str] = []

    def ok(msg: str) -> None:
        lines.append(f"[verify] ok   {msg}")

    def bad(msg: str) -> None:
        lines.append(f"[verify] FAIL {msg}")
        fails.append(msg)

    try:
        doc = Document(out_path)
        tpl = Document(template)
    except Exception as exc:  # noqa: BLE001
        return False, [f"[verify] FAIL 无法重新打开生成的 docx: {type(exc).__name__}: {exc}"]

    # 1) 段落样式分布
    expected = _expected_style_counts(blocks, resolved, cfg)
    if title:
        tname = resolved["styles"]["title"].name
        expected[tname] = expected.get(tname, 0) + 1
        # 封面编制时间行 + 目录标题行都用正文样式，按实际生成内容计入期望值
        bname = resolved["styles"]["body"].name
        expected[bname] = expected.get(bname, 0) + 1 + (1 if cover_note else 0)
    actual: dict[str, int] = {}
    body_paras: list[Paragraph] = []
    body_names = {resolved["styles"]["body"].name, resolved["styles"]["source"].name}
    for kind, obj in iter_blocks(doc):
        if kind != "paragraph":
            continue
        p: Paragraph = obj  # type: ignore[assignment]
        if not p.text.strip():
            continue
        nm = para_style_name(p)
        actual[nm] = actual.get(nm, 0) + 1
        if nm in body_names:
            body_paras.append(p)
    if actual == expected:
        detail = ", ".join(f"{k}={v}" for k, v in sorted(actual.items()))
        ok(f"paragraph_styles: {detail}")
    else:
        bad(
            "paragraph_styles 不匹配: expected "
            f"{json.dumps(expected, ensure_ascii=False, sort_keys=True)} / actual "
            f"{json.dumps(actual, ensure_ascii=False, sort_keys=True)}"
        )

    # 2) 标题数量（按角色）
    def count_style(name: str) -> int:
        return actual.get(name, 0)

    exp_h = {1: 0, 2: 0, 3: 0}
    for block in blocks:
        if block["kind"] == "heading":
            exp_h[min(block["level"], 3)] += 1
    act_h = {
        1: count_style(resolved["styles"]["h1"].name),
        2: count_style(resolved["styles"]["h2"].name),
        3: count_style(resolved["styles"]["h3"].name),
    }
    # 样式回退可能把多级映射到同一名字，此时按「至少」口径校验总数
    if resolved["styles"]["h1"].name == resolved["styles"]["h2"].name == resolved["styles"]["h3"].name:
        total_exp, total_act = sum(exp_h.values()), sum(act_h.values())
        if total_act == total_exp:
            ok(f"headings: {total_act} (各级样式合并为 {resolved['styles']['h1'].name})")
        else:
            bad(f"headings 总数不匹配: expected {total_exp} / actual {total_act}")
    elif act_h == exp_h:
        ok(f"headings: h1={act_h[1]} h2={act_h[2]} h3={act_h[3]}")
    else:
        bad(f"headings 不匹配: expected {exp_h} / actual {act_h}")

    # 3) 表格
    actual_tables: list[Table] = [obj for kind, obj in iter_blocks(doc) if kind == "table"]
    if len(actual_tables) == len(expected_table_shapes):
        ok(f"tables: {len(actual_tables)}")
    else:
        bad(f"tables 数量不匹配: expected {len(expected_table_shapes)} / actual {len(actual_tables)}")
    shapes = [(len(t.rows), len(t.columns)) for t in actual_tables]
    if shapes == expected_table_shapes:
        if shapes:
            ok("table_shapes: " + ", ".join(f"{r}x{c}" for r, c in shapes))
    else:
        bad(f"table_shapes 不匹配: expected {expected_table_shapes} / actual {shapes}")
    if resolved["table_style"] is not None:
        want = resolved["table_style"].name
        got = []
        for t in actual_tables:
            try:
                got.append(t.style.name if t.style is not None else None)
            except Exception:
                got.append(None)
        if got and all(g == want for g in got):
            ok(f"table_style: {want}")
        elif got:
            bad(f"table_style 不匹配: expected {want} / actual {got}")

    # 4) 标题段落
    if title:
        tname = resolved["styles"]["title"].name
        hit = [p for kind, p in iter_blocks(doc) if kind == "paragraph" and p.text.strip() == title]
        if hit and para_style_name(hit[0]) == tname:
            ok(f"title: 「{title}」({tname})")
        else:
            bad(f"title 缺失或样式不符: 期望「{title}」/「{tname}」")

    # 5) 页面设置与页眉页脚引用（原样保留，不臆造页边距）
    def fmt_snap(s: dict) -> str:
        parts = []
        parts.append("pgSz=" + (str(s["pgSz_attrs"]) if s["pgSz_present"] else "inherited"))
        parts.append("pgMar=" + (str(s["pgMar_attrs"]) if s["pgMar_present"] else "inherited"))
        parts.append(f"hdr={s['header_refs']} ftr={s['footer_refs']}")
        return " ".join(parts)

    snap_out = sect_snapshot(doc)
    snap_tpl = sect_snapshot(tpl)
    if snap_out == snap_tpl:
        ok(f"page_setup: {fmt_snap(snap_out)}")
    else:
        bad(f"page_setup 与模板不一致: template={fmt_snap(snap_tpl)} / out={fmt_snap(snap_out)}")

    # 6) eastAsia 字体
    want_ea = resolved["east_asia"]
    total_runs = 0
    missing = 0
    for p in body_paras:
        for run in p.runs:
            if not run.text.strip():
                continue
            total_runs += 1
            rPr = run._element.find(qn("w:rPr"))
            rFonts = rPr.find(qn("w:rFonts")) if rPr is not None else None
            ea = rFonts.get(qn("w:eastAsia")) if rFonts is not None else None
            if ea != want_ea:
                missing += 1
    if total_runs and missing == 0:
        ok(f"east_asia: {total_runs}/{total_runs} 正文 run 显式 w:eastAsia={want_ea}")
    elif total_runs == 0:
        bad("east_asia: 没有找到正文 run，无法校验")
    else:
        bad(f"east_asia: {missing}/{total_runs} 个正文 run 未按 {want_ea} 设置")

    # 7) 样式表仍在
    missing_styles = [
        role
        for role in STYLE_ROLES
        if resolved["styles"][role] is not None
        and resolved["styles"][role].name not in [s.name for s in doc.styles]
    ]
    if not missing_styles:
        ok(f"styles_preserved: {', '.join(sorted({resolved['styles'][r].name for r in STYLE_ROLES}))}")
    else:
        bad(f"模板样式缺失: {missing_styles}")

    ok("done") if not fails else None
    return (not fails), lines


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="build_report_docx.py",
        description="以模板 docx 为基底，按 markdown 重建正文并自检（UTF-8 无 BOM）。",
    )
    ap.add_argument("--template", required=True, help="模板 .docx（保留其页面设置、样式表、页眉页脚）")
    ap.add_argument("--content", required=True, help="markdown 内容文件（UTF-8）")
    ap.add_argument("--out", required=True, help="输出 .docx 路径")
    ap.add_argument("--style-map", default=None,
                    help="样式映射 YAML（title/h1/h2/h3/body/source/table 及可选配置项）")
    ap.add_argument("--title", default=None, help="文档标题（用 Title 样式落在第一段）")
    ap.add_argument("--cover-note", default="", help="封面副行（如编制时间），留空则不加")
    ap.add_argument("--force", action="store_true", help="输出已存在时覆盖")
    ap.add_argument("--version", action="version", version="%(prog)s " + __version__)
    return ap


def main(argv: list[str] | None = None) -> int:
    _silence_console()
    args = build_parser().parse_args(argv)

    template = os.path.abspath(args.template)
    content = os.path.abspath(args.content)
    out = os.path.abspath(args.out)

    if not os.path.isfile(template):
        print(f"ERROR: 模板不存在: {args.template}", file=sys.stderr)
        return 2
    if not template.lower().endswith(".docx"):
        print(f"ERROR: 模板必须是 .docx: {args.template}", file=sys.stderr)
        return 2
    if not os.path.isfile(content):
        print(f"ERROR: 内容文件不存在: {args.content}", file=sys.stderr)
        return 2
    if os.path.normcase(template) == os.path.normcase(out):
        print("ERROR: --out 不能与 --template 相同（会破坏模板）", file=sys.stderr)
        return 2

    try:
        style_map = load_style_map(args.style_map)
        unknown = [k for k in style_map if k not in ALLOWED_MAP_KEYS]
        if unknown:
            raise BuildError(
                f"--style-map 含未知键: {unknown}；允许的键: {', '.join(ALLOWED_MAP_KEYS)}"
            )
        cfg = resolve_config(style_map)

        try:
            with open(content, "r", encoding="utf-8-sig") as f:
                md_text = f.read()
        except OSError as exc:
            raise BuildError(f"读取内容文件失败: {exc}") from exc
        if not md_text.strip():
            raise BuildError(f"内容文件为空: {content}")

        # 图片相对路径的基准目录：--content 所在目录（正文.md 与 图表/ 同级）
        cfg["image_base"] = os.path.dirname(os.path.abspath(content))
        blocks, md_warnings = parse_markdown(md_text, cfg)
        if not blocks:
            raise BuildError("markdown 解析后没有任何内容块")

        doc = Document(template)
        resolved = resolve_styles(doc, style_map)
        for note in resolved["notes"]:
            print(f"WARN: {note}", file=sys.stderr)
        for w in md_warnings:
            print(f"WARN: {w}", file=sys.stderr)

        doc, stats = build_document(doc, blocks, resolved, cfg, args.title, args.cover_note)
        setup_footer(doc)
        enable_title_page(doc)
        enable_update_fields(doc)
        save_docx(doc, out, args.force)
    except BuildError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: 构建失败: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1

    expected_shapes = [(1 + len(b["rows"]), b["cols"]) for b in blocks if b["kind"] == "table"]

    names = {role: resolved["styles"][role].name for role in STYLE_ROLES}
    print(f"OK build {out}")
    print(f"  template={args.template}")
    print(
        "  styles="
        + " ".join(f"{role}:{names[role]}({resolved['origins'][role]})" for role in STYLE_ROLES)
    )
    print(
        f"  table_style={resolved['table_style'].name if resolved['table_style'] else '-'}"
        f"({resolved['table_origin']})  east_asia={resolved['east_asia']}"
    )
    print(
        f"  built: content_blocks={len(blocks)} paragraphs={stats['paragraphs']} "
        f"headings={stats['headings']} tables={stats['tables']} sources={stats['sources']} "
        f"pagebreaks={stats['pagebreaks']} removed_template_blocks={stats['removed_template_blocks']}"
    )

    ok, vlines = verify(out, template, blocks, resolved, cfg, args.title, expected_shapes,
                        args.cover_note)
    for line in vlines:
        print(line)
    if ok:
        print("[verify] ok")
        return 0
    print(f"[verify] FAIL ({len([l for l in vlines if 'FAIL' in l])} 项)")
    return 1


if __name__ == "__main__":
    sys.exit(main())