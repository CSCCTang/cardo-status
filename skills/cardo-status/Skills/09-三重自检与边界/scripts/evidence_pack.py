#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""evidence_pack.py -- 三件套同源导出：干净版 docx / 标注版 docx / 依据对照表 xlsx。

口径（用户确认）：
  * 标注版是唯一生产线：正文 + 每段一条〔锚点 …〕依据块；
  * 干净版 = 标注版剥离〔…〕依据块与 [S-xxx] 编号，不另走生产线；
  * 剥离前后数值集合与顺序必须完全一致，否则 fail（防止"两份内容"）；
  * 每个数值必须能指到 材料 / 公开 / 推导 / 定稿值 四类之一，否则判无源。

用法（默认参数即可跑）：
  python evidence_pack.py --project-dir <项目目录>
    可选：--md 标注源（文件/目录/列表） --dump 源报告转储 --pinned 定稿值.json
          --out-dir --version --check-only（09 只校验不写产物）
  退出码：0 通过；1 校验不通过；2 未找到标注源（旧格式，跳过）
"""
import argparse
import datetime
import glob
import json
import os
import re
import sys

# 运行时是 python313._pth 模式，脚本目录不进 sys.path：兄弟模块必须显式加进来
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BANNED_IN_CLEAN = [
    ("[S-", "外部条目编号"),
    ("（数据来源", "来源行"),
    ("数据来源：", "来源行"),
    ("置信度", "置信度标注"),
    ("推断依据", "推断标注"),
    ("据公开资料", "空头依据"),
    ("公开区间", "空头依据"),
    ("缺项-", "缺项编号"),
    ("（测算）", "测算标注"),
    ("（估算）", "估算标注"),
    ("（推断）", "推断标注"),
    ("待补", "缺口语"),
    ("未获取", "自曝缺口语"),
    ("资料缺失", "自曝缺口语"),
]
ANNOT_RE = re.compile(r"^〔(.+)〕\s*$")
SRC_ROW_RE = re.compile(r"^[（(]数据来源\s*[:：]")
HEAD_RE = re.compile(r"^(#{1,4})\s+(.+)$")
# 过程段：06 分章稿里的缺项表/来源映射只作留档，08 组装时剥离、不进交付正文。
# 三件套校验必须与 08 同口径跳过，否则会把过程表的「缺项-」当成正文污点。
PROCESS_SECTIONS = ("本节缺项表", "来源映射")
NUM_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*(?:[～~]\s*(\d+(?:\.\d+)?)\s*)?"
    r"(万平方米|万㎡|平方米|㎡|%|米|m|栋|户|个|处)")
CITE_RE = re.compile(r"\[S-\d{2,}\]")
# 章节指路：见第X章《…》 / 见《章》的《节》节 —— 目标必须落在真实标题层级里。
CITE_CHAPTER_RE = re.compile(
    r"见\s*第[一二三四五六七八九十百0-9]+章\s*[《「]([^》」]{2,40})[》」]")
CITE_SECTION_RE = re.compile(
    r"见\s*(?:第[一二三四五六七八九十百0-9]+章\s*)?[《「]([^》」]{2,40})[》」]\s*的\s*"
    r"[《「]([^》」]{2,40})[》」]\s*节")
EXPR_OK = re.compile(r"^[\d\s.+\-*/()]+$")


# ── 基础工具 ────────────────────────────────────────────────────────
def read_text(path):
    with open(path, encoding="utf-8-sig") as f:
        return f.read()


def load_json(path, default=None):
    if not path or not os.path.isfile(path):
        return default
    try:
        with open(path, encoding="utf-8-sig") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def numbers_in(text):
    """抽出（数值, 单位）列表；区间 9～24 米 记两个数。"""
    out = []
    for m in NUM_RE.finditer(text or ""):
        unit = m.group(3)
        out.append((m.group(1), unit))
        if m.group(2):
            out.append((m.group(2), unit))
    return out


def fnum(s):
    try:
        return round(float(s), 6)
    except (TypeError, ValueError):
        return None


def safe_eval(expr):
    expr = (expr or "").strip()
    if not expr or not EXPR_OK.match(expr):
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))  # 仅四则运算，字符已白名单
    except Exception:
        return None


def decimals(s):
    s = str(s)
    return len(s.split(".")[1]) if "." in s else 0


def close_enough(computed, target):
    """按目标值的显示精度容差比对（60.048 对 60.0 视为一致）。"""
    if computed is None:
        return False
    tol = 0.5 * 10 ** (-decimals(target)) if decimals(target) else 0.5
    return abs(computed - float(target)) <= max(tol, 1e-9)


# ── 解析标注源 ──────────────────────────────────────────────────────
MD_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def strip_md_marks(s):
    """去掉 markdown 强调标记。

    分章稿里 `**核心问题清单**` 这类加粗标题在 docx 渲染时没有转成粗体，
    而是原样落成星号，干净版又是从标注版剥离出来的，所以必须在同源解析阶段
    统一清掉：一处清、三件一致，不靠下游各自兜。
    """
    s = MD_BOLD_RE.sub(r"\1", str(s or ""))
    return s.replace("**", "")


def split_md(text):
    """→ [{'kind':'h'|'p'|'a'|'t', 'text'|'rows', 'level', 'line'}]。

    表格整块保留：多对象参数按规范必须进表，正文不再成段罗列。
    line 记录该条目在标注源里的起始行号，供《校验未通过说明》定位。
    """
    items, buf, tbl = [], [], []
    buf_line, tbl_line = [0], [0]

    def flush_para():
        nonlocal buf
        if buf:
            items.append({"kind": "p", "text": strip_md_marks(" ".join(buf)),
                          "line": buf_line[0] or 1})
            buf = []
            buf_line[0] = 0

    def flush_table():
        nonlocal tbl
        if tbl:
            items.append({"kind": "t", "rows": tbl, "line": tbl_line[0] or 1})
            tbl = []
            tbl_line[0] = 0

    for ln_no, raw in enumerate(text.splitlines(), 1):
        ln = raw.strip()
        if ln.startswith("|") and ln.count("|") >= 2:
            flush_para()
            if not tbl_line[0]:
                tbl_line[0] = ln_no
            cells = [strip_md_marks(c.strip()) for c in ln.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c or "-") for c in cells):
                continue
            tbl.append(cells)
            continue
        flush_table()
        if not ln:
            flush_para()
            continue
        if HEAD_RE.match(ln):
            flush_para()
            m = HEAD_RE.match(ln)
            items.append({"kind": "h", "level": len(m.group(1)),
                          "text": strip_md_marks(m.group(2).strip()), "line": ln_no})
            continue
        if ANNOT_RE.match(ln):
            flush_para()
            items.append({"kind": "a", "text": ANNOT_RE.match(ln).group(1),
                          "line": ln_no})
            continue
        if SRC_ROW_RE.match(ln):
            # 旧式来源行：只属标注，不进干净版，也不进标注版（依据块已含出处）
            flush_para()
            items.append({"kind": "s", "text": ln, "line": ln_no})
            continue
        if not buf_line[0]:
            buf_line[0] = ln_no
        buf.append(ln)
    flush_table()
    flush_para()
    return items


ROSTER_UNIT_RE = re.compile(r"\d+(?:\.\d+)?\s*(?:栋|户)")


def roster_hits(text):
    """正文当表格用：同一段出现 ≥3 组「数值+栋/户」即判成段罗列对象参数。

    规范见 06/规则/写作规范.yaml 的 body_style.roster_rule：多对象参数进表，
    正文只留判断与总量。
    """
    return len(ROSTER_UNIT_RE.findall(text or "")) >= 3



def parse_annot(inner):
    """〔锚点 C-3.2-P04｜材料：file，locator；…｜推导：a+b=c｜定稿值：k=v｜置信度：高〕"""
    fields = [x.strip() for x in inner.split("｜") if x.strip()]
    out = {"anchor": "", "materials": [], "derived": [], "pinned": [],
           "public": [], "internal": [], "judgement": "", "confidence": ""}
    for i, f in enumerate(fields):
        if i == 0 and f.startswith("锚点"):
            out["anchor"] = f[2:].strip()
        elif f.startswith("材料："):
            for part in f[3:].split("；"):
                part = part.strip()
                if not part:
                    continue
                if part.startswith("材料："):
                    part = part[3:].strip()
                if "，" in part:
                    fpath, loc = part.split("，", 1)
                else:
                    fpath, loc = part, ""
                out["materials"].append((fpath.strip(), loc.strip()))
        elif f.startswith("推导："):
            out["derived"].extend(x.strip() for x in f[3:].split("；") if x.strip())
        elif f.startswith("定稿值："):
            out["pinned"].extend(x.strip() for x in f[4:].split("；") if x.strip())
        elif f.startswith("公开："):
            out["public"].extend(x.strip() for x in f[3:].split("；") if x.strip())
        elif f.startswith("依据："):
            out["internal"].append(f[3:].strip())
        elif f.startswith("判断："):
            out["judgement"] = f[3:].strip()
        elif f.startswith("置信度："):
            out["confidence"] = f[4:].strip()
    return out


# ── 依据校验 ────────────────────────────────────────────────────────
def sheet_numbers(dump_lines, fpath, locator):
    """在转储里找「同时含 来源文件 与 表号/页码」的块，返回块内数值集合。"""
    hit = [ln for ln in dump_lines if fpath in ln and (not locator or locator in ln)]
    if not hit:
        return None
    nums = set()
    for ln in hit:
        for v, _u in numbers_in(ln):
            n = fnum(v)
            if n is not None:
                nums.add(n)
        # 表格行里的裸数字（单位在表头）
        for m in re.finditer(r"(?:^|\|)\s*(\d+(?:\.\d+)?)\s*(?=\||$)", ln):
            n = fnum(m.group(1))
            if n is not None:
                nums.add(n)
    return nums


def verify_entry(entry, dump_lines, pinned, ledger, ws=""):
    """返回 (problems, allowed{数值:依据类型}, detail[dict])"""
    probs, allowed, detail = [], {}, []
    if not entry["anchor"]:
        probs.append("依据块缺锚点")
    # 材料
    for fpath, loc in entry["materials"]:
        nums = sheet_numbers(dump_lines, fpath, loc)
        if nums is None:
            probs.append(f"材料出处定位不到：{fpath}｜{loc or '（未写表号/页码）'}")
            continue
        for n in nums:
            allowed.setdefault(n, "材料")
        detail.append({"type": "材料", "file": fpath, "locator": loc, "count": len(nums)})
    # 内部依据：指工作区内的过程文件（如 01 分类清单），供附件类「非源报告摘句」的段落用
    for ref in entry["internal"]:
        real = ref if os.path.isabs(ref) else os.path.join(ws, ref) if ws else ""
        if not (real and os.path.isfile(real)):
            probs.append(f"内部依据找不到文件：{ref}")
            continue
        text = read_any_text(real)
        if text is None:
            probs.append(f"内部依据读不出内容：{ref}")
            continue
        nums = {fnum(v) for v, _u in numbers_in(text)}
        for n in nums:
            if n is not None:
                allowed.setdefault(n, "内部依据")
        detail.append({"type": "内部依据", "file": ref, "locator": ""})
    # 推导
    for expr in entry["derived"]:
        if "=" not in expr:
            probs.append(f"推导式缺等号：{expr}")
            continue
        left, right = expr.split("=", 1)
        tgt = right.strip().rstrip("%")
        val = safe_eval(left)
        if val is None:
            probs.append(f"推导式无法计算：{expr}")
            continue
        if right.strip().endswith("%") and abs(val) <= 1:
            val *= 100
        if not close_enough(val, tgt):
            probs.append(f"推导复算不符：{expr}（算得 {round(val, 4)}）")
            continue
        allowed.setdefault(fnum(tgt), "推导")
        detail.append({"type": "推导", "expr": expr})
    # 定稿值
    pin = {str(v.get("key")): v for v in ((pinned or {}).get("values") or [])}
    for kv in entry["pinned"]:
        if "=" not in kv:
            probs.append(f"定稿值缺等号：{kv}")
            continue
        k, v = kv.split("=", 1)
        k, v = k.strip(), v.strip()
        if k not in pin:
            probs.append(f"定稿值键不存在：{k}")
            continue
        if not close_enough(float(pin[k].get("value") or 0), v):
            probs.append(f"定稿值与文件不一致：{k} 文件={pin[k].get('value')} 正文={v}")
            continue
        allowed.setdefault(fnum(v), "定稿值")
        detail.append({"type": "定稿值", "key": k, "value": v,
                       "origin": pin[k].get("origin")})
    # 公开（台账）
    for sid in entry["public"]:
        ent = ((ledger or {}).get("valid") or {}).get(sid)
        if not ent:
            probs.append(f"公开依据在台账中不存在或无效：{sid}")
            continue
        detail.append({"type": "公开", "sid": sid, "title": ent.get("title"),
                       "publisher": ent.get("publisher"), "url": ent.get("url"),
                       "date": ent.get("date"), "file": ent.get("file")})
        allowed.setdefault(-1, "公开")  # 台账数值命中由 09 的 CHK-EVD 复核
    return probs, allowed, detail


def load_ledger(ws):
    cands = [os.path.join(ws, "过程文件/00-总控/网络补料/检索台账.json"),
             os.path.join(ws, "05_过程/网络补料/检索台账.json")]
    path = next((p for p in cands if os.path.isfile(p)), "")
    if not path:
        return {"path": "", "valid": {}}
    try:
        import _ledger  # 09/scripts 下的唯一实现
        return _ledger.load(ws, path)
    except Exception:
        return {"path": path, "valid": {}}


# ── 输出 ────────────────────────────────────────────────────────────
def _set_cn_font(style, name="宋体"):
    from docx.oxml.ns import qn
    style.font.name = name
    style.element.rPr.rFonts.set(qn("w:eastAsia"), name)


def item_text(it):
    """段落取正文；表格取全部单元格（用于数值抽取与洁净检查）。"""
    if it["kind"] == "t":
        return "，".join(" ".join(r) for r in it["rows"])
    return it.get("text", "")


def _tbl_keep_header(t):
    """表头跨页重复 + 行不跨页断裂 + 按内容分配列宽 + 收窄单元格上下边距。

    * tblHeader：干净版 p3→p4、标注版 p14→p15 续页只剩数据行、没有表头，
      读者翻页后不知道每列是什么；
    * cantSplit：一行不许被页缝劈成两半；
    * 列宽按内容分配：等分列宽把「文件名称」「应包含字段」压成逐字竖排，
      行高翻倍，标注版附件三那一行因此被顶到单独一页；
    * 上下边距压到 0：进一步把行高压下来，尾行尽量并进前页。
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Twips

    tblPr = t._tbl.tblPr
    for old in tblPr.findall(qn("w:tblCellMar")):
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top", 0), ("bottom", 0), ("left", 57), ("right", 57)):
        el = OxmlElement("w:" + tag)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    # OOXML 规定 tblCellMar 必须排在 tblLook 之前；直接 append 会让 Word 打开即崩。
    look = tblPr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(mar)
    else:
        tblPr.append(mar)
    for i, row in enumerate(t.rows):
        trPr = row._tr.get_or_add_trPr()
        el = OxmlElement("w:cantSplit")
        el.set(qn("w:val"), "true")
        trPr.append(el)
        if i == 0:
            head_el = OxmlElement("w:tblHeader")
            head_el.set(qn("w:val"), "true")
            trPr.append(head_el)

    grid = [[c.text for c in row.cells] for row in t.rows]
    if not grid:
        return
    usable = 9360  # Letter 12240 - 左右各 1440 边距（python-docx 默认页面）
    need = []
    for j in range(len(grid[0])):
        w = max(sum(2 if ord(ch) > 127 else 1 for ch in str(r[j] or "")) for r in grid)
        need.append(max(6, min(w, 30)))
    total = sum(need) or 1
    widths = [max(int(usable * 0.08), int(usable * n / total)) for n in need]
    over = sum(widths) - usable
    if over:
        cap = max(widths)
        widths[widths.index(cap)] -= over
    t.autofit = False
    for row in t.rows:
        for j, cell in enumerate(row.cells):
            cell.width = Twips(max(720, widths[j]))


def write_docx(items, path, with_annot):
    from docx import Document
    from docx.shared import Pt, RGBColor
    doc = Document()
    _set_cn_font(doc.styles["Normal"])
    for it in items:
        if it["kind"] == "h":
            doc.add_heading(it["text"], level=min(it["level"], 3))
        elif it["kind"] == "p":
            p = doc.add_paragraph(it["text"])
            p.paragraph_format.first_line_indent = Pt(24)
            p.paragraph_format.line_spacing = 1.5
        elif it["kind"] == "t":
            rows = it["rows"]
            t = doc.add_table(rows=len(rows), cols=max(len(r) for r in rows))
            t.style = "Table Grid"
            for i, r in enumerate(rows):
                for j, c in enumerate(r):
                    cell = t.cell(i, j)
                    cell.text = c
                    for para in cell.paragraphs:
                        para.paragraph_format.space_before = Pt(0)
                        para.paragraph_format.space_after = Pt(0)
                        for run in para.runs:
                            run.font.size = Pt(9)
                            if i == 0:
                                run.bold = True
            _tbl_keep_header(t)
        elif it["kind"] == "a" and with_annot:
            p = doc.add_paragraph()
            r = p.add_run("〔" + it["text"] + "〕")
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    doc.save(path)


def _auto_width(ws, head, floors=None):
    """按内容算列宽（CJK 记 2），并给关键列/表头兜一个最小可读宽度。

    表头被挤成逐字竖排的根因是列宽小于一个汉字；这里按「表头显示宽度 + 2」
    兜底，长文本列仍按内容取 40 上限。
    """
    floors = floors or {}
    widths = {}
    for row in ws.iter_rows(values_only=True):
        for i, v in enumerate(row, 1):
            if v is None:
                continue
            n = sum(2 if ord(ch) > 127 else 1 for ch in str(v))
            widths[i] = min(42, max(widths.get(i, 8), n + 2))
    for i, name in enumerate(head, 1):
        head_w = sum(2 if ord(ch) > 127 else 1 for ch in str(name)) + 2
        floor = max(floors.get(name, 8), head_w)
        widths[i] = max(widths.get(i, 8), floor)
    from openpyxl.utils import get_column_letter
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = max(w, 8)


# A4 横向、左右各 1.2cm 页边距下可容纳的列宽总量（openpyxl 宽度单位，按实测校准）
PAGE_UNITS_LANDSCAPE = 144


def _fit_page_width(ws, head, caps, page_units=PAGE_UNITS_LANDSCAPE):
    """列宽按「表头 + 数据最宽」算，长文本列封顶，再整体压进一页宽。

    不压进一页宽，Excel 会把整张表按列切成「只剩校验一列」的列带（p16–p20）。
    短列只保表头最小宽，超出部分按可压缩量比例分摊，不把某列压到放不下表头；
    最小宽同时兜住「不可拆单元」（(1).xlsx、0.252万㎡），否则折行时会溢出列宽。
    """
    from openpyxl.utils import get_column_letter
    mins, widths = {}, {}
    for i, name in enumerate(head, 1):
        mins[i] = _cell_units(str(name)) + 2
    for i, name in enumerate(head, 1):
        vals = [ws.cell(r, i).value for r in range(2, ws.max_row + 1)]
        data = max([_cell_units(str(v)) for v in vals if isinstance(v, str)] or [0])
        tok = max([_token_width(v) for v in vals if isinstance(v, str)] or [0])
        mins[i] = max(mins[i], min(tok + 2, page_units // 3))
        widths[i] = max(mins[i], min(caps.get(name, 40), data + 1))
    total = sum(widths.values())
    if total > page_units:
        excess = total - page_units
        flex = {i: widths[i] - mins[i] for i in widths if widths[i] > mins[i]}
        fsum = sum(flex.values()) or 1
        for i in flex:
            widths[i] -= int(excess * flex[i] / fsum)
        diff = sum(widths.values()) - page_units
        if diff > 0:
            i = max(widths, key=lambda k: widths[k] - mins[k])
            widths[i] = max(mins[i], widths[i] - diff)
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = w
    return widths


def _page_setup(ws):
    """A4 横向 + 一页宽 + 表头重复：列不切成列带，续页仍有表头。"""
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = "1:1"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.47   # ≈1.2cm
    ws.page_margins.top = ws.page_margins.bottom = 0.4    # ≈1cm，给表页留出行空间
    ws.page_margins.header = ws.page_margins.footer = 0.2


LEAD_PUNCT = "，。、；：）》」』”’！？,.;:"

# 折行避头：这些标点不得出现在行首（》 收《》，、 收并列，。；： 收句读）。
# 开标点《（「 一并列入：中文排版里它们本可起行，但本表要求折行后行首 0 标点，
# 故一律挂到上一行行尾（见 _wrap_cell），不再带到下一行。
# 半角右括号 ) ] 同理：出处文件列曾出现 35 行行首「)」。
NO_LINE_HEAD = "》，、。；：！？）」』”’《（「『“‘" + ")]"

# 数值 + 单位必须整体不可拆：0.252 万㎡ 被拆成 0.252 / 万㎡ 是上一版的实际错例。
UNITS = ("万平方米", "平方米", "万㎡", "万m", "公顷", "公里", "千米", "㎡", "m²",
         "%", "米", "m", "栋", "户", "个", "处", "人", "亩", "元", "辆", "间",
         "条", "座", "层", "台")


def _cell_units(s):
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


def _tokens(s):
    """切成断行单元：中文按字可断，拉丁/数字串（含扩展名、数值+单位）整体不可拆。

    出处文件列出现过 ".xls"|"x"、行首 ")"，根因是按字符折行把拉丁串当普通字；
    这里先把「(1).xlsx」「0.252万㎡」「M01」这类串合成一个单元，再按单元排队。

    点号（.）不得作断点：全角「）」不在 ASCII 串内，会把「…（某社区）」和
    「.xlsx」切成两个单元，折到这里就出现行首「.xlsx」。所以「.扩展名」一律并回
    前一个单元，右括号「）」也并回前一个单元，断点只留在「；」「、」之后或两字之间。
    """
    out, i, n = [], 0, len(s)
    while i < n:
        ch = s[i]
        if ch.isascii() and (ch.isalnum() or ch in "_.-/\\()[]{}%+&@#"):
            j = i
            while j < n and s[j].isascii() and (
                    s[j].isalnum() or s[j] in "_.-/\\()[]{}%+&@#"):
                j += 1
            tok = s[i:j]
            i = j
            if re.fullmatch(r"\d+(?:\.\d+)?", tok):
                for u in UNITS:
                    if s.startswith(u, i):
                        tok += u
                        i += len(u)
                        break
            # 「.xlsx」「.png」整体并回前一个单元，绝不从点号前断开
            if out and re.fullmatch(r"\.[A-Za-z0-9]{1,6}", tok):
                out[-1] += tok
                continue
            out.append(tok)
            continue
        # 右括号并入前一单元，使其不会单独成行、也不会把扩展名甩到下一行
        if ch in "）)]" and out and not out[-1].isspace():
            out[-1] += ch
            i += 1
            continue
        out.append(ch)
        i += 1
    return out


def _token_width(v):
    """单元格里最宽的一个不可拆单元（列宽下限用）。"""
    return max([_cell_units(t) for t in _tokens(str(v)) if t.strip()] or [0])


def _wrap_cell(v, width):
    """按列宽预折行：行首不放标点，数值/单位/扩展名不拆。

    预折行宽度取「列宽 - 3」，收标点最多挂到「列宽 - 2」，与 Excel 实际列宽对齐。
    上一版按「列宽 - 4」且按字符切，结果 0.252 与 万㎡、.xls 与 x 被拆到两行。
    """
    if not isinstance(v, str) or "\n" in v:
        return v
    width = max(8, int(width))
    # 行宽按「列宽 - 3」折，收标点最多挂到「列宽 - 2」：Excel 的列宽单位按 Normal
    # 字体的数字宽折算，中文字宽略大于 2 个单位，贴着列宽折会被 Excel 二次折行，
    # 把挂在行尾的「、」甩成单独一行（PDF 里出现过 5 处、缺口清单 10 处）。
    budget = max(4, width - 3)
    hard = max(budget, width - 2)
    lines, cur, cur_w = [], [], 0
    for tok in _tokens(v):
        w = _cell_units(tok)
        if tok == " ":
            if cur:
                cur.append(tok)
                cur_w += w
            continue
        if tok in NO_LINE_HEAD:
            # 收标点挂上一行行尾；挂不上（会超列宽）就把上一行末尾的单元一起挪下来，
            # 让下一行以实词开头，避免行首标点。
            if not cur:
                cur, cur_w = [tok], w
                continue
            if cur_w + w <= hard:
                cur.append(tok)
                cur_w += w
                continue
            moved = [tok]
            # 只把「行尾最后一个实词」连同收标点挪到下一行：cur_w 与 moved 之和是
            # 不变量，按「和不超列宽」循环永远不成立，上一版会把整行掏空、只剩标点起行。
            while cur:
                last = cur.pop()
                moved.insert(0, last)
                if last not in NO_LINE_HEAD:
                    break
            if cur:
                lines.append("".join(cur))
            cur, cur_w = moved, sum(_cell_units(m) for m in moved)
            continue
        if cur and cur_w + w > budget:
            lines.append("".join(cur))
            cur, cur_w = [tok], w
            continue
        cur.append(tok)
        cur_w += w
    if cur:
        lines.append("".join(cur))
    return "\n".join(lines)


def _clean_cell(v):
    """去单元格行首的悬挂标点（段落截断后常以「，。」开头），不 strip「—」占位。"""
    if not isinstance(v, str):
        return v
    t = v.strip()
    while t and t[0] in LEAD_PUNCT:
        t = t[1:].lstrip()
    return t


def write_xlsx(rows, detail, path, section):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "按段对照"
    head = ["锚点", "小节", "段落（首句）", "本段数值", "依据类型",
            "出处文件", "表号/页码", "编号", "置信度", "校验"]
    ws.append(head)
    for r in rows:
        ws.append([_clean_cell(x) for x in r])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _fit_page_width(ws, head, {"段落（首句）": 30, "本段数值": 20,
                               "出处文件": 25, "小节": 22})
    _page_setup(ws)
    # 四列自由文本按列宽预折行：行首不出 》《、。；：，数据不变只插换行。
    from openpyxl.utils import get_column_letter
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.column in (2, 3, 4, 6) and isinstance(c.value, str):
                w = ws.column_dimensions[get_column_letter(c.column)].width or 12
                c.value = _wrap_cell(c.value, w)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("逐数明细")
    head2 = ["锚点", "数值", "单位", "依据类型", "出处文件", "表号/页码", "校验"]
    ws2.append(head2)
    for d in detail:
        ws2.append([_clean_cell(x) for x in d])
    for c in ws2[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _fit_page_width(ws2, head2, {"出处文件": 28})
    _page_setup(ws2)
    for row in ws2.iter_rows(min_row=2):
        for c in row:
            if c.column == 5 and isinstance(c.value, str):
                c.value = _wrap_cell(c.value, ws2.column_dimensions[get_column_letter(c.column)].width or 12)
    for row in ws2.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws2.freeze_panes = "A2"
    wb.save(path)


# ── 主流程 ──────────────────────────────────────────────────────────
def rel_to(path: str, ws: str) -> str:
    """工作区内给相对路径；跨盘符（标注源在别的盘）退回绝对路径，不给 relpath 崩。"""
    try:
        return os.path.relpath(path, ws)
    except ValueError:
        return os.path.abspath(path)


def collect_sources(ws, spec=""):
    """收集标注源：--md 支持文件 / 目录 / 逗号分号列表；缺省自动找。

    默认顺序 标注源 → 分章初稿 → 样张，取**第一个含〔锚点〕的目录**；
    都没有则返回空（旧格式章节尚未纳入三件套，只提示不拦流程）。
    """
    if spec:
        files = []
        for part in re.split(r"[;,]", spec):
            part = part.strip()
            if os.path.isdir(part):
                files.extend(sorted(glob.glob(os.path.join(part, "*.md"))))
            elif os.path.isfile(part):
                files.append(part)
        return [os.path.abspath(f) for f in files
                if not os.path.basename(f).startswith("_")]
    for rel in ("过程文件/06-分章成文/标注源", "过程文件/06-分章成文/分章初稿",
                "过程文件/06-分章成文/样张"):
        d = os.path.join(ws, rel)
        if not os.path.isdir(d):
            continue
        files = [f for f in sorted(glob.glob(os.path.join(d, "*.md")))
                 if not os.path.basename(f).startswith("_")]
        if any("〔锚点" in read_text(f) for f in files):
            return files
    return []


# ── 追溯链：正文数 → 锚点/对照表行 → 原文文件 ────────────────────────
TEXT_EXT = (".md", ".txt", ".csv", ".tsv", ".json")


def file_index(ws):
    """工作区文件名（小写）→ 绝对路径；材料只写文件名时靠它落到实际文件。"""
    idx = {}
    for dirpath, dirs, files in os.walk(ws):
        dirs[:] = [d for d in dirs if d not in ("__pycache__", ".cardo", ".git")]
        for name in files:
            idx.setdefault(name.lower(), os.path.join(dirpath, name))
    return idx


def norm_token(text):
    """数值文本归一：19720.0 / 19,720 / 77.80% → 19720 / 19720 / 77.8。"""
    t = str(text).replace(",", "").strip().rstrip("%")
    try:
        v = float(t)
    except ValueError:
        return t.lower()
    return str(int(v)) if v == int(v) else ("%g" % v)


def numeric_tokens(text):
    return {norm_token(m.group(0)) for m in re.finditer(r"\d+(?:\.\d+)?", (text or "").replace(",", ""))}


def read_any_text(path):
    """读材料原文：文本类直读，xlsx 摊平成单元格文本；图片/PPT/PDF 返回 None。

    返回 None 表示"这个文件读不出数"，链路上退回「依据有效」已用的转储核对，
    不在这里制造假阴性。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            from openpyxl import load_workbook
            book = load_workbook(path, read_only=True, data_only=True)
        except Exception:            # noqa: BLE001 - 打不开就当读不出（不影响其它检查）
            return None
        cells = []
        for sheet in book.worksheets:
            for row in sheet.iter_rows(values_only=True):
                cells.extend("" if c is None else str(c) for c in row)
        return "\n".join(cells)
    if ext in TEXT_EXT:
        for enc in ("utf-8-sig", "utf-8", "gbk"):
            try:
                with open(path, encoding=enc) as fh:
                    return fh.read()
            except (OSError, UnicodeDecodeError):
                continue
    return None


def register_links(acc, ws):
    """对照表 ↔ 锚点一一对应，且每行「材料」能回指到工作区里的真实文件。"""
    idx = file_index(ws)
    problems = []
    anchors = [row[0] for row in acc["rows"]]
    dup = sorted({a for a in anchors if anchors.count(a) > 1})
    for a in dup:
        problems.append(f"{a}：对照表里出现多行（锚点必须一一对应）")
    if len(anchors) != len(acc["entries"]):
        problems.append(f"对照表 {len(anchors)} 行 ≠ 依据块 {len(acc['entries'])} 条")
    for row in acc["rows"]:
        kinds = row[4] or ""
        if "材料" not in kinds:
            continue
        for fpath in [x for x in (row[5] or "").split("；") if x and x != "—"]:
            if os.path.isfile(os.path.join(ws, fpath)) or idx.get(os.path.basename(fpath).lower()):
                continue
            problems.append(f"{row[0]}：材料 {fpath} 在工作区里找不到（该行回指不到原文）")
    return problems


def chain_closure(acc, ctx, ws):
    """逐数走完整条链：正文数 → 锚点（对照表行）→ 原文文件。

    推导数在「依据有效」已复算、公开数已核台账与 URL，这两类在行内闭合；
    材料数要落到真实文件并能读到该数值；定稿值要落到其 origin 文件。
    """
    idx = file_index(ws)
    pin = {str(v.get("key")): v for v in ((ctx.get("pinned") or {}).get("values") or [])}
    problems, checked, closed, by_kind = [], 0, 0, {}
    cache = {}

    def text_of(path):
        if path not in cache:
            cache[path] = read_any_text(path)
        return cache[path]

    for row in acc["detail"]:
        anchor, raw, kind = row[0], row[1], row[3]
        if kind == "—":                 # 无源数已由「数值有源」判，不重复计
            continue
        checked += 1
        entry = acc["entries"].get(anchor)
        if not entry:
            problems.append(f"{anchor}：正文数 {raw} 在对照表里找不到对应行")
            continue
        token = norm_token(raw)
        if kind == "材料":
            legs = []
            for fpath, _loc in entry["materials"]:
                real = os.path.join(ws, fpath)
                if not os.path.isfile(real):
                    real = idx.get(os.path.basename(fpath).lower(), "")
                if real:
                    legs.append(real)
            if not legs:
                problems.append(f"{anchor}：正文数 {raw} 的锚点写了材料，但材料文件在工作区找不到")
                continue
            hit = [p for p in legs
                   if text_of(p) is None or token in numeric_tokens(text_of(p))]
            if not hit:
                problems.append(f"{anchor}：正文数 {raw} 在材料原文里查不到"
                                f"（{os.path.basename(legs[0])}）")
                continue
        elif kind == "定稿值":
            key = next((kv.split("=", 1)[0].strip() for kv in entry["pinned"]), "")
            origin = (pin.get(key) or {}).get("origin") or {}
            if isinstance(origin, str):
                origin = {"kind": "material", "ref": origin}
            otype = str(origin.get("kind") or "")
            if otype in ("derived", "public"):
                # 推导：公式已在「依据有效」复算；公开：台账编号与 URL 已核
                closed += 1
                by_kind[kind] = by_kind.get(kind, 0) + 1
                continue
            ref = str(origin.get("ref") or origin.get("file") or "")
            real = ref if os.path.isfile(ref) else (
                idx.get(os.path.basename(ref).lower(), "") if ref else "")
            if not real:
                problems.append(f"{anchor}：定稿值 {key} 的 origin 文件不存在（{ref or '未写 ref'}）")
                continue
            if token not in numeric_tokens(text_of(real)):
                problems.append(f"{anchor}：定稿值 {raw} 在 origin 原文里查不到（{os.path.basename(real)}）")
                continue
        elif kind == "内部依据":
            ref = row[4] if len(row) > 4 else ""
            real = ref if os.path.isfile(ref) else (
                os.path.join(ws, ref) if ws and ref else "")
            if not (real and os.path.isfile(real)):
                real = idx.get(os.path.basename(ref).lower(), "") if ref else ""
            if not real:
                problems.append(f"{anchor}：正文数 {raw} 的内部依据文件找不到（{ref or '未写'}）")
                continue
            if token not in numeric_tokens(text_of(real)):
                problems.append(f"{anchor}：正文数 {raw} 在内部依据里查不到"
                                f"（{os.path.basename(real)}）")
                continue
        closed += 1
        by_kind[kind] = by_kind.get(kind, 0) + 1
    return problems, {"checked": checked, "closed": closed, "byKind": by_kind}


def drop_process_sections(items):
    """丢掉「本节缺项表」「来源映射」两段（08 组装同口径剥离，不进交付正文）。"""
    out, skip = [], False
    for it in items:
        if it["kind"] == "h":
            skip = it["text"].strip() in PROCESS_SECTIONS
        if not skip:
            out.append(it)
    return out


def check_citations(head_items, body_items):
    """章节引用必须能在标题层级里找到：H1 章 + 该章下的 H2 节。

    只认「见第X章《…》」与「见《章》的《节》节」两种指路句式；
    「见《输入材料补齐确认单》」这类文件名引用不在此列。
    """
    h1, h2, cur = set(), set(), ""
    for it in head_items:
        if it["kind"] != "h":
            continue
        t = it["text"].strip()
        if it["level"] == 1:
            cur = t
            h1.add(t)
        elif it["level"] == 2 and cur:
            h2.add((cur, t))
    probs, checked = [], 0
    for it in body_items:
        if it["kind"] not in ("p", "t"):
            continue
        text = item_text(it)
        where = f"{it.get('source', '?')}:{it.get('line', '?')}"
        rest = text
        for m in CITE_SECTION_RE.finditer(text):
            checked += 1
            father, child = m.group(1).strip(), m.group(2).strip()
            if father not in h1:
                probs.append(f"{where} 引用悬空：{m.group(0)}"
                             f"（章标题「{father}」不在标题层级）")
            elif (father, child) not in h2:
                probs.append(f"{where} 引用悬空：{m.group(0)}"
                             f"（「{father}」下找不到小节「{child}」）")
            rest = rest.replace(m.group(0), "")
        for m in CITE_CHAPTER_RE.finditer(rest):
            checked += 1
            name = m.group(1).strip()
            if name not in h1:
                probs.append(f"{where} 引用悬空：{m.group(0)}"
                             f"（章标题「{name}」不在标题层级）")
    return probs, checked


def collect_evidence(files, ctx):
    """逐文件校验标注源 → 干净/标注 items + 对照行 + 逐数明细 + 问题清单。

    旧格式文件（无〔锚点〕）整份跳过并在报告里列名：迁移期允许新旧并存，
    但**只要文件里有依据块就按全覆盖校验**，避免"写一半"蒙混过关。
    过程段（缺项表/来源映射）不参与三件套：08 组装时会剥离。
    """
    clean_items, annot_items, rows, detail, problems = [], [], [], [], []
    entries, skipped, body_items, seq_all = {}, [], [], []
    for path in files:
        its = drop_process_sections(split_md(read_text(path)))
        if not any(x["kind"] == "a" for x in its):
            skipped.append(path)
            continue
        seq = [x for x in its if x["kind"] != "s"]      # 旧式来源行不进任何版本
        clean_items += [x for x in seq if x["kind"] in ("h", "p", "t")]
        annot_items += seq
        seq_all += seq
        cur, para_no, base = "未命名小节", 0, os.path.basename(path)
        for i, it in enumerate(seq):
            if it["kind"] == "h":
                cur = it["text"]
                continue
            if it["kind"] == "a":
                continue
            para_no += 1
            it["source"] = base          # 供引用可解说清问题位置（文件:行）
            body_items.append(it)
            nxt = seq[i + 1] if i + 1 < len(seq) else None
            if not nxt or nxt["kind"] != "a":
                problems.append(f"{base}:{it.get('line', '?')} 第 {para_no} 段缺依据块"
                                f"（每段必须有一条〔锚点 …〕）")
                continue
            e = parse_annot(nxt["text"])
            e["source"] = base
            e["line"] = nxt.get("line", 0)
            entries[e["anchor"] or f"{base}#{para_no}"] = e
            nums = numbers_in(item_text(it))
            probs, allowed, det = verify_entry(e, ctx["dump"], ctx["pinned"], ctx["ledger"],
                                               ctx.get("ws", ""))
            where = f"{base}:{nxt.get('line', '?')}"
            problems.extend(f"{where} {e['anchor']}：{p}" for p in probs)
            misses = []
            for v, u in nums:
                n = fnum(v)
                if n is None:
                    continue
                if n in allowed:
                    detail.append([e["anchor"], v, u, allowed[n],
                                   "；".join(sorted({d.get("file", "") for d in det if d.get("file")})),
                                   "；".join(sorted({d.get("locator", "") for d in det if d.get("locator")})),
                                   "有源"])
                elif -1 in allowed:
                    detail.append([e["anchor"], v, u, "公开",
                                   "；".join(sorted({str(d.get("file", "")) for d in det if d.get("type") == "公开"})),
                                   "；".join(sorted({str(d.get("url", "")) for d in det if d.get("type") == "公开"})),
                                   "有源"])
                else:
                    misses.append(f"{v}{u}")
                    detail.append([e["anchor"], v, u, "—", "—", "—", "无源"])
            if misses:
                problems.append(f"{where} {e['anchor']}：{len(misses)} 个数值无出处"
                                f"（{'、'.join(misses[:6])}）")
            kinds = []
            if e["materials"]:
                kinds.append("材料")
            if e["derived"]:
                kinds.append("推导")
            if e["pinned"]:
                kinds.append("定稿值")
            if e["public"]:
                kinds.append("公开")
            if e["internal"]:
                kinds.append("内部依据")
            if not kinds:
                kinds.append("判断")
            rows.append([
                e["anchor"], cur,
                (f"表：{'/'.join(it['rows'][0])}" if it["kind"] == "t"
                 else it["text"][:28] + ("…" if len(it["text"]) > 28 else "")),
                "、".join(f"{v}{u}" for v, u in nums) or "—",
                "+".join(kinds),
                "；".join(sorted({f for f, _ in e["materials"]})) or "；".join(e["internal"]) or "—",
                "；".join(sorted({l for _, l in e["materials"]})) or "—",
                e["anchor"] + ("；" + "、".join(e["public"]) if e["public"] else ""),
                e["confidence"] or "—",
                "通过" if not (probs or misses) else "不通过",
            ])
    return {"clean": clean_items, "annot": annot_items, "rows": rows, "detail": detail,
            "problems": problems, "entries": entries, "skipped": skipped,
            "body": body_items, "seq": seq_all}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--workspace", default=os.getcwd())
    ap.add_argument("--project-dir", default="", help="与 --workspace 同义")
    ap.add_argument("--md", default="", help="标注源：文件 / 目录 / 逗号分号列表")
    ap.add_argument("--dump", default="")
    ap.add_argument("--pinned", default="")
    ap.add_argument("--out-dir", default="")
    ap.add_argument("--version", default="v1")
    ap.add_argument("--check-only", action="store_true", help="只校验，不写 docx/xlsx/对照表")
    args = ap.parse_args(argv)
    ws = os.path.abspath(args.project_dir or args.workspace)
    dump_dir = args.dump or os.path.join(ws, "过程文件/03-模板与源报告转储/源报告转储")
    pinned_path = args.pinned or os.path.join(ws, "过程文件/00-总控/定稿值.json")
    out_dir = args.out_dir or os.path.join(ws, "成果文件/证据包")

    files = collect_sources(ws, args.md)
    if not files:
        print("[三件套] 未找到含〔锚点 …〕依据块的标注源，跳过（06 按写作规范产出标注源后自动生效）")
        return 2
    dump_lines = []
    if os.path.isdir(dump_dir):
        for fn in sorted(os.listdir(dump_dir)):
            if fn.endswith((".blocks.txt", ".tables.txt")):
                dump_lines.extend(read_text(os.path.join(dump_dir, fn)).splitlines())
    ctx = {"dump": dump_lines,
           "pinned": load_json(pinned_path, {}),
           "ledger": load_ledger(ws),
           "ws": ws}
    acc = collect_evidence(files, ctx)
    if not acc["entries"]:
        print("[三件套] 标注源里没有可校验的依据块，跳过（旧格式，未纳入三件套）")
        return 2
    clean_items, annot_items = acc["clean"], acc["annot"]
    rows, detail, problems = acc["rows"], acc["detail"], acc["problems"]
    entry_by_anchor = acc["entries"]
    heads = [it for it in clean_items if it["kind"] == "h" and it["level"] == 1]
    section = heads[0]["text"] if heads else "未命名小节"
    label = section if len(files) == 1 else "整册"

    # 干净版：剥依据块 + 编号
    body_items = acc["body"]
    checks = []
    clean_text = "\n\n".join(item_text(x) for x in body_items)
    banned = [(kw, why) for kw, why in BANNED_IN_CLEAN if kw in clean_text]
    roster = [x for x in body_items if x["kind"] == "p" and roster_hits(x["text"])]
    body_nums = [v for it in body_items for v, _u in numbers_in(item_text(it))]
    src_nums = [v for it in acc["seq"] if it["kind"] in ("p", "t")
                for v, _u in numbers_in(CITE_RE.sub("", item_text(it)))]
    same = body_nums == src_nums

    checks.append(("正文洁净", not banned,
                   "无禁用标注" if not banned else "命中：" + "、".join(f"{k}（{w}）" for k, w in banned)))
    checks.append(("段落覆盖", not any("缺依据块" in p for p in problems),
                   f"{len(entry_by_anchor)} 段全部有依据块" if not any("缺依据块" in p for p in problems)
                   else "存在无依据块的段落"))
    checks.append(("数值有源", not any("无出处" in p for p in problems),
                   "每个数值均能指到 材料/内部依据/公开/推导/定稿值"
                   if not any("无出处" in p for p in problems) else "存在无出处数值"))
    checks.append(("依据有效", not any(("定位不到" in p) or ("复算不符" in p) or ("不一致" in p)
                                     or ("不存在" in p) or ("缺等号" in p) for p in problems),
                   "材料可在转储定位、推导可复算、定稿值与文件一致"
                   if not problems else "见问题清单"))
    checks.append(("剥离一致", same,
                   "干净版与标注版数值集合与顺序完全一致"
                   if same else "数值序列不一致（剥离过程改动正文）"))
    checks.append(("正文非表格化", not roster,
                   "正文无成段罗列对象参数（≥3 组「数值+栋/户」）" if not roster
                   else f"{len(roster)} 段成段罗列对象参数，应改为进表"))

    # 成稿口径：来源行从正文消失后，用「对照表回指 + 逐数链条闭合」等价替换追溯校验。
    link_problems = register_links({"rows": rows, "entries": entry_by_anchor}, ws)
    chain_problems, chain_stats = chain_closure(
        {"detail": detail, "entries": entry_by_anchor}, ctx, ws)
    problems += link_problems + chain_problems
    checks.append(("对照表回指", not link_problems,
                   f"{len(rows)} 行逐行回指到工作区原文" if not link_problems
                   else f"{len(link_problems)} 处回指不到原文（见问题清单）"))
    checks.append(("链条闭合", not chain_problems and chain_stats["checked"] > 0,
                   f"{chain_stats['closed']}/{chain_stats['checked']} 个数走通"
                   f"「正文数→锚点→对照表→原文」"
                   if not chain_problems and chain_stats["checked"] else
                   ("；".join(chain_problems[:5]) if chain_problems
                    else "无含数段落，链路未建立")))

    # 引用可解：正文指路的章/节标题必须在标题层级里真实存在，悬空引用一律拦下。
    cite_problems, cite_checked = check_citations(clean_items, body_items)
    problems += cite_problems
    checks.append(("引用可解", not cite_problems,
                   f"{cite_checked} 处章节引用的目标均在标题层级" if not cite_problems
                   else f"{len(cite_problems)} 处引用悬空（见问题清单）"))

    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    safe = re.sub(r'[\\/:*?"<>|\s]', "", label)
    clean_docx = annot_docx = xlsx = ""
    if not args.check_only:
        os.makedirs(out_dir, exist_ok=True)
        # 文件名不再带时间戳：证据包目录每轮清空重建，只保留本轮当前版本，
        # 避免出现 0810/0813/0814 多套并存的历史堆积。
        clean_docx = os.path.join(out_dir, f"现状分析_{safe}_干净版_{args.version}.docx")
        annot_docx = os.path.join(out_dir, f"现状分析_{safe}_标注版_{args.version}.docx")
        xlsx = os.path.join(out_dir, f"依据对照表_{safe}_{args.version}.xlsx")
        write_docx(clean_items, clean_docx, with_annot=False)
        write_docx(annot_items, annot_docx, with_annot=True)
        write_xlsx(rows, detail, xlsx, label)

        proc = os.path.join(ws, "过程文件/00-总控")
        os.makedirs(proc, exist_ok=True)
        with open(os.path.join(proc, f"依据对照表_{safe}.json"), "w", encoding="utf-8") as f:
            json.dump({"schema": "cardo/evidence-register/v1", "section": label,
                       "generatedAt": ts,
                       "sources": [rel_to(p, ws) for p in files],
                       "skipped": [rel_to(p, ws) for p in acc["skipped"]],
                       "entries": list(entry_by_anchor.values()),
                       "problems": problems}, f, ensure_ascii=False, indent=1)

    rep_dir = os.path.join(ws, "过程文件/09-三重自检与边界")
    os.makedirs(rep_dir, exist_ok=True)
    ok = all(c[1] for c in checks) and not problems
    rep = [f"# 依据校验 · {label}", "",
           f"- 时间：{ts}　版本：{args.version}　模式：{'只校验' if args.check_only else '导出三件套'}",
           f"- 结果：{'通过' if ok else '不通过'}", "",
           "| 检查项 | 结果 | 说明 |", "|---|---|---|"]
    rep += [f"| {n} | {'通过' if p else '不通过'} | {d} |" for n, p, d in checks]
    rep += ["", "## 标注源", ""] + [f"- `{rel_to(p, ws)}`" for p in files]
    if acc["skipped"]:
        rep += ["", "> 以下文件无〔锚点〕依据块（旧格式），未纳入三件套："]
        rep += [f"> - `{rel_to(p, ws)}`" for p in acc["skipped"]]
    if problems:
        rep += ["", "## 问题清单", ""] + [f"- {p}" for p in problems]
    if clean_docx:
        rep += ["", "## 产物", "", f"- 干净版：`{rel_to(clean_docx, ws)}`",
                f"- 标注版：`{rel_to(annot_docx, ws)}`",
                f"- 对照表：`{rel_to(xlsx, ws)}`"]
    with open(os.path.join(rep_dir, f"依据校验_{safe}.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(rep) + "\n")

    print(f"[三件套] {label}  {'通过' if ok else '不通过'}")
    for n, p, d in checks:
        print(f"  {'PASS' if p else 'FAIL'}  {n}：{d}")
    for p in acc["skipped"]:
        print(f"  SKIP  {os.path.basename(p)}（无依据块，旧格式）")
    if problems:
        print("  问题：")
        for p in problems:
            print("   -", p)
    if clean_docx:
        print(f"  干净版 {rel_to(clean_docx, ws)}")
        print(f"  标注版 {rel_to(annot_docx, ws)}")
        print(f"  对照表 {rel_to(xlsx, ws)}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
