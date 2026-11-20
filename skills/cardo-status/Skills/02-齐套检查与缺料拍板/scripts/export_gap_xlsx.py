#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""导出《数据缺口清单.xlsx》：材料缺口 / 章节缺项 / 事实缺失 三张表。"""
import argparse
import json
import os
from pathlib import Path
import re
import sys

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:  # pragma: no cover
    print("缺少 openpyxl：pip install openpyxl")
    raise SystemExit(2)

GAP_RE = re.compile(r"【(缺项-[A-F]\d{1,3})：(.+?)｜建议来源：(.+?)】")
# 缺项表在成稿里是**表格行**（| 缺项-A01 | 缺什么 | 找谁补 | 现有替代性判断 |），
# 与 08 步 assemble_chapters 的 ROW_RE 口径一致。
# 旧写法只认行内格式「【缺项-A01：…｜建议来源：…】」，实测采集 0 条——
# 「章节缺项」表只剩表头，缺口清单漏掉全部章节缺项。
GAP_ROW_RE = re.compile(r"^\|\s*缺项-([A-Za-z]?)(\d{1,3})\s*\|(.*)\|\s*$")


def style_header(ws, ncol):
    fill = PatternFill("solid", fgColor="D9E2F3")
    for c in range(1, ncol + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


# ── 版式：一页宽、表头跨页重复、按不可拆单元折行 ─────────────────────
# 三张表原来不设打印参数，列宽合计 80–210 个单位，打印时被 Excel 按列竖切成
# 「只剩说明一列」「整页只有落位目录」的列带。这里统一按 A4 横向一页宽排版。
PAGE_UNITS_LANDSCAPE = 144
NO_LINE_HEAD = "》，、。；：！？）」』”’《（「『“‘" + ")]"
UNITS = ("万平方米", "平方米", "万㎡", "公顷", "公里", "㎡", "%", "米", "m",
         "栋", "户", "个", "处", "人", "亩", "元")


def _cell_units(s):
    return sum(2 if ord(ch) > 127 else 1 for ch in s)


def _tokens(s):
    """断行单元：中文按字可断，拉丁/数字串（含扩展名、数值+单位）整体不可拆。"""
    out, i, n = [], 0, len(s)
    while i < n:
        ch = s[i]
        if ch.isascii() and (ch.isalnum() or ch in "_.-/\\()[]{}%+&@#"):
            j = i
            while j < n and s[j].isascii() and (
                    s[j].isalnum() or s[j] in "_.-/\\()[]{}%+&@#"):
                j += 1
            out.append(s[i:j])
            i = j
            if re.fullmatch(r"\d+(?:\.\d+)?", out[-1]):
                for u in UNITS:
                    if s.startswith(u, i):
                        out[-1] += u
                        i += len(u)
                        break
            continue
        out.append(ch)
        i += 1
    return out


def _wrap_cell(v, width):
    """按列宽预折行：行首不放标点，数值/单位/扩展名不拆。

    行宽按「列宽 - 3」折、收标点最多挂到「列宽 - 2」：Excel 列宽按数字宽折算，
    中文字略大于 2 个单位，贴着列宽折会被二次折行，把行尾「、」甩成单独一行。
    """
    if not isinstance(v, str) or "\n" in v:
        return v
    width = max(8, int(width))
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
            if not cur:
                cur, cur_w = [tok], w
                continue
            if cur_w + w <= hard:
                cur.append(tok)
                cur_w += w
                continue
            moved = [tok]
            # 只把「行尾最后一个实词」连同收标点挪到下一行；按「和不超列宽」循环
            # 是不变量，永远不成立，会把整行掏空、只剩标点起行。
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


def _fit_page_width(ws, head, caps, page_units=PAGE_UNITS_LANDSCAPE):
    """列宽按「表头 + 数据最宽」算、长文本列封顶，再整体压进一页宽。"""
    from openpyxl.utils import get_column_letter
    mins, widths = {}, {}
    for i, name in enumerate(head, 1):
        mins[i] = _cell_units(str(name or "")) + 2
    for i, name in enumerate(head, 1):
        vals = [ws.cell(r, i).value for r in range(2, ws.max_row + 1)]
        data = max([_cell_units(str(v)) for v in vals if isinstance(v, str)] or [0])
        tok = max([max([_cell_units(t) for t in _tokens(str(v)) if t.strip()] or [0])
                   for v in vals if isinstance(v, str)] or [0])
        mins[i] = max(mins[i], min(tok + 2, page_units // 3))
        widths[i] = max(mins[i], min(caps.get(str(name), 40), data + 1))
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
    """A4 横向 + 一页宽 + 表头每页重复：列不再被切成列带。"""
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = "1:1"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.47   # ≈1.2cm
    ws.page_margins.top = ws.page_margins.bottom = 0.4    # ≈1cm
    ws.page_margins.header = ws.page_margins.footer = 0.2


def layout_sheet(ws, caps):
    """列宽压进一页、长文本按列宽预折行、A4 横向打印设置。"""
    from openpyxl.utils import get_column_letter
    head = [c.value for c in ws[1]]
    _fit_page_width(ws, head, caps)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, str):
                w = ws.column_dimensions[get_column_letter(c.column)].width or 12
                c.value = _wrap_cell(c.value, w)
            c.alignment = Alignment(wrap_text=True, vertical="top")
    _page_setup(ws)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--materials", required=True, help="材料齐套表.json")
    ap.add_argument("--chapters", required=True, help="分章初稿目录或成稿正文")
    ap.add_argument("--facts", required=False, default="")
    ap.add_argument("--delivery-root", default="",
                    help="投递区路径；只取末级目录名（如 输入文件）拼相对落位目录，不写盘符")
    ap.add_argument("--data-gap", action="append", default=[],
                    help="01 数据缺口登记.json / 06 叙述缺口.json（可多次传入），并入材料缺口表")
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)

    delivery_root = args.delivery_root or ""
    # 这份清单是对外附件，落位目录统一写项目内相对路径（输入文件\…），
    # 与本机盘符无关——收件人照着放到项目输入目录即可。
    delivery_name = Path(delivery_root).name if delivery_root else ""

    def dest_rel(sub: str) -> str:
        if not delivery_name:
            return sub or "—"
        return str(Path(delivery_name) / sub) if sub else delivery_name

    wb = Workbook()
    ws = wb.active
    ws.title = "材料缺口"
    ws.append(["编号", "材料名称", "级别", "状态", "说明", "找谁要", "落位目录（补齐后放到这里）"])
    mats = json.load(open(args.materials, encoding="utf-8")) if os.path.exists(args.materials) else {}
    for it in mats.get("items", []):
        if it.get("status") in ("缺失", "疑似缺失", "待补"):
            sub = str(it.get("delivery_dir") or "").strip()
            dest = dest_rel(sub)
            ws.append([it.get("id"), it.get("name"), it.get("level"), it.get("status"),
                       it.get("why", ""), it.get("source") or "—", dest])
    # 01 低置信度/不可读、06 叙述推导缺口：一并进材料缺口表，让缺口清单是完整的一本。
    for path in args.data_gap or []:
        if not path or not os.path.exists(path):
            continue
        try:
            d = json.load(open(path, encoding="utf-8"))
        except Exception:
            continue
        for i, item in enumerate(d.get("lowConfidence") or [], 1):
            ws.append([f"L{i:02d}", str(item.get("file") or ""), "数据",
                       "低置信度入档", str(item.get("reason") or ""),
                       "提供方可解释/补件", dest_rel("")])
        for i, item in enumerate(d.get("unreadable") or [], 1):
            ws.append([f"G{i:02d}", str(item.get("file") or ""), "数据",
                       "文件不可读", str(item.get("reason") or ""),
                       "提供方可读版本", dest_rel("")])
        for i, item in enumerate(d.get("narrativeGaps") or [], 1):
            ws.append([f"N{i:02d}", str(item.get("field") or ""), "叙述",
                       "正文已跳过", str(item.get("note") or "材料侧推不出"),
                       "项目申报材料/信息卡", dest_rel("")])
    style_header(ws, 7)

    # 「找谁要 → 落位目录」映射：从材料齐套表反查。
    # 章节缺项表里的「建议来源」是自然语言（"区发改部门、…"），
    # 与材料表的 source 大体同源，故按关键词包含关系反查所属目录；
    # 查不到时归到「其他补充材料」目录。
    source_dir: dict[str, list[tuple[str, str]]] = {}
    for it in mats.get("items", []):
        src = str(it.get("source") or "").strip()
        sub = str(it.get("delivery_dir") or "").strip()
        if src and sub:
            source_dir.setdefault(src, []).append((sub, str(it.get("name") or "")))
    all_subs = sorted({sub for pairs in source_dir.values() for sub, _ in pairs})

    def dest_for(source_text: str) -> str:
        """按「找谁要」推出唯一落位目录（项目内相对路径）：取来源里排在最前的主责提供方。

        一个单元里只给一个目录；多命中的其余目录不并列，只在内部日志留痕。
        """
        if not delivery_name:
            return "—"
        text = source_text or ""
        hits = []  # (在来源文本中的位置, 子目录)
        for src, pairs in source_dir.items():
            # 「区发改部门」能命中 source「建设单位／发改、住建主管部门」里的片段
            for frag in re.split(r"[／/、,，]", src):
                frag = frag.strip()
                if len(frag) >= 2 and frag in text:
                    pos = text.find(frag)
                    hits.extend((pos, sub) for sub, _ in pairs if sub)
        if not hits:
            return dest_rel("其他补充材料")
        hits.sort(key=lambda x: x[0])
        subs = []
        for _, sub in hits:
            if sub not in subs:
                subs.append(sub)
        if len(subs) > 1:
            print(f"  [落位目录] 来源「{text}」命中 {len(subs)} 个目录，"
                  f"取主责：{subs[0]}；其余不显示：{'、'.join(subs[1:])}")
        return dest_rel(subs[0])

    ws2 = wb.create_sheet("章节缺项")
    ws2.append(["章节", "编号", "缺项内容", "建议来源", "落位目录（补齐后放到这里）"])
    seen_gap = set()

    def collect_gap_rows(txt, chapter):
        """从一段文本里抽缺项行。返回 [(编号, 内容, 来源)]。"""
        out = []
        # ① 表格行格式（成稿里的实际形态）：
        #    | 缺项-A01 | 缺什么 | 找谁补 | 现有替代性判断 |
        #    逐行匹配（正则带 ^ 锚点，不加 re.M 时只在首行生效）
        for ln in txt.splitlines():
            m = GAP_ROW_RE.match(ln.strip())
            if not m:
                continue
            cells = [c.strip() for c in m.group(3).split("|")]
            while len(cells) < 2:
                cells.append("")
            gid = f"缺项-{(m.group(1) or chapter).upper()}{int(m.group(2)):02d}"
            out.append((gid, cells[0], cells[1]))
        # ② 行内格式（旧形态，保留兼容）：
        #    【缺项-A01：内容｜建议来源：来源】
        for m in GAP_RE.finditer(txt):
            out.append((m.group(1), m.group(2), m.group(3)))
        return out

    chap_path = Path(args.chapters)
    if chap_path.is_file():
        # 单文件模式：成稿正文（含汇总后的附件四缺项清单）
        txt = chap_path.read_text(encoding="utf-8")
        for gid, content, source in collect_gap_rows(txt, ""):
            if gid in seen_gap:
                continue
            ws2.append([gid[3:4], gid, content, source, dest_for(source)])
            seen_gap.add(gid)
    elif chap_path.is_dir():
        for fn in sorted(os.listdir(chap_path)):
            if not fn.endswith(".md") or fn.startswith("_"):
                continue
            chapter = fn.split("_")[0]
            txt = (chap_path / fn).read_text(encoding="utf-8")
            for gid, content, source in collect_gap_rows(txt, chapter):
                if gid in seen_gap:
                    continue
                ws2.append([chapter, gid, content, source, dest_for(source)])
                seen_gap.add(gid)
    style_header(ws2, 5)

    ws3 = wb.create_sheet("事实缺失")
    ws3.append(["章节", "缺失内容", "落位目录（补齐后放到这里）"])
    facts = json.load(open(args.facts, encoding="utf-8")) if args.facts and os.path.exists(args.facts) else {}
    missing = facts.get("missing") or []
    for m in missing:
        if isinstance(m, dict):
            # 事实缺失条目的字段是 section/heading/note（title 为空），
            # 旧写法只取 title/reason → 导出后「缺失内容」整列为空。
            detail = (m.get("heading") or m.get("title") or m.get("note")
                      or m.get("reason") or "")
            note = m.get("note") or m.get("reason") or ""
            if note and note != detail:
                detail = f"{detail}（{note}）" if detail else note
            ws3.append([m.get("section", ""), detail, dest_for("")])
        else:
            ws3.append(["", str(m), dest_for("")])
    style_header(ws3, 3)

    # 三张表统一按 A4 横向一页宽排版：材料名称/说明/缺项内容/缺失内容封顶，
    # 落位目录给 20 个单位（够放「输入文件\其他补充材料」），其余按内容自适应。
    layout_sheet(ws, {"编号": 8, "材料名称": 26, "说明": 38, "找谁要": 20,
                      "落位目录（补齐后放到这里）": 20})
    layout_sheet(ws2, {"章节": 6, "编号": 12, "缺项内容": 38, "建议来源": 26,
                       "落位目录（补齐后放到这里）": 20})
    layout_sheet(ws3, {"章节": 8, "缺失内容": 52,
                       "落位目录（补齐后放到这里）": 20})
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    wb.save(args.out)
    print(f"[02] 数据缺口清单 → {args.out}"
          f"（材料 {ws.max_row - 1} / 章节 {ws2.max_row - 1} / 事实 {ws3.max_row - 1}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
