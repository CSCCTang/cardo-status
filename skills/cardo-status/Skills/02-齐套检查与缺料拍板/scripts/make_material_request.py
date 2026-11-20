#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_material_request.py -- 把「材料齐套表」转成一份可发给业务专家确认的《输入材料补齐确认单》。

输入（只读）：
  * `05_过程\\齐套检查\\材料齐套表.json`（check_materials.py 产出，schema cardo/material-gap/v1）
  * `05_过程\\分类清单\\分类清单.json`（可选；用于列出"已收到哪些材料"，不传则该节省略）
  * `--extra <md>`（可选；人工撰写、要一并请专家确认的「其他口径事项」）

输出：
  * `<out>`：.docx（A4、宋体/黑体、表格带框线、表头跨页重复）
  * `--md`：同名 markdown（便于改字/留档/再生成）

设计约束：
  * 不依赖任何模板 docx（这是一份**工作单据**，不是标准报告，不能用报告模板的章节编号样式）；
  * 零硬编码：项目名、方向、投递区、收件人全部来自参数或齐套表 JSON；
  * 只读输入、只写指定输出；不改投递区、不改项目其它文件。

用法（T20 起可用 --project-dir，输入/输出相对路径以项目目录为基准）：
    python make_material_request.py --gap-json <材料齐套表.json> --out <确认单.docx> \\
        [--md <确认单.md>] [--classification <分类清单.json>] [--extra <其他事项.md>] \\
        [--to <收件人>] [--report-name <报告文件名>] [--date YYYY-MM-DD]
    python make_material_request.py --project-dir <用户自选项目目录> [--md 同名.md]
        # 缺省 --gap-json=<项目>/05_过程/齐套检查/材料齐套表.json
        # 缺省 --out=<项目>/04_成果报告/对外单据/输入材料补齐确认单-<日期>.docx（对外交付物，不进 05_过程）
        # 缺省 --md = 同名 .md

退出码：0 成功；1 运行失败（含只读目录写入失败）；2 参数/输入错误。
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _common import (  # noqa: E402
    UserProjectError,
    disp_path,
    resolve_user_project_dir,
)

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

__version__ = "1.0.0"
DEFAULT_EAST_ASIA = "宋体"
HEAD_EAST_ASIA = "黑体"
GRID_STYLE = "Table Grid"

# 缺失/疑似材料表列：一眼能看出「要什么、为什么要、现在报告怎么写、找谁要、怎么确认」
MISSING_COLS = ["编号", "材料名称", "为什么需要（影响报告章节）", "目前报告的替代写法",
                "建议提供方", "落位目录", "专家确认"]
SUSPECT_COLS = ["编号", "材料名称", "现有疑似文件", "请确认的问题", "专家意见"]
SUGGEST_COLS = ["编号", "材料名称", "作用", "建议提供方", "是否需要"]

CN_NUM = "一二三四五六七八九十"
ASK_MARKS = "①②③④⑤"


def section_plan(head: dict, extra_md: str | None) -> list[dict]:
    """按本次实际会出现的栏目排序并编中文序号。

    二/四/五/六节是条件节（没有对应材料/人工事项就整节省略），
    序号必须跟着「实际生成」走，不能写死成四、六，否则导语会指向不存在的章节。
    """
    plan: list[dict] = []

    def add(code: str, title: str) -> None:
        n = len(plan) + 1
        plan.append({"code": code, "title": title,
                     "no": CN_NUM[n - 1] if n <= len(CN_NUM) else str(n)})

    add("overview", "材料齐套总体情况")
    if head.get("received"):
        add("received", "已收到的材料（无需重复提供）")
    add("missing", "必需材料·缺失清单（建议优先补齐）")
    if head.get("suspect"):
        add("suspect", "疑似材料·请确认是否可用")
    if head.get("suggested"):
        add("suggested", "建议补充材料（有则更扎实，没有可先出稿）")
    if extra_md:
        add("extra", "其他需要一并确认的口径问题")
    add("howto", "回填方式与后续处理")
    return plan


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def read_text(path: str | os.PathLike) -> str:
    return Path(path).read_text(encoding="utf-8-sig")


def read_json(path: str | os.PathLike) -> dict:
    try:
        return json.loads(read_text(path))
    except json.JSONDecodeError as exc:
        raise SystemExit(f"错误：JSON 解析失败 {path}：{exc}")


# --------------------------------------------------------------------------- #
# docx 基础排版
# --------------------------------------------------------------------------- #
def set_cell_text(cell, text: str, *, bold: bool = False, size: float = 9.0,
                  center: bool = False) -> None:
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = DEFAULT_EAST_ASIA
    run._element.rPr.rFonts.set(qn("w:eastAsia"), DEFAULT_EAST_ASIA)
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.space_after = Pt(1)


def shade_cell(cell, hex_color: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def mark_header_row(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def mark_row_cant_split(row) -> None:
    """整行不跨页：材料行折行多，跨页后下一页只剩半截且首列空白，读起来像断表。"""
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def _disp_width(text: str) -> int:
    total = 0
    for ch in text or "":
        total += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return total


def _min_widths(usable_twips: int, header: list, rows: list, ncols: int, font_size: float) -> list:
    """每列的最小可读宽度（twips）：至少两个字宽、最长不可断 ASCII 串、外加单元格内边距。"""
    pad = 300  # 0.19cm×2 内边距 + 边框，折 twips
    mins = []
    for c in range(ncols):
        texts = [str(header[c]) if c < len(header) else ""]
        texts += [str(r[c]) for r in rows if c < len(r)]
        tokens = [tok for txt in texts for tok in re.findall(r"[A-Za-z0-9_.%\-/\\]+", txt)]
        ascii_w = max([len(t) for t in tokens] or [1]) * font_size * 10
        two_cjk = 2 * font_size * 20
        mins.append(int(max(ascii_w, two_cjk) + pad))
    # 最小宽度总和不得超过可用宽度：超了就按比例压
    total = sum(mins)
    if total > usable_twips:
        mins = [int(m * usable_twips / total) for m in mins]
    return mins


def column_widths(usable_twips: int, header: list, rows: list, ncols: int,
                  weights: list | None = None, font_size: float = 8.5) -> list:
    """列宽：优先用调用方给的相对权重，否则按内容 85 分位；合计 = 可用宽度。

    每列强制不低于最小可读宽度，避免「编号」「住建部门/街道办」被挤成逐字竖排；
    差额从最宽的列里扣。
    """
    if ncols <= 0:
        return []
    if weights and len(weights) == ncols:
        total = float(sum(weights)) or 1.0
        raw = [usable_twips * w / total for w in weights]
    else:
        scores = []
        for c in range(ncols):
            lens = [_disp_width(header[c] if c < len(header) else "")]
            for r in rows:
                lens.append(_disp_width(r[c] if c < len(r) else ""))
            lens.sort()
            scores.append(float(max(lens[int(round(0.85 * (len(lens) - 1)))], 4)))
        total = sum(scores) or 1.0
        raw = [usable_twips * s / total for s in scores]
    ints = [int(round(x)) for x in raw]
    mins = _min_widths(usable_twips, header, rows, ncols, font_size)
    for c in range(ncols):
        ints[c] = max(ints[c], mins[c])
    overflow = sum(ints) - usable_twips
    while overflow > 0:
        # 从可压缩空间最大的列里扣，且不低于其最小值
        candidates = [(ints[i] - mins[i], i) for i in range(ncols) if ints[i] > mins[i]]
        if not candidates:
            break
        _, idx = max(candidates)
        take = min(overflow, ints[idx] - mins[idx])
        ints[idx] -= take
        overflow -= take
    ints[-1] += usable_twips - sum(ints)
    return ints


def add_table(doc, header: list, rows: list, widths: list, *, font_size: float = 9.0) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(header))
    table.style = GRID_STYLE
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for c, w in enumerate(widths):
        table.columns[c].width = Pt(w / 20.0)  # twips -> pt
    for r_idx, row_data in enumerate([header] + rows):
        row = table.rows[r_idx]
        for c_idx in range(len(header)):
            cell = row.cells[c_idx]
            cell.width = Pt(widths[c_idx] / 20.0)
            text = row_data[c_idx] if c_idx < len(row_data) else ""
            set_cell_text(cell, str(text), bold=(r_idx == 0), size=font_size,
                          center=(r_idx == 0 or c_idx == 0))
            if r_idx == 0:
                shade_cell(cell, "EDEDED")
        if r_idx == 0:
            mark_header_row(row)
        mark_row_cant_split(row)


def add_heading(doc, text: str, level: int = 1) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(13 if level == 1 else 11)
    run.font.name = HEAD_EAST_ASIA
    run._element.rPr.rFonts.set(qn("w:eastAsia"), HEAD_EAST_ASIA)


def add_para(doc, text: str, *, size: float = 10.5, bold: bool = False,
             indent: bool = False, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.25
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    if align is not None:
        p.alignment = align
    # 支持 **加粗** 片段，其余按正文；不把 markdown 标记漏进 docx
    for idx, chunk in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not chunk:
            continue
        run = p.add_run(chunk)
        run.bold = bold or (idx % 2 == 1)
        run.font.size = Pt(size)
        run.font.name = DEFAULT_EAST_ASIA
        run._element.rPr.rFonts.set(qn("w:eastAsia"), DEFAULT_EAST_ASIA)
    return p


def build_doc(head: dict, extra_md: str | None) -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)
    usable = int((21.0 - 4.4) / 2.54 * 1440)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(head["title"])
    run.bold = True
    run.font.size = Pt(20)
    run.font.name = HEAD_EAST_ASIA
    run._element.rPr.rFonts.set(qn("w:eastAsia"), HEAD_EAST_ASIA)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(head["project_line"])
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = DEFAULT_EAST_ASIA
    run._element.rPr.rFonts.set(qn("w:eastAsia"), DEFAULT_EAST_ASIA)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run(head["subtitle"])
    run.font.size = Pt(10.5)
    run.font.name = DEFAULT_EAST_ASIA
    run._element.rPr.rFonts.set(qn("w:eastAsia"), DEFAULT_EAST_ASIA)

    for line in head["intro"]:
        add_para(doc, line, indent=True)

    plan = head["sections"]
    no = {s["code"]: s["no"] for s in plan}

    # 一、总体情况
    add_heading(doc, f"{no['overview']}、材料齐套总体情况")
    s = head["summary"]
    add_para(doc, (
        f"本项目按「{head['direction']}」方向检查，投递区现有材料 {s.get('files_scanned', '—')} 份；"
        f"对照该方向的必备材料清单（共 {s.get('required_total', '—')} 项必需），"
        f"已齐 {s.get('required_hit', '—')} 项（齐套率 {head['rate_text']}），"
        f"待确认 {s.get('required_suspect', 0)} 项，缺失 {max((s.get('required_missing', 0) or 0) - (s.get('required_suspect', 0) or 0), 0)} 项；"
        f"建议补充材料已齐 {s.get('suggested_hit', 0)}/{s.get('suggested_total', 0)}。"
    ), indent=True)
    add_para(doc, (
        ("现状分析报告已按模板出稿：正文按「初步判定＋判断依据」成文，判断分"
         "「（可核对）」「（推断，需现场核实）」两类标注，"
         "未能取得原始数据的定量口径集中汇总为《数据缺口清单》（随报告一并交付），正文不再逐处标注缺项。"
         if head.get("report_ready") else
         "本单在报告成稿前生成，先把需要补的材料一次说清；补齐后重跑即可把对应章节由定性判断升级为定量结论。")
        + "本次请您确认的，是这份清单里**属于输入材料缺失**的部分——哪些能补、向谁补、"
        "哪些确认无需补充。"
    ), indent=True)

    # 二、已收到材料
    if head["received"]:
        add_heading(doc, f"{no['received']}、已收到的材料（无需重复提供）")
        rows = [[str(i + 1), item["rel"], item.get("category") or "—", item.get("size") or "—"]
                for i, item in enumerate(head["received"])]
        add_table(doc, ["序号", "文件", "初判类别", "大小"], rows,
                  column_widths(usable, ["序号", "文件", "初判类别", "大小"], rows, 4,
                                weights=[1, 8, 2.4, 1.6]), font_size=8.5)
        add_para(doc, "说明：以上为投递区现有全部材料，已按类别归集；这些材料不需要再提供。", size=9.5)

    # 三、必需·缺失
    add_heading(doc, f"{no['missing']}、必需材料·缺失清单（建议优先补齐）")
    add_para(doc, (
        "补齐材料请放到本项目输入目录（输入文件）下对应的子目录，见下表「落位目录」列（按目录放好即可被自动识别）；"
        "放好后重跑流水线即可自动识别。"
    ), size=9.5)
    add_para(doc, (
        "下列材料属于「做一份标准现状分析报告必须有」的输入。缺项对应的章节目前只能写定性判断，"
        "无法给出量化结论。请在最后一列标注：可用（已提供或不需提供）／需补／不需要。"
    ), size=9.5)
    rows = []
    for item in head["missing"]:
        rows.append([
            item["id"], item["name"],
            item.get("why_full") or item.get("why") or "；".join(item.get("chapters") or []),
            item.get("fallback") or "定性描述 + 缺项标注",
            item.get("source") or "—",
            item.get("dest") or "—",
            "",
        ])
    add_table(doc, MISSING_COLS, rows,
              column_widths(usable, MISSING_COLS, rows, len(MISSING_COLS),
                            weights=[1.0, 2.9, 4.0, 3.1, 2.2, 3.2, 2.6]), font_size=8.5)
    # 按「找谁要」分组的落位说明：同一提供方的材料放到同一目录，一次拷贝就够
    groups: dict[str, list[str]] = {}
    for item in head["missing"]:
        key = item.get("source") or "—"
        groups.setdefault(key, []).append(item.get("dest") or "—")
    if groups:
        add_para(doc, "按提供方分组的落位目录（补齐后按目录放好，重跑流水线即可被自动识别）：",
                 size=9.5, bold=True)
        for n, (who, dests) in enumerate(sorted(groups.items()), 1):
            uniq = sorted(set(dests))
            add_para(doc, f"{n}. {who} → " + "；".join(uniq), size=9.0)

    # 四、疑似待确认
    if head["suspect"]:
        add_heading(doc, f"{no['suspect']}、疑似材料·请确认是否可用")
        add_para(doc, (
            "下列材料在投递区内有名称相近的文件，但无法确认是否就是所需材料（可能是其他项目的图纸/报告）。"
            "请确认：① 该文件是否可用作本项目依据；② 若不可用，是否有独立的正式版本。"
        ), size=9.5)
        rows = [[item["id"], item["name"], item.get("matched_text") or "—",
                 item.get("question") or "现有文件能否作为本项目依据？", ""]
                for item in head["suspect"]]
        add_table(doc, SUSPECT_COLS, rows,
                  column_widths(usable, SUSPECT_COLS, rows, len(SUSPECT_COLS),
                                weights=[1.1, 3.4, 4.4, 4.4, 2.0]), font_size=8.5)

    # 五、建议补充
    if head["suggested"]:
        add_heading(doc, f"{no['suggested']}、建议补充材料（有则更扎实，没有可先出稿）")
        rows = [[item["id"], item["name"], item.get("why") or "—",
                 item.get("source") or "—", ""] for item in head["suggested"]]
        add_table(doc, SUGGEST_COLS, rows,
                  column_widths(usable, SUGGEST_COLS, rows, len(SUGGEST_COLS),
                                weights=[1.1, 3.4, 4.6, 3.0, 2.2]), font_size=8.5)

    # 六、其他待确认（人工撰写）
    if extra_md:
        add_heading(doc, f"{no['extra']}、其他需要一并确认的口径问题")
        n_item = 0
        for block in parse_extra(extra_md):
            if block["kind"] == "h2":
                add_heading(doc, block["text"], level=2)
                n_item = 0
            elif block["kind"] == "li":
                n_item += 1
                add_para(doc, f"{n_item}. " + block["text"], size=10)
            else:
                add_para(doc, block["text"], indent=True)

    # 七、回填方式（另起一页，避免整页只剩一行尾注）
    doc.add_page_break()
    add_heading(doc, f"{no['howto']}、回填方式与后续处理")
    howto = head["howto"]
    for i, line in enumerate(howto):
        p = add_para(doc, f"{i + 1}. " + line, size=10)
        if i == len(howto) - 1:
            p.paragraph_format.keep_with_next = True
    add_para(doc, head["footer"], size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)
    return doc


def parse_extra(md: str) -> list:
    blocks: list = []
    for raw in md.splitlines():
        line = raw.strip()
        if not line or line in ("---",):
            continue
        if line.startswith("## "):
            blocks.append({"kind": "h2", "text": line[3:].strip()})
        elif line.startswith(("- ", "* ")):
            blocks.append({"kind": "li", "text": line[2:].strip()})
        else:
            blocks.append({"kind": "p", "text": line})
    return blocks


# --------------------------------------------------------------------------- #
# 组装
# --------------------------------------------------------------------------- #
def collect(gap: dict, classification: dict | None, gap_map: dict | None = None,
            delivery_rel: str = "", other_files: set[str] | None = None) -> dict:
    items = gap.get("items") or []
    summary = gap.get("summary") or {}
    gap_map = gap_map or {}
    other_files = other_files or set()
    # 落位目录只用项目内相对路径（输入文件\…），绝对路径对收件人无用且会带出本机目录
    delivery_rel = (delivery_rel or gap.get("delivery") or "输入文件").strip().strip("/\\")

    received = []
    if classification:
        for item in classification.get("files") or []:
            rel = item.get("rel") or ""
            cat = item.get("category_label") or ""
            # 投递区常混入别的项目材料（如某项目与本项目同批移交）：
            # 识别结果已列出这些文件，这里如实标注，避免专家误以为本项目引用了别片区台账。
            if rel.replace("/", "\\") in other_files:
                cat = (cat + "（非本项目材料，未用于本报告）") if cat else "非本项目材料，未用于本报告"
            received.append({
                "rel": rel,
                "category": cat,
                "size": item.get("size_h") or "",
            })

    missing, suspect, suggested = [], [], []
    for item in items:
        level = item.get("level")
        status = item.get("status")
        matched = item.get("matched") or []
        if status == "已齐":
            continue
        mid = item.get("id") or ""
        gap_ref = str(gap_map.get(mid, "") or "").replace("缺项-", "").strip()
        # 落位目录：项目内相对路径（输入文件\管网与市政），不放本机绝对路径
        sub = (item.get("delivery_dir") or "").strip().strip("/\\")
        dest = f"{delivery_rel}\\{sub}" if sub else delivery_rel
        base = {
            "id": mid,
            "name": item.get("name") or "",
            "why": item.get("why") or "",
            "chapters": item.get("chapters") or [],
            "source": item.get("source") or "",
            "gap_ref": gap_ref,
            "delivery_dir": sub,
            "dest": dest,
        }
        ref_note = f"（报告中已标注：缺项-{gap_ref}）" if gap_ref else ""
        if status == "待确认":
            files = matched or []
            base["matched_text"] = (
                "；".join(files[:2]) + (f" 等 {len(files)} 份" if len(files) > 2 else "")
            ) if files else "（未定位到具体文件）"
            base["question"] = (
                f"上述文件能否作为本项目「{base['name']}」的依据？"
                f"若不能，请说明另需哪份材料，或确认本项目不需要。" + ref_note
            )
            suspect.append(base)
        elif level == "必需":
            chapters = "、".join(base["chapters"])
            base["why_full"] = base["why"] + (f"（影响：{chapters}）" if chapters else "")
            base["fallback"] = "仅定性描述" + (f"；报告已标缺项-{gap_ref}" if gap_ref else "＋缺项标注")
            missing.append(base)
        else:
            suggested.append(base)

    rate = summary.get("rate")
    rate_text = f"{float(rate) * 100:.0f}%" if isinstance(rate, (int, float)) else "—"
    return {
        "summary": summary,
        "rate_text": rate_text,
        "direction": gap.get("direction") or "—",
        "project": gap.get("project") or "",
        "project_name": gap.get("project_name") or "",
        "delivery": gap.get("delivery") or "",
        "received": received,
        "missing": missing,
        "suspect": suspect,
        "suggested": suggested,
    }


def to_markdown(head: dict) -> str:
    s = head["summary"]
    plan = head["sections"]
    no = {x["code"]: x["no"] for x in plan}
    lines = [f"# {head['title']}", "", head["project_line"], "", head["subtitle"], ""]
    lines += [line + "\n" for line in head["intro"]]
    lines += [f"## {no['overview']}、材料齐套总体情况", "",
              f"- 方向：{head['direction']}",
              f"- 投递区：{head['delivery']}（现有 {s.get('files_scanned', '—')} 份材料）",
              f"- 必需 {s.get('required_hit', 0)}/{s.get('required_total', 0)} 齐（{head['rate_text']}）；"
              f"待确认 {s.get('required_suspect', 0)}；缺失 {len(head['missing'])}",
              f"- 建议 {s.get('suggested_hit', 0)}/{s.get('suggested_total', 0)} 齐", ""]
    if head["received"]:
        lines += [f"## {no['received']}、已收到的材料（无需重复提供）", "",
                  "| 序号 | 文件 | 初判类别 | 大小 |", "|---|---|---|---|"]
        for i, item in enumerate(head["received"]):
            lines.append(f"| {i + 1} | {item['rel']} | {item.get('category') or '—'} | {item.get('size') or '—'} |")
        lines.append("")
    lines += [f"## {no['missing']}、必需材料·缺失清单（建议优先补齐）", "",
              "补齐材料请放到本项目输入目录（输入文件）下对应的子目录，见下表「落位目录」列（按目录放好即可被自动识别）；"
              "放好后重跑流水线即可自动识别。", "",
              "| " + " | ".join(MISSING_COLS) + " |", "|" + "---|" * len(MISSING_COLS)]
    for item in head["missing"]:
        lines.append("| " + " | ".join([
            item["id"], item["name"],
            item.get("why_full") or item.get("why") or "；".join(item.get("chapters") or []),
            item.get("fallback") or "定性描述 + 缺项标注",
            item.get("source") or "—", item.get("dest") or "—", "",
        ]) + " |")
    lines.append("")
    # 按提供方分组的落位目录
    groups: dict[str, list[str]] = {}
    for item in head["missing"]:
        groups.setdefault(item.get("source") or "—", []).append(item.get("dest") or "—")
    if groups:
        lines += ["按提供方分组的落位目录（补齐后按目录放好，重跑流水线即可被自动识别）：", ""]
        for who, dests in sorted(groups.items()):
            lines.append(f"- {who} → " + "；".join(sorted(set(dests))))
        lines.append("")
    if head["suspect"]:
        lines += [f"## {no['suspect']}、疑似材料·请确认是否可用", "",
                  "| " + " | ".join(SUSPECT_COLS) + " |", "|" + "---|" * len(SUSPECT_COLS)]
        for item in head["suspect"]:
            lines.append("| " + " | ".join([
                item["id"], item["name"], item.get("matched_text") or "—",
                item.get("question") or "", ""]) + " |")
        lines.append("")
    if head["suggested"]:
        lines += [f"## {no['suggested']}、建议补充材料", "",
                  "| " + " | ".join(SUGGEST_COLS) + " |", "|" + "---|" * len(SUGGEST_COLS)]
        for item in head["suggested"]:
            lines.append("| " + " | ".join([
                item["id"], item["name"], item.get("why") or "—",
                item.get("source") or "—", ""]) + " |")
        lines.append("")
    if head.get("extra_md"):
        extra = head["extra_md"].strip()
        # 与 docx 保持一致：extra 里的一级小节降为三级，并挂到口径问题节下
        extra = "\n".join(
            ("### " + ln[3:]) if ln.startswith("## ") else ln for ln in extra.splitlines()
        )
        lines += [f"## {no['extra']}、其他需要一并确认的口径问题", "", extra, ""]
    lines += [f"## {no['howto']}、回填方式与后续处理", ""]
    lines += [f"- {x}" for x in head["howto"]]
    lines += ["", head["footer"]]
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    ap = argparse.ArgumentParser(description="生成《输入材料补齐确认单》docx + md（只读齐套表）")
    ap.add_argument("--project-dir", default="",
                    help="用户自选项目目录（T20 最高优先级；相对输入/输出以它为基准，可省略 --gap-json/--out）")
    ap.add_argument("--gap-json", default="", help="材料齐套表 JSON（cardo/material-gap/v1）")
    ap.add_argument("--out", default="", help="输出 .docx 路径")
    ap.add_argument("--md", default="", help="可选：同时输出 markdown")
    ap.add_argument("--classification", default="", help="可选：分类清单 JSON（列出已收到材料）")
    ap.add_argument("--project-id", default="",
                    help="可选：项目识别.json；用于把其它项目材料在已收到清单里标注出来")
    ap.add_argument("--extra", default="", help="可选：其他待确认事项 markdown 文件")
    ap.add_argument("--gap-map", default="",
                    help="可选：材料编号 → 报告缺项编号 的映射 JSON（如 {\"M07\": \"缺项-B03、B20\"}）")
    ap.add_argument("--to", default="", help="收件人（如：贝贝 总）")
    ap.add_argument("--report-name", default="", help="配套送审的报告文件名")
    ap.add_argument("--report-ready", action="store_true",
                    help="报告已出稿（09 步重新生成时使用）；不传则按「成稿前」口径开头")
    ap.add_argument("--date", default="", help="落款日期 YYYY-MM-DD，默认今天")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args(argv)

    # T20：--project-dir → 相对路径以项目目录为基准；缺省 gap-json/out 从项目规范目录取
    base: Path | None = None
    if str(args.project_dir or "").strip():
        try:
            base = resolve_user_project_dir(args.project_dir)
        except UserProjectError as exc:
            print(f"错误：{exc}", file=sys.stderr)
            return 2
        date_hint = args.date or datetime.date.today().strftime("%Y%m%d")
        gap_default = base / "05_过程" / "齐套检查" / "材料齐套表.json"
        # 确认单是**对外交付物**，按 SKILL.md §5.5 落在 04_成果报告/对外单据/，不进 05_过程。
        out_default = base / "04_成果报告" / "对外单据" / f"输入材料补齐确认单-{date_hint}.docx"
        args.gap_json = args.gap_json or str(gap_default)
        args.out = args.out or str(out_default)
        # 缺省同时出同名 markdown（便于改字/留档/再生成）
        if not args.md:
            args.md = str(out_default.with_suffix(".md"))

    def resolve_arg(value: str) -> Path | None:
        if not value:
            return None
        path = Path(value).expanduser()
        if base is not None and not path.is_absolute():
            path = base / path
        return path

    gap_path = resolve_arg(args.gap_json)
    out_path = resolve_arg(args.out)
    if gap_path is None or out_path is None:
        print("错误：请给 --gap-json 与 --out（或改用 --project-dir 用缺省位置）。", file=sys.stderr)
        return 2
    md_path = resolve_arg(args.md)
    classification_path = resolve_arg(args.classification)
    extra_path = resolve_arg(args.extra)
    gap_map_path = resolve_arg(args.gap_map)

    if not gap_path.is_file():
        print(f"错误：齐套表不存在：{disp_path(gap_path) if base else gap_path}", file=sys.stderr)
        return 2
    gap = read_json(gap_path)
    if gap.get("schema") not in (None, "cardo/material-gap/v1"):
        print(f"警告：齐套表 schema={gap.get('schema')!r}，按 v1 结构尽力解析", file=sys.stderr)
    classification = read_json(classification_path) if classification_path and classification_path.is_file() else None
    extra_md = read_text(extra_path) if extra_path and extra_path.is_file() else None

    # 投递区相对路径：确认单是发给建设单位的，落位目录只写「输入文件\…」，不写本机绝对路径
    delivery_rel = (gap.get("delivery") or "输入文件").strip().strip("/\\")
    other_files = set()
    pid_path = resolve_arg(args.project_id) if args.project_id else None
    if pid_path and pid_path.is_file():
        pid = read_json(pid_path)
        # 只有「已判为独立项目」的才标注为本项目之外的材料。
        # 判为区域（area）或尚未判定的（ask）一律不标注——不按名字片段静默排除。
        for op in pid.get("scopeCandidates") or []:
            if str(op.get("decision") or "") != "separate":
                continue
            for f in op.get("files") or []:
                rel = str(f).replace("/", "\\")
                # 包内依据 是本项目随包提供的证据文件，即便被切词算法误判也不标注。
                if rel.startswith("包内依据"):
                    continue
                other_files.add(rel)
    data = collect(gap, classification,
                   read_json(gap_map_path) if gap_map_path and gap_map_path.is_file() else None,
                   delivery_rel, other_files)
    # 抬头优先用项目全称（齐套表的 project_name），缺省用 project（目录名）
    project = data.get("project_name") or data["project"] or "本项目"
    date = args.date or datetime.date.today().strftime("%Y-%m-%d")
    title = "输入材料补齐确认单"
    # 节号按实际生成的栏目动态编：四、六这类条件节缺失时，后续节自动前移。
    plan = section_plan(data, extra_md)
    no = {x["code"]: x["no"] for x in plan}
    asks = [f"第{no['missing']}节里“缺失”的材料，哪些确实可以补到、应向谁要"]
    if "suspect" in no:
        asks.append(f"第{no['suspect']}节里“疑似”的文件，现有那份能否直接作为本项目依据")
    if "extra" in no:
        asks.append(f"第{no['extra']}节里几个口径问题（专名、调研期限、图件要求等）按哪个写法定稿")
    ask_text = "；".join(f"{ASK_MARKS[i]} {a}" for i, a in enumerate(asks))
    head = {
        "title": title,
        "project_line": f"{project}（{data['direction']}）" if data["direction"] not in ("—", "") else project,
        "subtitle": f"《现状分析报告》配套文件 ｜ 送 {args.to or '业务专家'}确认 ｜ {date}",
        "direction": data["direction"],
        "delivery": data["delivery"],
        "summary": data["summary"],
        "rate_text": data["rate_text"],
        "received": data["received"],
        "missing": data["missing"],
        "suspect": data["suspect"],
        "suggested": data["suggested"],
        "report_ready": bool(getattr(args, "report_ready", False)),
        "extra_md": extra_md or "",
        "sections": plan,
        "intro": [
            (f"您好！{project}的《城市更新项目现状分析报告》已按标准模板完成拆解出稿"
             + (f"（同名文件：{args.report_name}）" if args.report_name else "")
             if args.report_ready else
             f"您好！{project}的《城市更新项目现状分析报告》正在编制，先与您对齐所需输入材料")
            + "。为使报告能作为正式成果使用，我们对“做这份报告所需的输入材料”做了一次齐套检查，"
              "把缺口整理成下面这张单子，请您确认。",
            f"请重点确认以下 {len(asks)} 件事：{ask_text}。",
            "确认后我会按您的意见补齐报告：能补的材料写成定量结论，"
            "补不到的按初步判定与「（推断，需现场核实）」成文，并同步更新《数据缺口清单》。",
        ],
        "howto": [
            "请在本单各表最后一列直接批注：可用（已提供或不需提供）／需补／不需要；补充说明写在同格内即可。",
            f"也可以只回复编号，例如：“M07 需补、向住建局要；M06 可用”"
            + (f"；第{no['extra']}节第 1 条按「重大装备」订正" if "extra" in no else "") + "。",
            "材料补齐后放回项目投递区即可（原件不用改名、不用移动），我会重跑齐套检查并更新报告。",
            "若某项您判断“不需要”，也请标注，我会把报告里对应缺项改为“经确认不适用”，避免长期挂着。",
        ],
        "footer": f"编制：项目组　｜　{date}",
    }

    out = out_path
    doc = build_doc(head, extra_md)
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))
    except OSError as exc:
        print(f"错误：无法写入确认单（可能是只读目录、无权限或文件被占用）：{disp_path(out)}", file=sys.stderr)
        print(f"      原因：{type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
    print(f"[ok] docx: {disp_path(out) if base else out}")

    if md_path is not None:
        try:
            md_path.parent.mkdir(parents=True, exist_ok=True)
            md_path.write_text(to_markdown(head), encoding="utf-8", newline="\n")
        except OSError as exc:
            print(f"错误：无法写入 markdown（可能是只读目录、无权限）：{disp_path(md_path)}", file=sys.stderr)
            return 1
        print(f"[ok] md  : {disp_path(md_path) if base else md_path}")

    # 自检：回读
    check = Document(str(out))
    tables = len(check.tables)
    paras = len([p for p in check.paragraphs if p.text.strip()])
    print(f"[verify] 段落 {paras}、表格 {tables}（期望 3–5 张：已收到/缺失/疑似/建议）")
    if tables < 2:
        print("[verify] FAIL：表格数量异常", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())