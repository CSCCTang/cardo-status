#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""assemble_chapters.py -- 把分章成文的 md 片段组装成一份可渲染的正文（通用、零硬编码）。

约定（与 cards-status-report 的分章产物格式一致）：
  * 分章文件命名：`<章字母>_<章名>.md`，章字母为 A/B/C/D…（同一字母多文件时取字典序第一个）；
    章序 = 字母序，因此 A/B/C/D/E 就是正文的先后顺序。
  * 每个分章文件内可以有若干“过程段”，它们的二级标题形如：
        ## 本节缺项表      （每个分章自己的缺项表，编号缺项-A01 / 缺项-B07 …）
        ## 来源映射        （本节内容取自原文哪一节，过程留档，不进交付正文）
    这两个段会被剥离：缺项表被汇总成正文末尾的「数据缺口清单」，来源映射被单独落到
    `--mapping-out` 指定的文件里（供追溯），都不进交付正文。
  * 正文里的内联标记形如 `【缺项-01：缺什么｜建议来源：谁】`：脚本按所属分章补上字母前缀
    （A 章 → `缺项-A01`），保证全篇编号唯一、可与汇总表一一对应。
  * 正文一级/二级标题、段落、markdown 表格原样保留，交给 build_report_docx.py 渲染；
    `（数据来源：…）` 行与标注源的 `〔锚点 …〕` 依据块一律剥离（出处只进《依据对照表》）。
    脚本不改变任何正文用字。

用法：
    python assemble_chapters.py --extract-dir <分章目录> --out <body.md>
                                [--mapping-out <来源映射.md>]
                                [--gap-section <二级标题关键词，默认 本节缺项表>]
                                [--mapping-section <关键词，默认 来源映射>]
                                [--gap-heading <汇总节标题，默认 数据缺口清单>]
                                [--append-to-chapter <字母，默认最后一章>]

退出码：0 成功；2 参数/输入错误（缺目录、缺分章文件、找不到可汇总的缺项）。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _skill_paths  # noqa: E402  （同目录资源定位底座，零第三方依赖）

# 检索台账的读写规则只有一份（09 下的 _ledger.py）。这里只用来生成"附件六"，
# 校验权威仍在 09 的 CHK-CITE。找不到 09 时降级为不生成附件，不影响正文组装。
_SKILLS = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    sys.path.insert(0, os.path.join(_SKILLS, "09-三重自检与边界", "scripts"))
    import _ledger  # noqa: E402
except ImportError:  # pragma: no cover
    _ledger = None

__version__ = "1.2.0"

FILE_RE = re.compile(r"^(?P<letter>[A-Za-z])_[^/\\]+\.md$")
INLINE_RE = re.compile(r"【缺项-([A-Za-z]?)(\d{1,3})：([^】]*)】")
ROW_RE = re.compile(r"^\|\s*缺项-([A-Za-z]?)(\d{1,3})\s*\|(.*)\|\s*$")
H2_RE = re.compile(r"^##\s+(.*\S)\s*$")
CITE_RE = re.compile(r"\[(S-\d{2,})\]")
# 标注源的依据块（〔锚点 …｜材料：…｜推导：…〕）：只给三件套用，不进交付正文
ANNOT_RE = re.compile(r"^〔.+〕\s*$")
# 旧式来源行（（数据来源：…））：成稿口径正文 0 处，出处只进《依据对照表》
SRC_ROW_RE = re.compile(r"^[（(]数据来源\s*[:：]")


def norm_id(letter: str, num: str, fallback: str) -> str:
    return f"{(letter or fallback).upper()}{int(num):02d}"


def find_ledger_near(out_path: str) -> str:
    """从 --out 目录向上找工作区里的检索台账（不猜路径，只认两种固定布局）。"""
    cur = os.path.dirname(os.path.abspath(out_path))
    for _ in range(4):
        for rel in ("过程文件/00-总控/网络补料/检索台账.json", "05_过程/网络补料/检索台账.json"):
            cand = os.path.join(cur, *rel.split("/"))
            if os.path.isfile(cand):
                return cand
        cur = os.path.dirname(cur)
    return ""


def build_citation_block(parts: list[str], letters: list[str], heading: str,
                         ledger_path: str) -> str:
    """生成《外部数据引用清单》（附件六）：只列正文实际引用到的条目。

    正文里 `[S-012]` = 引用台账 S-012 指向的落盘文件；这里把编号、机构、
    标题、URL、检索日期、所在章节摊开，让专家随手能核。未被引用的条目不列，
    免得附件变成"看起来查了很多"的装饰。
    """
    if _ledger is None or not ledger_path:
        return ""
    pd = os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.dirname(os.path.abspath(ledger_path)))))
    led = _ledger.load(pd, ledger_path)
    order: list[str] = []
    part_of: dict[str, int] = {}
    for i, part in enumerate(parts):
        for m in CITE_RE.finditer(part):
            sid = m.group(1)
            if sid not in part_of:
                part_of[sid] = i
                order.append(sid)
    if not order:
        return ""

    def chapter_of(i: int) -> str:
        for line in parts[i].splitlines():
            if line.startswith("# "):
                return line[2:].strip()
        return f"{letters[i]} 章"

    rows, not_found, unresolved = [], [], []
    for sid in order:
        entry = led["valid"].get(sid)
        if not entry:
            unresolved.append(sid)
            continue
        where = chapter_of(part_of[sid])
        if entry.get("notFound"):
            not_found.append(
                f"| [{sid}] | {entry.get('metric','')} | {entry.get('quote','')} | "
                f"{entry.get('date','')} | {entry.get('reason','')} | {where} |")
            continue
        value = f"{entry.get('value','')} {entry.get('unit','')}".strip()
        rows.append(
            f"| [{sid}] | {entry.get('metric','')} | {value} | "
            f"{entry.get('publisher','')}·{entry.get('title','')} | "
            f"{entry.get('url','')} | {entry.get('date','')} | {where} |")

    lines = [f"## {heading}", "",
             "本节逐条列出正文中以 [S-xxx] 标注引用的外部资料，编号与《检索台账》一致，"
             "正文引用编号即引用台账 `file` 字段指向的过程文件；专家可按 URL 与检索日期复核。", ""]
    if rows:
        lines += ["| 编号 | 指标 | 数值或区间 | 来源机构·标题 | URL | 检索日期 | 所在章节 |",
                  "|---|---|---|---|---|---|---|"] + rows + [""]
    if not_found:
        lines += ["**已查无果（确实检索过，无公开数据，不作为依据）**", "",
                  "| 编号 | 指标 | 检索口径 | 检索日期 | 原因 | 所在章节 |",
                  "|---|---|---|---|---|---|"] + not_found + [""]
    if unresolved:
        lines += [f"> 以下编号未能在台账中反查到有效条目，正文不应引用：{'、'.join(unresolved)}", ""]
    return "\n".join(lines)


CHAPTER_LETTERS = "ABCDEF"


def contract_level1_units() -> list[str]:
    """读章节契约的 H1 节点顺序，返回其**成文章节字母**（A、B、C…）。

    为什么不用契约里的 writing_unit：
    契约的 writing_unit 与模板章节并非一一对应——实测契约里 t01～t08 八个节点
    （跨「项目概述」「片区基础概况分析」「现状详细分析」三章）都标 writing_unit: A，
    而它们分属成稿的第 1、2、3 章。check_contracts 的 CHK-GAP-1 明确按
    「H1 依次对应 A、B、C、D、E、F」校验缺项字母（见 _contract_engine.py
    node_chapter_letters 的注释）。故这里必须按 **H1 顺序**编号，
    否则 08 步重映射会把 B 章的缺项编成 A01，09 步必然判「字母与所属章节不符」。

    资源定位走 _skill_paths：优先本 skill 自带 `规则/章节契约.yaml`，
    再回退兄弟 cardo-status-spec 的副本；解析失败返回空表（不阻断组装）。
    """
    spec = _skill_paths.resolve_rule("章节契约.yaml")
    if not spec:
        return []
    units, node = [], None
    try:
        text = open(spec, encoding="utf-8-sig").read()
    except OSError:
        return []
    for line in text.splitlines():
        if re.match(r"\s*-\s+id: *t\d+ *$", line):
            if node and node.get("level") == 1:
                units.append(node["unit"])
            node = {"unit": ""}
            continue
        m = re.match(r"\s+level: *(\d+) *$", line)
        if m and node is not None:
            node["level"] = int(m.group(1))
            continue
    if node and node.get("level") == 1:
        units.append(node["unit"])
    # H1 顺序 → 成文章节字母
    return [CHAPTER_LETTERS[i] if i < len(CHAPTER_LETTERS) else ""
            for i in range(len(units))]


def remap_to_contract_units(parts: list[str], letters: list[str], gap_rows: dict):
    """把分章字母缺项（A–F）按**契约 H1 顺序**重新归一编号；仅对 6 章且各章字母自洽时生效。

    归一后的字母 = 该节点所属的**成文章节**字母（与 09 步 CHK-GAP-1 的口径一致），
    不是契约里的 writing_unit——两者在实测中并不一致（见 contract_level1_units 的说明）。
    """
    units = contract_level1_units()
    if len(parts) != 6 or len(units) != 6:
        return parts, gap_rows
    for i, part in enumerate(parts):
        found = {m.group(1).upper() for m in re.finditer(r"缺项-([A-Za-z])\d{2}", part)}
        if found - {letters[i].upper()}:
            return parts, gap_rows
    mapping, counters = {}, {}

    def repl(unit):
        def _r(m):
            old = (m.group(1) + m.group(2)).upper()
            if old not in mapping:
                counters[unit] = counters.get(unit, 0) + 1
                mapping[old] = f"{unit}{counters[unit]:02d}"
            return "缺项-" + mapping[old]
        return _r

    new_parts = [re.sub(r"缺项-([A-Za-z])(\d{2})", repl(units[i]), p) for i, p in enumerate(parts)]
    new_rows = {}
    for gid, (_, content, source, judged) in gap_rows.items():
        ngid = mapping.get(gid, gid)
        new_rows[ngid] = (ngid, content, source, judged)
    return new_parts, new_rows


def split_process_sections(text: str, gap_kw: str, map_kw: str):
    """剥离过程段，返回 (正文行, 缺项表行, 来源映射行)。"""
    main, gaps, mapping = [], [], []
    state = "main"
    for ln in text.splitlines():
        m = H2_RE.match(ln)
        if m:
            title = m.group(1)
            if gap_kw and gap_kw in title:
                state = "gaps"
                continue
            if map_kw and map_kw in title:
                state = "mapping"
                continue
            if state != "main":
                state = "main"
        if ln.strip() in ("---", "***", "___") and state == "main":
            continue
        if state == "main":
            if ANNOT_RE.match(ln.strip()) or SRC_ROW_RE.match(ln.strip()):
                continue          # 依据块/来源行只留在标注源，正文及其渲染件不背标注
            main.append(ln)
        elif state == "gaps":
            if ln.lstrip().startswith("|"):
                gaps.append(ln)
        else:
            mapping.append(ln)
    return main, gaps, mapping


def main() -> int:
    ap = argparse.ArgumentParser(description="组装分章 md 为可渲染正文")
    ap.add_argument("--extract-dir", required=True, help="分章 md 所在目录")
    ap.add_argument("--out", required=True, help="输出的正文 md")
    ap.add_argument("--mapping-out", default="", help="来源映射留档输出路径（可选）")
    ap.add_argument("--gap-section", default="本节缺项表", help="分章内「缺项表」二级标题关键词")
    ap.add_argument("--mapping-section", default="来源映射", help="分章内「来源映射」二级标题关键词")
    ap.add_argument("--gap-heading", default="数据缺口清单", help="正文末尾汇总表的一级标题文本（不含 #）")
    ap.add_argument("--gap-intro", default="", help="汇总表前的说明段（留空用默认句）")
    ap.add_argument("--append-to-chapter", default="", help="汇总表追加到哪个章字母（默认最后一章）")
    ap.add_argument("--gap-out", default="", help="把汇总的数据缺口清单另存到此文件（不进正文，干净版用）")
    ap.add_argument("--strip-inline-gaps", action="store_true",
                    help="干净版：剥离正文里的内联【缺项-XNN：…】标记")
    ap.add_argument("--gap-index-only", action="store_true",
                    help="缺口版：附件四只出「缺口索引」一页（指向 数据缺口清单.xlsx），不生成完整表格")
    ap.add_argument("--gap-json", default="",
                    help="06 落盘的缺口登记 json（正文不写缺项表时从这里汇总，干净版用）")
    ap.add_argument("--ledger", default="", help="检索台账 json（缺省从 --out 向上找）")
    ap.add_argument("--citation-heading", default="附件六 外部数据引用清单",
                    help="外部数据引用清单的二级标题（默认「附件六 外部数据引用清单」）")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args()

    if not os.path.isdir(args.extract_dir):
        print(f"[error] 分章目录不存在: {args.extract_dir}", file=sys.stderr)
        return 2

    by_letter: dict[str, str] = {}
    for name in sorted(os.listdir(args.extract_dir)):
        m = FILE_RE.match(name)
        if not m:
            continue
        letter = m.group("letter").upper()
        by_letter.setdefault(letter, os.path.join(args.extract_dir, name))
    if not by_letter:
        print(f"[error] {args.extract_dir} 下没有 <章字母>_<章名>.md 形式的分章文件", file=sys.stderr)
        return 2

    letters = sorted(by_letter)
    body_parts: list[str] = []
    gap_rows: dict[str, tuple] = {}
    mapping_parts: list[str] = []

    for letter in letters:
        raw = open(by_letter[letter], encoding="utf-8-sig").read()
        main_lines, gap_lines, map_lines = split_process_sections(
            raw, args.gap_section, args.mapping_section)
        body_parts.append("\n".join(main_lines).strip("\n"))

        for ln in gap_lines:
            m = ROW_RE.match(ln)
            if not m:
                continue
            cells = [c.strip() for c in m.group(3).split("|")]
            if not any(cells):
                continue
            while len(cells) < 3:
                cells.append("")
            gid = norm_id(m.group(1), m.group(2), letter)
            gap_rows.setdefault(gid, (gid, cells[0], cells[1], cells[2] if len(cells) > 2 else ""))

        if map_lines:
            if mapping_parts:
                mapping_parts.append("")
            mapping_parts.append(f"## [{letter}] {os.path.basename(by_letter[letter])}")
            mapping_parts.extend([ln for ln in map_lines if ln.strip()])

    # 缺口登记（06 落盘）：正文不含缺项表时，缺口仍要独立汇总成件
    if args.gap_json and os.path.isfile(args.gap_json):
        try:
            with open(args.gap_json, encoding="utf-8-sig") as f:
                registry = json.load(f)
        except (OSError, json.JSONDecodeError):
            registry = {}
        for row in (registry.get("gaps") or []):
            raw_id = str(row.get("id") or "").strip()
            match = re.match(r"(?:缺项-)?([A-Za-z]+)(\d+)", raw_id)
            if not match:
                continue
            gid = norm_id(match.group(1), match.group(2), str(row.get("chapter") or letters[-1]))
            gap_rows.setdefault(gid, (gid, str(row.get("内容") or ""),
                                      str(row.get("来源") or ""),
                                      str(row.get("现有替代性判断") or "")))

    # 内联编号补前缀（在已剥离过程段的正文里做）
    fixed_parts = []
    for letter, part in zip(letters, body_parts):
        def repl(m, _l=letter):
            return f"【缺项-{norm_id(m.group(1), m.group(2), _l)}：{m.group(3)}】"
        fixed_parts.append(INLINE_RE.sub(repl, part))

    # 干净版：正文不保留任何缺口标记（缺口只进单独的数据缺口清单）
    if args.strip_inline_gaps:
        fixed_parts = [INLINE_RE.sub("", p) for p in fixed_parts]

    # 汇总缺项表（先按契约写作单元归一字母，保证成品与章节契约一致）
    fixed_parts, gap_rows = remap_to_contract_units(fixed_parts, [x.upper() for x in letters], gap_rows)
    ordered = sorted(gap_rows.values(), key=lambda r: (r[0][0], int(re.sub(r"\D", "", r[0]))))
    if ordered:
        if args.gap_index_only:
            # 缺口版 v2：附件四只保留「缺口索引」一页，指向同目录的《数据缺口清单.xlsx》。
            # 不硬造完整缺口表：正文不做内联留痕时，造一张空表反而误导读者。
            gap_md = [f"## {args.gap_heading}", "",
                      f"本版为缺口留痕版：正文各节均按现有材料给出初步判定（推断依据与置信度随文标注），"
                      f"共登记缺口 {len(ordered)} 条，明细见与本报告同目录的《数据缺口清单.xlsx》——"
                      f"按「编号 / 缺口内容 / 建议来源 / 现有替代性判断」逐条列示，"
                      f"并注明补齐后能强化哪一章；补齐后重跑即可收敛为干净版。",
                      ""]
            gap_block = "\n".join(gap_md)
        else:
            intro = args.gap_intro or (
                "下表汇总本报告正文中标注的全部数据缺口。缺口不影响既有结论的定性成立，"
                "但建议在方案设计阶段前完成补充，否则相应的定量分析无法闭环。")
            gap_md = [f"## {args.gap_heading}", "", intro, "",
                      "| 编号 | 缺项内容 | 建议来源 | 现有替代性判断 |", "|---|---|---|---|"]
            for gid, content, source, judged in ordered:
                gap_md.append("| 缺项-{g} | {c} | {s} | {j} |".format(
                    g=gid, c=content.replace("|", "／"), s=source.replace("|", "／"),
                    j=judged.replace("|", "／")))
            gap_block = "\n".join(gap_md)
    else:
        gap_block = ""

    target = (args.append_to_chapter or letters[-1]).upper()
    if target not in letters:
        target = letters[-1]
    if gap_block and args.gap_out:
        # 干净版：缺口清单单独落文件（供导出 数据缺口清单.xlsx），不进正文
        os.makedirs(os.path.dirname(os.path.abspath(args.gap_out)), exist_ok=True)
        with open(args.gap_out, "w", encoding="utf-8", newline="\n") as f:
            f.write(gap_block + "\n")
        gap_block = ""
    if gap_block:
        idx = letters.index(target)
        fixed_parts[idx] = fixed_parts[idx].rstrip("\n") + "\n\n" + gap_block + "\n"

    # 附件六：外部数据引用清单（按正文实际引用生成；无引用则不生成）
    ledger_path = args.ledger or find_ledger_near(args.out)
    cite_block = build_citation_block(fixed_parts, [x.upper() for x in letters],
                                      args.citation_heading, ledger_path)
    if cite_block:
        idx = letters.index(target)
        fixed_parts[idx] = fixed_parts[idx].rstrip("\n") + "\n\n" + cite_block + "\n"

    out_text = "\n\n".join(p.strip("\n") for p in fixed_parts) + "\n"
    with open(args.out, "w", encoding="utf-8", newline="\n") as f:
        f.write(out_text)

    if args.mapping_out and mapping_parts:
        with open(args.mapping_out, "w", encoding="utf-8", newline="\n") as f:
            f.write("# 来源映射（过程留档，不进交付正文）\n\n" + "\n".join(mapping_parts) + "\n")

    h1 = [ln for ln in out_text.splitlines() if ln.startswith("# ")]
    h2 = [ln for ln in out_text.splitlines() if ln.startswith("## ")]
    chars = len(re.sub(r"\s", "", out_text))
    print(f"[assemble] chapters={len(letters)} ({','.join(letters)}) H1={len(h1)} H2={len(h2)} "
          f"非空字符={chars} 缺项={len(ordered)}")
    for ln in h1:
        print("  " + ln)
    return 0


if __name__ == "__main__":
    sys.exit(main())