#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_contract_engine.py -- 契约校验引擎：报告结构映射 + 各断言实现（不解析命令行/不找路径）。"""

from __future__ import annotations

import bisect
import json
import os
import re
import sys

# 运行时是 python313._pth 模式，脚本目录不进 sys.path：兄弟模块必须显式加进来
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import _contract_lib as lib  # noqa: E402
import _ledger as ledger_lib  # noqa: E402
import _num_match as numlib  # noqa: E402

E_OK, E_FAIL, E_INPUT, E_ASK, E_DOCX = 0, 1, 2, 3, 4


class Check:
    def __init__(self, cid: str):
        self.id, self.status, self.items = cid, "pass", []

    def fail(self, msg: str) -> None:
        self.status = "fail"
        # 失败信息置顶：命令行日志与差异清单都取 items[0]，若把提示项排在前面，
        # 打印出来的「FAIL 原因」会和真实失败项对不上（实测踩过）。
        self.items.insert(0, msg)

    def hint(self, msg: str) -> None:
        self.items.append(msg)
        if self.status != "fail":
            self.status = "hint"

    def skip(self, msg: str) -> None:
        self.status = "skip"
        self.items.append(msg)


def build_mapping(tpl: list[tuple[int, str]], heads: list[dict]) -> tuple[dict, list, dict]:
    matched: dict[int, dict] = {}
    pos = 0
    for i, (lvl, title) in enumerate(tpl):
        for j in range(pos, len(heads)):
            h = heads[j]
            if h["level"] == lvl and lib.norm_title(h["title"]) == lib.norm_title(title):
                matched[i], pos = h, j + 1
                break
    owner, cur, k = [], None, 0
    for h in heads:
        if k < len(tpl) and h["level"] == tpl[k][0] and \
                lib.norm_title(h["title"]) == lib.norm_title(tpl[k][1]):
            cur, k = k, k + 1
        owner.append(cur)
    ranges = {}
    for i, h in matched.items():
        end = min((x["line"] for x in heads if x["line"] > h["line"] and
                   x["level"] <= h["level"]), default=None)
        ranges[i] = (h["line"], end if end is not None else 10 ** 9)
    return matched, owner, ranges


def owner_at(owner: list, heads: list[dict], line_no: int):
    idx = bisect.bisect_right([h["line"] for h in heads], line_no) - 1
    return owner[idx] if 0 <= idx < len(owner) else None


def enclosing_scope(heads: list[dict], line_no: int) -> tuple[int, int] | None:
    idx = bisect.bisect_right([h["line"] for h in heads], line_no) - 1
    if idx < 0:
        return None
    start, lvl = heads[idx]["line"], heads[idx]["level"]
    end = min((h["line"] for h in heads if h["line"] > start and h["level"] <= lvl),
              default=10 ** 9)
    return start, end



def check_contract(ctx: dict) -> Check:
    chk, chapter, tables = Check("CHK-CONTRACT-1"), ctx["chapter"], ctx["table_contract"]
    if chapter.get("schema") != "cardo/chapter-contract/v1":
        chk.fail(f"章节契约 schema 非法: {chapter.get('schema')!r}")
    if tables.get("schema") != "cardo/table-contract/v1":
        chk.fail(f"表格契约 schema 非法: {tables.get('schema')!r}")
    tpl_cfg = chapter.get("template") or {}
    if lib.sha256_file(ctx["template"]) != str(tpl_cfg.get("sha256") or "").lower():
        chk.fail("模板 sha256 与契约不一致（契约漂移）：请重新抽取契约并升 version")
    nodes = chapter.get("nodes") or []
    if len(nodes) != len(ctx["tpl"]):
        chk.fail(f"契约节点数 {len(nodes)} ≠ 模板标题数 {len(ctx['tpl'])}")
    units, dirs, node_ids = set(), set(ctx["directions"]), [str(n.get("id")) for n in nodes]
    for i, node in enumerate(nodes):
        nid = str(node.get("id") or "")
        if nid != f"t{i + 1:02d}":
            chk.fail(f"第 {i + 1} 个节点 id 应为 t{i + 1:02d}，实际 {nid!r}")
        if i < len(ctx["tpl"]) and int(node.get("level") or 0) != ctx["tpl"][i][0]:
            chk.fail(f"{nid} level 与模板第 {i + 1} 个标题层级不一致")
        unit = str(node.get("writing_unit") or "")
        units.add(unit)
        if unit not in set("ABCDE"):
            chk.fail(f"{nid} writing_unit 非法: {unit!r}（必须 A–E）")
        for d in ((node.get("requirements") or {}).get("by_direction") or {}):
            if d not in dirs:
                chk.fail(f"{nid} 的方向键 {d!r} 不在项目库配置 directions 里")
    if not units.issuperset(set("ABCDE")):
        chk.fail(f"写作单元未覆盖 A–E（缺 {sorted(set('ABCDE') - units)}）")
    if not (chapter.get("extra_sections") or {}).get("titles"):
        chk.fail("extra_sections.titles 为空（模板外 H1/H2 白名单未登记）")
    seen = set()
    for t in tables.get("tables") or []:
        tid = str(t.get("id") or "")
        if tid in seen:
            chk.fail(f"表格契约 id 重复: {tid}")
        seen.add(tid)
        if not t.get("columns"):
            chk.fail(f"{tid} columns 为空")
        if str(t.get("anchor_node")) not in node_ids:
            chk.fail(f"{tid} anchor_node {t.get('anchor_node')!r} 在章节契约中不存在")
        if t.get("status") not in ("active", "planned"):
            chk.fail(f"{tid} status 非法: {t.get('status')!r}")
    mat = ctx.get("materials_path") or ""
    h1 = {lib.norm_title(t) for lvl, t in ctx["tpl"] if lvl == 1}
    if mat and os.path.isfile(mat):
        try:
            data = lib.parse_yaml(lib.read_text(mat))
            items = list(data.get("common") or [])
            for extra in (data.get("direction_extra") or {}).values():
                items.extend(extra or [])
            for item in items:
                for ch in item.get("chapters") or []:
                    if lib.norm_title(str(ch)) not in h1:
                        chk.fail(f"必备材料清单 {item.get('id')} 的 chapters 含非模板 H1 名: {ch}")
        except ValueError as exc:
            chk.fail(f"必备材料清单解析失败: {exc}")
    return chk

def _paragraph(lines: list[str], line_no: int) -> str:
    """取 line_no（1 起）所在的自然段（以空行分隔）。"""
    i = max(0, line_no - 1)
    start = i
    while start > 0 and lines[start - 1].strip():
        start -= 1
    end = i
    while end + 1 < len(lines) and lines[end + 1].strip():
        end += 1
    return "\n".join(lines[start:end + 1])


def _is_table_head(lines: list[str], line_no: int) -> bool:
    """line_no（1 起）是否为 markdown 表头行或 |---| 分隔行。"""
    s = lines[line_no - 1].strip()
    if not s.startswith("|"):
        return False
    if set(s) <= set("|-: "):
        return True
    nxt = lines[line_no].strip() if line_no < len(lines) else ""
    return nxt.startswith("|") and set(nxt) <= set("|-: ")


def _cite2_hits(lines: list[str]) -> list[tuple[int, str]]:
    """空头依据：①「据公开资料/公开区间」类措辞；②直接贴 URL 或写检索日期却无编号。

    ② 是 ① 的补漏：只写"网址 …，检索日期 …"、不写编号，读者仍然无法确认这条
    是否真检索过、原文在不在；按同一类处理。
    """
    out = []
    for no, line in enumerate(lines, start=1):
        if _is_table_head(lines, no):
            continue      # 表头/分隔行的「来源与检索日期」是列名，不是空头依据
        if "[S-xxx]" in line or "[S-XXX]" in line:
            continue      # 「须写成 [S-xxx]」这类编号口径说明，本身不是引用
        for phrase in ledger_lib.CITE2_PHRASES:
            if phrase in line:
                out.append((no, phrase))
                break
        else:
            if ledger_lib.URL_RE.search(line) or "检索日期" in line:
                out.append((no, "外部网址/检索日期"))
    return out


def check_cite(ctx: dict) -> Check:
    """CHK-CITE-1..5：外部依据必须能落到检索台账条目，且条目本身有效。

    这条闸针对的是盲评里最致命的一类问题——"据公开资料""按同类项目公开区间"
    这类空头依据：读者核不了，专家一问就穿。规则是：写外部来源就必须带
    `[S-012]`，编号必须能在台账反查到一条**有效**条目（缺 URL/日期/摘录/落盘文件
    的条目等同不存在），附件还要有一页可核的引用清单。
    """
    chk = Check("CHK-CITE-1")
    lines = ctx["lines"]
    refs = [(sid, no) for no, line in enumerate(lines, start=1)
            for sid in ledger_lib.cite_refs_in_line(line)]
    led = ctx.get("ledger") or {}
    if not led.get("path"):
        if refs:
            chk.fail(f"正文引用了 [{refs[0][0]}]，但项目内没有检索台账.json（行 {refs[0][1]}）")
            return chk
        hits2 = _cite2_hits(lines)
        for no, phrase in hits2[:10]:
            chk.fail(f"「{phrase}」未给条目编号（行 {no}）：外部依据必须写成 [S-xxx] 并能在台账反查")
        if not hits2:
            chk.skip("本项目未使用外部检索台账，正文也无「公开区间/据公开资料」类空头依据")
        return chk

    for problem in led.get("problems") or []:
        chk.fail(problem)
    valid = led.get("valid") or {}

    # CITE-1 引用可反查：无效条目等同不存在
    unresolved: dict[str, list[int]] = {}
    for sid, no in refs:
        if sid not in valid:
            unresolved.setdefault(sid, []).append(no)
    for sid, nos in unresolved.items():
        chk.fail(f"[{sid}] 在台账中不存在或无效（行 {','.join(map(str, nos[:5]))}）")

    # CITE-2 空头依据：有"据公开…"没有编号
    for no, phrase in _cite2_hits(lines):
        if not ledger_lib.cite_refs_in_line(lines[no - 1]):
            chk.fail(f"「{phrase}」没有条目编号（行 {no}）：请补 [S-xxx]，或改为明确标注的内部推断")

    # CITE-3 低置信推断若声称外部核实过，台账必须有对应「已查无果」记录
    for no, line in enumerate(lines, start=1):
        if "推断" not in line or "置信度：低" not in line:
            continue
        if not any(w in line for w in ledger_lib.CITE3_EXTERNAL):
            continue
        para = _paragraph(lines, no)
        ok = any(sid in valid and valid[sid].get("notFound")
                 for sid in ledger_lib.cite_refs_in_line(para))
        if not ok:
            chk.fail(f"低置信推断声称外部核实，但台账没有对应「已查无果」记录（行 {no}）")

    # CITE-4 附件引用清单必须覆盖正文全部引用（清单由 08 组装自动生成）
    if refs:
        idx = next((i for i, line in enumerate(lines) if "外部数据引用清单" in line), None)
        if idx is None:
            chk.fail("正文引用了外部条目，但附件缺少《外部数据引用清单》（重跑 08 组装）")
        else:
            listed = {m.group(1) for m in ledger_lib.CITE_RE.finditer("\n".join(lines[idx:]))}
            missing = sorted({sid for sid, _ in refs} - listed)
            if missing:
                chk.fail(f"附件引用清单缺条目：{'、'.join(missing[:10])}")

    # CITE-5 台账有条目但正文未引用 —— 只提示，不阻塞
    unused = sorted(set(valid) - {sid for sid, _ in refs})
    if unused:
        chk.hint(f"台账条目未被正文引用：{'、'.join(unused[:10])}（不阻塞，可留作交叉校核）")
    if not chk.items:
        chk.hint(f"外部引用闭环通过：正文引用 {len({s for s, _ in refs})} 条，台账有效条目 {len(valid)} 条")
    return chk



def check_tpl(ctx: dict) -> tuple[Check, Check]:
    tpl1, tpl2 = Check("CHK-TPL-1"), Check("CHK-TPL-2")
    for i, (lvl, title) in enumerate(ctx["tpl"]):
        if i not in ctx["matched"]:
            tpl1.fail(f"模板标题缺失/错序: [H{lvl}] {title}")
        elif int(ctx["matched"][i]["level"]) != lvl:
            tpl1.fail(f"标题层级不符: {title} 期望 H{lvl} 实际 H{ctx['matched'][i]['level']}")
    tpl_set = {(lvl, lib.norm_title(t)) for lvl, t in ctx["tpl"]}
    allow = [lib.norm_title(str(t)) for t in
             ((ctx["chapter"].get("extra_sections") or {}).get("titles") or [])]
    for h in ctx["heads"]:
        if h["level"] > int((ctx["chapter"].get("template") or {}).get("max_level") or 2):
            continue
        if (h["level"], lib.norm_title(h["title"])) in tpl_set:
            continue
        if lib.norm_title(h["title"]) not in allow:
            tpl2.fail(f"模板外标题未登记 extra_sections: [H{h['level']}] {h['title']}")
    return tpl1, tpl2


def has_annotated_source(ctx: dict) -> bool:
    """本章是否为标注源成文（分章稿里带〔锚点〕依据块）。

    成稿口径：出处由依据块 +《依据对照表》承载，正文不再背来源行；
    依据块的段落覆盖与数值有源由 09 的 check_evidence（CHK-EVID）负责，
    因此这里不再要求「每节恰 1 条（数据来源：…）」。
    """
    pd = str(ctx.get("project_dir") or "")
    d = os.path.join(pd, "过程文件", "06-分章成文", "分章初稿")
    if not pd or not os.path.isdir(d):
        return False
    for name in sorted(os.listdir(d)):
        if not name.endswith(".md"):
            continue
        try:
            if "〔锚点" in lib.read_text(os.path.join(d, name)):
                return True
        except OSError:
            continue
    return False


def check_src(ctx: dict) -> Check:
    """CHK-SRC-1（成稿口径）：正文 0 处来源行，出处由〔锚点〕+《依据对照表》承载。

    等价替换而非删项：来源行退出正文后，改由 09 依据校验的
    「对照表回指 + 逐数链条闭合」核到原文；这里守两道——
    正文不许再出现来源行；没有标注源就不存在可闭合的链，直接判不合格。
    """
    chk = Check("CHK-SRC-1")
    hits = [i + 1 for i, ln in enumerate(ctx["lines"])
            if ln.lstrip().startswith(("（数据来源：", "(数据来源："))]
    if hits:
        chk.fail(f"正文仍有 {len(hits)} 处「（数据来源：…）」来源行（首处第 {hits[0]} 行）："
                 "成稿口径正文 0 处，出处只进《依据对照表》")
        return chk
    if not has_annotated_source(ctx):
        chk.fail("正文无来源行，且 06 分章初稿无〔锚点〕依据块："
                 "追溯链断了（删掉来源行不等于成稿，须按标注源重跑 06）")
        return chk
    chk.skip("正文 0 处来源行；出处由〔锚点〕与《依据对照表》承载（由 09 依据校验逐数核到原文）")
    return chk


_BANNED_LINE_RE = re.compile(r"-\s*\{\s*pattern:\s*([^,}]+?)\s*,\s*scopes:\s*\[([^\]]*)\]\s*\}")


def _banned_ref_phrases(ctx: dict, ref: str, scope: str) -> list[str]:
    """按契约里的 `banned_phrases_ref` 读禁句单一来源（相对 `Skills/` 解析）。

    读不到或该 scope 为空时抛 ValueError —— 由调用方转成 FAIL，
    不允许「找不到禁句表就当没有禁句」。
    """
    base = os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(ctx.get("chapter_path") or ""))))
    path = os.path.normpath(os.path.join(base, str(ref)))
    try:
        with open(path, encoding="utf-8") as fh:
            text = fh.read()
    except OSError as exc:
        raise ValueError(f"禁句表不可读：{ref}（{path}；{type(exc).__name__}）") from exc
    phrases = [m.group(1).strip() for m in _BANNED_LINE_RE.finditer(text)
               if scope in [s.strip() for s in m.group(2).split(",") if s.strip()]]
    if not phrases:
        raise ValueError(f"禁句表无 scope={scope} 条目：{ref}")
    return phrases


def check_ban(ctx: dict) -> Check:
    chk = Check("CHK-BAN-1")
    head_lines = {h["line"] for h in ctx["heads"]}
    phrases = list(ctx["chapter"].get("banned_phrases") or [])
    ref = str(ctx["chapter"].get("banned_phrases_ref") or "").strip()
    if ref:
        try:
            phrases = _banned_ref_phrases(ctx, ref, "gap")
        except ValueError as exc:
            chk.fail(str(exc))
            return chk
    for phrase in phrases:
        hits = [n + 1 for n, ln in enumerate(ctx["lines"])
                if n not in head_lines and str(phrase) in ln]
        if hits:
            chk.fail(f"残留模板指令句「{phrase}」{len(hits)} 处（行 {hits[:5]}）")
    for marker in ctx["chapter"].get("banned_direction_markers") or []:
        hits = [n + 1 for n, ln in enumerate(ctx["lines"]) if str(marker) in ln]
        if hits:
            chk.fail(f"正文出现方向标记「{marker}」{len(hits)} 处（行 {hits[:5]}）")
    return chk


# 成文单元字母 A–F 对应模板的六个章（H1 依次为 项目概述 / 片区基础概况分析 /
# 现状详细分析 / 核心问题汇总 / 现状综合研判与更新思路 / 附件）。
CHAPTER_LETTERS = "ABCDEF"


def node_chapter_letters(nodes: list[dict]) -> dict[str, str]:
    """节点 id → 所属成文单元字母。

    按契约节点的 level 归组：每个 level 1 节点是一个章首（H1），
    其后的 level 2 节点归属该章；H1 依次对应 A、B、C、D、E、F。

    不能用契约里的 writing_unit 判缺项字母：契约的 writing_unit 与
    模板章节并非一一对应（t05「片区基础概况分析」H1 标的是 A，
    但它属第 2 章即 B；build_write_tasks.py 的注释也记了同一处错位）。
    缺项编号按**成文章节**编号才是写作规范的口径。
    """
    out: dict[str, str] = {}
    cur = ""
    idx = 0
    for n in nodes:
        lvl = int(n.get("level") or 0)
        if lvl == 1:
            cur = CHAPTER_LETTERS[idx] if idx < len(CHAPTER_LETTERS) else cur
            idx += 1
        if cur and n.get("id"):
            out[str(n["id"])] = cur
    return out


def check_gap(ctx: dict) -> Check:
    chk = Check("CHK-GAP-1")
    if ctx.get("edition") == "clean":
        # 干净版：正文不带缺项标记、也没有附件四，闭环要求不适用
        # （缺口统一由《数据缺口清单.xlsx》承载）。
        chk.skip("干净版不要求缺项闭环（缺口走数据缺口清单.xlsx）")
        return chk
    nodes, heads, lines = ctx["chapter"]["nodes"], ctx["heads"], ctx["lines"]
    node_letters = node_chapter_letters(nodes)
    inline: dict[str, int] = {}
    # 跨章引用：D 章引用 C 章的缺项编号时，若**显式说明**（如「详见第三章缺项登记」），
    # 就不判「字母与所属章节不符」——那是有意的交叉引用，不是编号错误。
    # 不显式说明的跨章编号仍判 fail：读者无从知道它在说哪一章的缺口。
    cross_ref_re = re.compile(r"见第\s*[一二三四五六七八九十\d]+\s*章|详见|参见|另见")
    for no, ln in enumerate(lines):
        explicit_cross = bool(cross_ref_re.search(ln))
        for m in lib.GAP_INLINE_RE.finditer(ln):
            letter, num = (m.group(1) or "").upper(), int(m.group(2))
            idx = owner_at(ctx["owner"], heads, no)
            nid = str(nodes[idx].get("id")) if idx is not None and idx < len(nodes) else ""
            unit = node_letters.get(nid, "")
            gid = f"{letter or unit}{num:02d}"
            if letter and unit and letter != unit and not explicit_cross:
                chk.fail(f"缺项-{gid} 位于 {nid}（所属章节 {unit}），字母与所属章节不符"
                         f"（确为跨章引用请写「详见第三章缺项登记」）")
            if gid in inline:
                chk.fail(f"正文缺项编号重复: 缺项-{gid}")
            inline.setdefault(gid, no + 1)
    t6 = next((t for t in ctx["table_contract"].get("tables") or []
               if "missing_id_closure" in (t.get("checks") or [])), None)
    rows: set[str] = set()
    insts = ctx["table_matches"].get(t6["id"]) if t6 else None
    if not insts:
        # v2（2026-11-11）：附件四只保留「缺口索引」一页，指向《数据缺口清单.xlsx》，
        # 不再在正文造完整缺口表；行级闭环由 xlsx 承担，正文只校验编号合法性。
        return chk
    if insts:
        for r in insts[0]["rows"]:
            m = re.match(r"^缺项-([A-Za-z])(\d{2})$", (r[0] if r else "").strip())
            if not m:
                chk.fail(f"附件四表格行编号不合法: {(r[0] if r else '')!r}")
                continue
            gid = m.group(1).upper() + m.group(2)
            if gid in rows:
                chk.fail(f"附件四表格编号重复: 缺项-{gid}")
            rows.add(gid)
    missing_rows, orphan_rows = sorted(set(inline) - rows), sorted(rows - set(inline))
    if missing_rows:
        chk.fail(f"正文有缺项标记但附件四表缺行: {'、'.join(missing_rows)}")
    if orphan_rows:
        chk.fail(f"附件四表有行但正文未引用: {'、'.join(orphan_rows)}")
    if inline and not insts:
        chk.fail("正文有缺项标记，但未找到附件四数据缺口清单表格")
    return chk


def check_requirements(ctx: dict) -> tuple[Check, list]:
    chk, manual = Check("CHK-REQ-1"), []
    direction = ctx.get("direction")
    for i, node in enumerate(ctx["chapter"]["nodes"]):
        nid = str(node.get("id"))
        reqs = node.get("requirements") or {}
        items = list(reqs.get("common") or []) + list(reqs.get("manual") or [])
        by = reqs.get("by_direction") or {}
        if direction and direction in by:
            items += list(by[direction] or [])
        a, b = ctx["ranges"].get(i, (0, 0))
        text = "\n".join(ctx["lines"][a:b])
        scope_heads = [h for h in ctx["heads"] if a <= h["line"] < b]
        for req in items:
            rid = str(req.get("id") or "")
            if str(req.get("check")) != "auto":
                manual.append({"node": nid, "id": rid, "note": str(req.get("note") or "")})
                continue
            typ = req.get("type")
            if i not in ctx["ranges"]:
                chk.fail(f"{rid}: 节点在正文中缺失，无法校验")
            elif typ == "keywords_any":
                kws = [str(k) for k in (req.get("keywords") or [])]
                if not any(k in text for k in kws):
                    chk.fail(f"{rid}（{nid}）关键词一个都没命中：{kws}")
            elif typ == "section_exists":
                for title in [str(t) for t in (req.get("titles") or [])]:
                    if not any(lib.norm_title(h["title"]).startswith(lib.norm_title(title))
                               for h in scope_heads):
                        chk.fail(f"{rid}（{nid}）缺少小节: {title}")
            elif typ == "min_chars":
                n = len(re.sub(r"\s", "", text))
                if n < int(req.get("min") or 0):
                    # 缺口版：本节**因缺料**而写不满字数，不是写作缺陷。
                    # 降级为提示——否则缺口版永远过不了（缺料节写不到 300 字）。
                    if ctx.get("gapMode"):
                        chk.hint(f"{rid}（{nid}）篇幅 {n} 字 < {req.get('min')} 字"
                                 f"（缺口版：缺料所致，仅提示）")
                    else:
                        chk.fail(f"{rid}（{nid}）篇幅 {n} 字 < {req.get('min')} 字")
            elif typ == "table_required":
                tid = str(req.get("table") or "")
                tc = ctx["table_map"].get(tid) or {}
                if ctx.get("edition") == "clean" and "missing_id_closure" in (tc.get("checks") or []):
                    # 干净版无附件四：该表要求不适用（缺口走数据缺口清单.xlsx），不判 fail
                    chk.hint(f"{rid}（{nid}）{tid} 在干净版不出（缺口走 数据缺口清单.xlsx）")
                    continue
                if "missing_id_closure" in (tc.get("checks") or []) and not ctx["table_matches"].get(tid) \
                        and not _has_inline_gap(ctx):
                    # 附件四为空是正常的（正文没有内联缺项标记），不判 fail
                    chk.hint(f"{rid}（{nid}）{tid} 为空（正文无内联缺项标记）")
                    continue
                if not ctx["table_matches"].get(tid):
                    if tc.get("status") == "planned":
                        chk.hint(f"{rid}（{nid}）计划表 {tid} 未产出（status=planned，只提示）")
                    else:
                        chk.fail(f"{rid}（{nid}）缺少表格 {tid}")
            else:
                chk.hint(f"{rid}（{nid}）未知要求类型 {typ!r}，未校验")
    return chk, manual


def _facts_hay(path: str) -> tuple[set, dict]:
    """可选事实底座（T22）：读事实库 → 数值对集合 + 缺口摘要。不存在/坏文件时降级为空（保持原行为）。"""
    if not path or not os.path.isfile(path):
        return set(), {}
    try:
        with open(path, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
    except (OSError, ValueError):
        return set(), {"note": "事实库解析失败（已降级为仅转储校验）"}
    if not isinstance(data, dict) or data.get("schema") != "cardo/facts/v1":
        return set(), {"note": "事实库 schema 不符（已降级为仅转储校验）"}
    pairs = set()
    for fact in data.get("facts") or []:
        unit = str((fact or {}).get("unit") or "")
        value = (fact or {}).get("value")
        if not unit or isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        num = f"{value:.6f}".rstrip("0").rstrip(".")
        pairs.add((num, unit))
        if unit.startswith("万") and unit != "万元":
            pairs.add((num, "万"))
        if unit.startswith("亿") and unit != "亿元":
            pairs.add((num, "亿"))
    return pairs, {"facts": len(data.get("facts") or []),
                   "missing": data.get("missing") or [],
                   "conflict": data.get("conflicts") or []}


def check_num(ctx: dict) -> Check:
    chk = Check("CHK-NUM-1")
    if not ctx["source_text"]:
        chk.skip("未找到原始报告转储（--source-dump / 05_过程/原文转储），跳过数字溯源")
        return chk
    hay = set(lib.numeric_pairs(lib.normalize_units(ctx["source_text"])))
    fact_pairs, fact_gaps = _facts_hay(str(ctx.get("facts_path") or ""))
    hay |= fact_pairs
    if fact_pairs:
        chk.hint(f"事实库命中集合已并入：{fact_gaps.get('facts', 0)} 条事实（数字可命中 facts[].value 或转储）")
    # 表格里的数值在转储里往往是**裸数字**（单位在表头或列名里），
    # 如附表「现状基本情况」的主楼层面积列，源里只有 512.45，
    # 正文写「512.45平方米」时 (数值, 单位) 配不上，会被误判为编造。
    # 故另建「裸数字集合」：数值命中即算数（单位由列名/语境补足），
    # 只把这类命中记为 hint——不放宽「数字必须来自源」这条红线。
    # 裸数字集合：源文本中出现的所有数值（不要求带单位）。
    # 表格里的数值在转储里往往只有数字本身（单位在表头或列名里）。
    bare = set(lib.bare_numbers(ctx["source_text"]))
    bare |= {n for n, _u in hay}
    # 数值比较按**数值**而非字符串：源里是 645.8、正文写 645.80 时，
    # 字符串不等会被误判为编造。转 float 后按数值比对（同一数值的不同写法等价）。
    def _numkey(v):
        try:
            return round(float(v), 6)
        except (TypeError, ValueError):
            return v
    hay_num = {_numkey(n) for n, _u in hay}
    bare_num = {_numkey(n) for n in bare}
    # 台账数字来源池（2026-11 增）：**有效**台账条目的值可以进正文，
    # 前提是该值所在行必须带 [S-xxx]。有值无编号照样按「无源数字」判失败——
    # 数字对账闸与引用闸不许互相放行。
    led = ctx.get("ledger") or {}
    cited_by_line: dict[int, list[str]] = {}
    for no, ln in enumerate(ctx["lines"], start=1):
        ids = ledger_lib.cite_refs_in_line(ln)
        if ids:
            cited_by_line[no] = ids

    def _ledger_value_hit(line_no: int, value) -> bool:
        try:
            val = float(value)
        except (TypeError, ValueError):
            return False
        for sid in cited_by_line.get(line_no, []):
            entry = (led.get("valid") or {}).get(sid)
            if not entry:
                continue
            for x in ledger_lib.values_of(entry):
                if abs(val - x) <= max(1e-6, abs(x) * 0.005):
                    return True
        return False

    ledger_hits = 0
    unit_matched = 0
    derived_count = 0
    if fact_gaps.get("note"):
        chk.hint(fact_gaps["note"])
    ctx["facts_gaps"] = fact_gaps
    wl = {str(x) for x in (ctx["chapter"].get("num_whitelist") or [])}
    skip = set()
    for tb in ctx["table_matches"].get("TBL-03", []) + ctx["table_matches"].get("TBL-04", []):
        skip.update(range(tb["line"], tb["end_line"]))
    norm_lines = lib.normalize_units("\n".join(ctx["lines"])).splitlines()
    # 推算标注：正文里写明「测算/加总/计算/折算/相除/推算」的数值，
    # 属据材料算出的派生量，不是编造。带这类标注的行不参与数字溯源，
    # 但仍要求作者显式标注——没标注的推算值照样判 fail。
    # 「估算/反算」也计入派生标注：v2 允许硬指标给「估算值＋反算依据」，
    # 只要正文显式写明估算与反算依据，就不是编造原值（用户确认的三条决策之一）。
    derived_re = re.compile(r"测算|加总|计算|折算|相除|推算|除以|累加|估算|反算")
    # 日期（2026 年 2 月 25 日 / 2 月 25 日）里的日、月不是统计数值，先剥离再溯源，
    # 否则「25日」这类日期片段会被当成无来源的带单位数值（误报）。
    date_re = re.compile(r"\d{2,4}\s*年\s*\d{1,2}\s*月(\s*\d{1,2}\s*日)?|\d{1,2}\s*月\s*\d{1,2}\s*日")
    # 表格的数值说明常写在**表后**（如「占比为按匡算总额测算所得」），
    # 故推算标注按**节**判定：只要该节里出现测算标注，该节内的数字都视为派生量。
    # 节 = 二级标题到下一个二级标题之间的行区间。
    derived_secs: set[int] = set()
    cur_sec = None
    for no, ln in enumerate(norm_lines):
        if ln.startswith("#"):
            cur_sec = None
            if re.match(r"^#{1,3}\s+\S", ln):
                cur_sec = no
        if cur_sec is not None and derived_re.search(ln):
            derived_secs.add(cur_sec)
    sec_bounds: list[tuple[int, int]] = []
    heads_idx = [i for i, ln in enumerate(norm_lines) if re.match(r"^#{1,3}\s+\S", ln)]
    for k, st in enumerate(heads_idx):
        en = heads_idx[k + 1] if k + 1 < len(heads_idx) else len(norm_lines)
        sec_bounds.append((st, en))
    def _sec_has_derived(no: int) -> bool:
        for st, en in sec_bounds:
            if st <= no < en:
                return any(d in derived_secs for d in derived_secs
                           if st <= d < en)
        return False
    misses: dict[tuple[str, str], int] = {}
    for no, ln in enumerate(norm_lines):
        if no in skip:
            continue
        ln_scan = date_re.sub("", ln)
        derived = bool(derived_re.search(ln)) or _sec_has_derived(no)
        for num, unit in lib.numeric_pairs(ln_scan):
            if num in wl or f"{num}{unit}" in wl:
                continue
            if (num, unit) in hay:
                continue
            if _numkey(num) in bare_num:
                # 数值在源里，单位由表头/列名补足 —— 计为命中
                unit_matched += 1
                continue
            # 万级收敛等价：正文写「约 4.73 万㎡」、源里是「47333 ㎡」时属同一数值的
            # 不同量级写法，不是无源数字。只对「万」量级做，容差 0.2%，且必须
            # 真的能在源里找到对应裸数值——不是放过任意近似。
            if unit in ("万", "万平方米", "万㎡", "万元", "万个", "万栋", "万户"):
                try:
                    scaled = float(num) * 10000
                except (TypeError, ValueError):
                    scaled = None
                if scaled is not None and scaled and any(
                        isinstance(b, float) and abs(scaled - b) <= abs(scaled) * 0.002
                        for b in bare_num):
                    unit_matched += 1
                    continue
            if _ledger_value_hit(no + 1, num):
                ledger_hits += 1
                continue
            if derived:
                derived_count += 1
                continue
            misses.setdefault((num, unit), no + 1)
    if misses:
        shown = "、".join(f"{n}{u}（行 {l}）" for (n, u), l in list(misses.items())[:10])
        chk.fail(f"{len(misses)} 个带单位数值在来源转储中未命中: {shown}")
    elif unit_matched or derived_count or ledger_hits:
        bits = []
        if unit_matched:
            bits.append(f"{unit_matched} 处数值的单位在源中由表头/列名补足（数值已命中来源）")
        if ledger_hits:
            bits.append(f"{ledger_hits} 处数值来自检索台账（正文已带 [S-xxx] 编号）")
        if derived_count:
            bits.append(f"{derived_count} 处为正文标注的测算/加总值（已显式标注，非引用原值）")
        chk.hint("；".join(bits))
    # CHK-NUM-2 算术对账。此前「有测算/加总标注」就等于放行，不做任何算术校验，
    # 于是「分项 19820 却写合计 17420」「10.2945 万㎡ 占 47333 ㎡ 的 88.1%（实为 217%）」
    # 「按 13% 计提却写 15% 的金额」三处硬伤在 09 自检里全部绿灯。
    # 现在四类关系写成硬校验，任一不成立即 fail，并给出两处位置（行号）。
    rec = numlib.reconcile("\n".join(ctx["lines"]))
    for issue in rec["issues"]:
        lines_hit = issue.get("lines") or [issue.get("line")]
        chk.fail(f"算术对账不成立（{issue['kind']}，行 {','.join(map(str, lines_hit))}）："
                 f"{issue['message']}")
    return chk


def _has_inline_gap(ctx: dict) -> bool:
    """正文是否存在内联缺项标记（【缺项-XNN：…】）。附件四只在内联留痕时才有内容。"""
    return any(lib.GAP_INLINE_RE.search(ln) for ln in ctx["lines"])


def _annotated_blocks(ctx: dict) -> list[dict]:
    """06 分章初稿逐段切块：[{file, text, anchor}]，anchor 表示该段后紧跟〔锚点〕依据块。"""
    root = os.path.join(ctx.get("project_dir") or "", "过程文件", "06-分章成文", "分章初稿")
    blocks = []
    if not os.path.isdir(root):
        return blocks
    for name in sorted(os.listdir(root)):
        if not name.endswith(".md"):
            continue
        para = []

        def flush(anchored: bool) -> None:
            if para:
                blocks.append({"file": name, "text": "\n".join(para), "anchor": anchored})
                para.clear()

        for raw in lib.read_text(os.path.join(root, name)).splitlines():
            s = raw.strip()
            if s.startswith("〔") and "锚点" in s:
                flush(True)
            elif not s or s.startswith("#"):
                flush(False)
            else:
                para.append(s)
        flush(False)
    return blocks


def _draft_table_block(blocks: list[dict], tb: dict) -> tuple[dict | None, str]:
    """在标注源里找这张表（按表头签名匹配），返回 (所在段, 问题说明)。"""
    def norm(text: str) -> str:
        return re.sub(r"[\s|]", "", str(text))

    sig = "".join(norm(h) for h in tb["header"])
    if not sig:
        return None, "表头为空，无法回指标注源"
    for block in blocks:
        if sig in norm(block["text"]):
            return block, ""
    return None, "在 06 分章初稿里找不到这张表（表头与正文不一致或表非标注源产出）"


def check_tables(ctx: dict) -> list[Check]:
    out = []
    defaults = ctx["table_contract"].get("defaults") or {}
    max_empty = float(defaults.get("empty_cell_ratio") or 0.0)
    annot = has_annotated_source(ctx)
    ctx["_annot_blocks"] = _annotated_blocks(ctx) if annot else []
    for c in ctx["table_contract"].get("tables") or []:
        chk = Check(f"CHK-TBL-{c['id']}")
        if ctx.get("edition") == "clean" and "missing_id_closure" in (c.get("checks") or []):
            # 干净版不设附件四：缺口统一出到 成果文件/数据缺口清单.xlsx，
            # 正文与附件都不留缺口清单（用户确认的「干净版彻底移除附件四」）。
            chk.skip("干净版不出附件四（缺口走 数据缺口清单.xlsx）")
            out.append(chk)
            continue
        insts = ctx["table_matches"].get(c["id"])
        if not insts and "missing_id_closure" in (c.get("checks") or []) and not _has_inline_gap(ctx):
            # 附件四是内联缺项标记的汇总产物：正文没有 【缺项-XNN】 时它自然为空，
            # 不是漏做（缺口版也可能整篇不内联留痕，缺口走数据缺口清单.xlsx）。
            chk.hint("正文无内联缺项标记，附件四为空（缺口未在正文内联留痕）")
            out.append(chk)
            continue
        if not insts:
            if c.get("status") == "planned":
                chk.hint(f"计划中的表未产出：{c['name']}（status=planned，只提示）")
            else:
                chk.fail(f"未找到表「{c['name']}」（表头应为 {c.get('columns')}）")
            out.append(chk)
            continue
        tb = insts[0]
        a = next((i for i, n in enumerate(ctx["chapter"]["nodes"])
                  if str(n.get("id")) == str(c.get("anchor_node"))), None)
        if a is not None and a in ctx["ranges"]:
            lo, hi = ctx["ranges"][a]
            if not (lo <= tb["line"] < hi):
                chk.fail(f"表「{c['name']}」不在 anchor 节点 {c.get('anchor_node')} 的作用域内")
        if c.get("unit_required") and not any(
                re.search(r"万元|亿元|㎡|平方米|平方公里|公里|米|%|％|个|亩|户|人|吨", str(h))
                for h in tb["header"]):
            chk.fail(f"表「{c['name']}」表头没有显式单位")
        total = sum(len(r) for r in tb["rows"])
        empty = sum(1 for r in tb["rows"] for cell in r if not str(cell).strip())
        if total and empty / total > max_empty:
            chk.fail(f"表「{c['name']}」空单元格 {empty}/{total}，超过阈值 {max_empty}")
        scope = enclosing_scope(ctx["heads"], tb["line"])
        if defaults.get("source_cell_required") and scope:
            # 成稿口径：正文不再背「（数据来源：…）」行（08 组装也把锚点剥掉了），
            # 表的出处改由 06 标注源承载。这里不再要来源行，改为回指 06：
            # 「这张表在标注源里有对应表，且表后紧跟〔锚点〕依据块」——断在哪一环就点名哪一环。
            a2, b2 = scope
            seg = ctx["lines"][a2:b2]
            has_src = any(ln.lstrip().startswith(("（数据来源：", "(数据来源：")) for ln in seg)
            if not has_src:
                if annot:
                    block, why = _draft_table_block(ctx.get("_annot_blocks") or [], tb)
                    if why:
                        chk.fail(f"表「{c['name']}」（第 {tb['line']} 行）{why}：表内数据无出处")
                    elif not block["anchor"]:
                        chk.fail(f"表「{c['name']}」（第 {tb['line']} 行）在 "
                                 f"{block['file']} 里后面没有〔锚点〕依据块：表内数据无出处")
                else:
                    # 整章没有锚点（旧格式）：出处断链已由 CHK-SRC-1 判失败，
                    # 这里不重复计同一条根因，避免把 exit 归到表检查上。
                    chk.skip("本章 06 无〔锚点〕（旧格式）：表出处断链由 CHK-SRC-1 统一判失败")
        sum_mode = str(c.get("sum_row") or "off")
        tot_rows = [r for r in tb["rows"] if r and str(r[0]).strip() in ("合计", "总计", "小计")]
        if sum_mode != "off":
            if not tot_rows and sum_mode == "required":
                chk.fail(f"表「{c['name']}」缺合计行（sum_row=required）")
            for r in tot_rows:
                data = [x for x in tb["rows"] if x is not r]
                for ci in range(1, min(len(r), len(tb["header"]))):
                    nums = [float(re.sub(r"[^\d.]", "", str(x[ci])) or 0)
                            for x in data if str(x[ci]).strip()]
                    try:
                        got = float(re.sub(r"[^\d.]", "", str(r[ci])) or 0)
                    except ValueError:
                        continue
                    if nums and abs(sum(nums) - got) > max(0.5, abs(got) * 0.01):
                        chk.fail(f"表「{c['name']}」第 {ci + 1} 列合计 {got} ≠ 明细之和 {sum(nums):.2f}")
        if len(insts) > 1:
            chk.fail(f"表「{c['name']}」在正文出现 {len(insts)} 次：契约未登记重复表，请去重")
        out.append(chk)
    return out

def check_docx(ctx: dict) -> Check:
    chk = Check("CHK-DOCX-1")
    if not ctx["docx"]:
        chk.skip("未找到报告 docx，跳过回读（渲染后重跑即可）")
        return chk
    if ctx.get("docx_error"):
        chk.fail(str(ctx["docx_error"]))
        return chk
    st = ctx.get("docx_struct") or {}
    md_h1 = sum(1 for h in ctx["heads"] if h["level"] == 1)
    md_h2 = sum(1 for h in ctx["heads"] if h["level"] == 2)
    if st["titles"] > 1:
        chk.fail(f"docx Title 段 {st['titles']} 个（应 ≤1）")
    dx_h1 = len([1 for lv, _ in st["headings"] if lv == 1])
    dx_h2 = len([1 for lv, _ in st["headings"] if lv == 2])
    if (dx_h1, dx_h2) != (md_h1, md_h2):
        chk.fail(f"docx 与 md 标题数不一致：docx H1={dx_h1} H2={dx_h2}，md H1={md_h1} H2={md_h2}")
    if st["tables"] != len(ctx["md_tables"]):
        chk.fail(f"docx 表格数 {st['tables']} ≠ md 表格数 {len(ctx['md_tables'])}")
    md_heads = [(h["level"], lib.norm_title(h["title"])) for h in ctx["heads"] if h["level"] <= 2]
    dx_heads = [(lv, lib.norm_title(t)) for lv, t in st["headings"]]
    if md_heads != dx_heads:
        for k, (x, y) in enumerate(zip(md_heads, dx_heads)):
            if x != y:
                chk.fail(f"第 {k + 1} 个标题 docx 与 md 不一致：md={x} docx={y}")
                break
        if len(md_heads) != len(dx_heads):
            chk.fail(f"标题序列长度不一致：md {len(md_heads)} / docx {len(dx_heads)}")
    return chk

def check_clean(ctx: dict) -> Check:
    """CHK-CLEAN-1：干净版正文不得出现缺口口语。

    缺料不是缺陷，但把「资料缺失/待补/缺什么找谁补」写进正文会让报告读起来像
    现状资料缺失报告（真实用户反馈）。干净版要求：缺口只进 数据缺口清单.xlsx。
    """
    chk = Check("CHK-CLEAN-1")
    if ctx.get("edition") != "clean":
        chk.skip("缺口版允许正文留痕")
        return chk
    ref = str(ctx["chapter"].get("banned_phrases_ref") or "").strip()
    if ref:
        try:
            words = _banned_ref_phrases(ctx, ref, "clean")
        except ValueError as exc:
            chk.fail(str(exc))
            return chk
    else:
        words = ["资料缺失", "未提供", "资料不足", "待补", "缺项-", "找谁补", "缺什么"]
    hits: dict[str, int] = {}
    for ln in ctx["lines"]:
        for w in words:
            if w in ln:
                hits[w] = hits.get(w, 0) + 1
    for w, n in hits.items():
        chk.fail(f"干净版正文出现缺口口语「{w}」{n} 处：缺失内容只进 数据缺口清单.xlsx")
    return chk


def run_checks(ctx: dict) -> tuple[list[Check], list, int]:
    """按契约 assertions 的固定顺序跑全部检查，返回 (checks, manual, exit_code)。"""
    checks = [check_contract(ctx)]
    dir_chk = Check("CHK-DIR-1")
    if ctx.get("direction") is None:
        dir_chk.fail("未判定项目方向：请用 --direction，或在 章节映射.yaml / .cardo/project.json 写 direction（闸门 G2）")
    checks.append(dir_chk)
    tpl1, tpl2 = check_tpl(ctx)
    checks.extend([tpl1, tpl2, check_src(ctx), check_ban(ctx), check_num(ctx),
                   check_cite(ctx), check_gap(ctx), check_clean(ctx)])
    req_chk, manual = check_requirements(ctx)
    checks.append(req_chk)
    checks.extend(check_tables(ctx))
    docx_chk = check_docx(ctx)
    checks.append(docx_chk)
    if dir_chk.status == "fail":
        code = E_ASK
    elif docx_chk.status == "fail":
        code = E_DOCX
    elif any(c.status == "fail" for c in checks):
        code = E_FAIL
    else:
        code = E_OK
    return checks, manual, code
