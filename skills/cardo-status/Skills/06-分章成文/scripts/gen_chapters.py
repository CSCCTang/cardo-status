#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""分章成文（取料成文，确定性）：模板章节 ↔ 原始报告转储 → A–F 分章稿。

原则：
1. 只搬运原始报告（转储）里的表述与数字，不新增事实、不补数；
2. 每个成文节点必须带「来源行」（数据来源 + 源报告段范围）；
3. 无来源的节点**不得留空、也不得写「资料缺失」**：按现有材料的间接证据给初步判定，
   标注推断依据与置信度（v2 结论优先），缺口只进 数据缺口清单.xlsx；
4. 表格逐字取自源报告表格转储（不重排口径），仅转成 markdown。

用法：
  python gen_chapters.py --template-outline 模板.outline.json \
      --source-dump 过程文件/03-模板与源报告转储/源报告 \
      --mapping 过程文件/04-章节映射/章节映射.yaml \
      --materials 过程文件/01-建档与材料归集/分类清单/分类清单.json \
      --facts 过程文件/05-事实抽取与事实库/facts/事实库.json \
      --out 过程文件/06-分章成文/分章初稿 [--gap-mode]
"""
import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None

CHAPTERS = ["A", "B", "C", "D", "E", "F"]
# 各章 H1 标题（模板真相，用于识别回填稿是否写完）
CHAPTER_TITLES = {
    "A": "项目概述",
    "B": "片区基础概况分析",
    "C": "现状详细分析",
    "D": "核心问题汇总",
    "E": "现状综合研判与更新思路",
    "F": "附件",
}
# 散材料模式：每节正文字数上限（避免整表/长条文倾泻）
SCATTER_SECTION_CAP = {"A": 900, "B": 900, "C": 1200, "D": 900, "E": 900, "F": 900}
# 散材料模式下各章的缺项建议来源
SCATTER_HINT = {
    "A": "项目资料、可研报告、项目批复",
    "B": "规划文件、片区策划、历史沿革资料",
    "C": "专项调查台账、人口/业态/交通/市政/公服/生态台账",
    "D": "项目资料、体检报告问题清单",
    "E": "基于前述分析与项目资料",
    "F": "项目资料、图件与台账",
}
H1_TO_CHAPTER = {
    "项目概述": "A",
    "片区基础概况分析": "B",
    "现状详细分析": "C",
    "核心问题汇总": "D",
    "现状综合研判与更新思路": "E",
    "科学研判与更新思路": "E",
    "附件": "F",
}
# 模板中无二级标题的章，其节标题由本脚本固定（与各章 chapter_check CONFIG 逐字一致）
D_SECTIONS = ["街巷空间与肌理问题", "建筑安全与风貌破损问题", "文脉保护与活化不足问题",
              "业态老旧与活力不足问题", "交通通行与消防隐患突出问题", "基础设施老旧滞后问题",
              "民生配套与人居品质偏低", "街区治理与运营短板"]
F_SECTIONS = ["附件一 现有基础资料清单", "附件二 图件与影像资料", "附件三 专项台账"]
F_GAP_TITLE = "附件四 数据缺口清单"
GAP_RE = re.compile(r"【(缺项-[A-F]\d{1,3})：(.+?)｜建议来源：(.+?)】")
SOURCE_RE = re.compile(r"^\s*（数据来源：.*）\s*$")

# 摘句阶段就要拦掉的「作者待办语」——它们是源报告写给自己的工作提示，
# 不是对现状的陈述，搬进正文既不成立又会命中 CHK-BAN-1。
# 清单只在 06-分章成文/规则/禁句表.yaml 维护（本文件不再复制）。
from _banned import TODO_PHRASES, EMPTY_PHRASES  # noqa: E402
# 表格搬运的行数上限：源报告的改造清单表动辄数百行，整表搬进现状章会淹没正文。
# 超限时截断并留一行说明，保证「不丢来源、不改数字」的同时不倾泻。
TABLE_MAX_ROWS = 12


def load_yaml(path):
    """读 yaml。优先 PyYAML；没有则用内置 _yaml_lite 子集解析器。

    便携环境里 PyYAML 可能没装（实测缺 yaml）。此前直接 SystemExit，
    06 步整个跑不起来——而契约/规则 yaml 都是受控的简单结构，
    内置子集解析器足够。
    """
    if yaml:
        with open(path, encoding="utf-8") as f:
            return yaml.safe_load(f)
    try:
        import _yaml_lite
        with open(path, encoding="utf-8-sig") as f:
            return _yaml_lite.parse(f.read())
    except Exception as exc:
        raise SystemExit(f"缺少 pyyaml 且内置解析器失败（{exc}）："
                         f"pip install pyyaml，或把 pyyaml 拷进 runtime 的 site-packages")


def read_blocks(path):
    with open(path, encoding="utf-8-sig") as f:
        return [ln.rstrip("\n") for ln in f]


def read_tables(path):
    """tables.txt -> [ {index, rows:[[...]]} ]"""
    tables, cur = [], None
    if not os.path.exists(path):
        return tables
    with open(path, encoding="utf-8-sig") as f:
        for ln in f:
            m = re.match(r"^=====\s*TABLE\s+(\d+)\s*\((\d+)x(\d+)\)", ln.strip())
            if m:
                cur = {"index": int(m.group(1)), "rows": []}
                tables.append(cur)
                continue
            if cur is not None:
                m = re.match(r"^\s+R\d+:\s*(.*)$", ln.rstrip("\n"))
                if m:
                    cells = [c.strip() for c in m.group(1).split(" | ")]
                    cur["rows"].append(cells)
    return tables


def find_one(path, suffix):
    if os.path.isdir(path):
        for fn in sorted(os.listdir(path)):
            if fn.endswith(suffix):
                return os.path.join(path, fn)
        return None
    return path


def read_blocks_source(dump_dir):
    """读转储里的 *blocks.txt（数据化成文仍以原始材料行取数）。"""
    path = find_one(dump_dir, ".blocks.txt")
    return read_blocks(path) if path else []


# 数据化成文：包内依据可用时由 main() 注入 DataWriter
BASELINE_WRITER = None


def norm_title(t):
    t = re.sub(r"\s+", "", str(t or ""))
    return re.sub(r"[：:、，,。.（）()《》“”\"'—\-]", "", t)


def similarity(a, b):
    a, b = norm_title(a), norm_title(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    sa, sb = set(a), set(b)
    return len(sa & sb) / len(sa | sb)


def md_table(rows, max_rows=None):
    """渲染 markdown 表格。

    max_rows 限数据行数（不含表头）：源报告的改造清单表动辄数百行，
    整表搬进现状章会淹没正文。超限时截断并留一行说明「共 N 行，此处摘录前 M 行」，
    既不丢来源、不改数字，也不倾泻。
    """
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        return ""
    total = len(rows) - 1
    dropped = 0
    if max_rows and total > max_rows:
        dropped = total - max_rows
        rows = rows[:max_rows + 1]
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    def esc(c):
        t = str(c).replace("|", "／").replace("\n", " ").strip()
        # 表格逐字搬运，但「作者待办语」格不能进正文（会命中 CHK-BAN-1）。
        # 只替换整格就是待办语的情况；含实质内容的格子保留原样，不改数字。
        if is_todo_phrase(t):
            return "（原文为作者待办事项，见源报告）"
        return t
    out = ["| " + " | ".join(esc(c) for c in rows[0]) + " |",
           "| " + " | ".join("---" for _ in rows[0]) + " |"]
    for r in rows[1:]:
        out.append("| " + " | ".join(esc(c) for c in r) + " |")
    if dropped:
        out.append("| " + " | ".join(["（本表共 %d 行，此处摘录前 %d 行，其余 %d 行见源报告）"
                                     % (total, max_rows, dropped)]
                                    + [""] * (width - 1)) + " |")
    return "\n".join(out)


# 政策原文的识别特征：带文号（〔2020〕23号 / [2020]23号）、书名号文件名、
# 或成文年份。带这些的句子是**可核验的政策引用**，不是空洞套话，
# 其中的「大力推进」「进一步加强」属于原文表述，不应按禁句拦掉。
_CITE_RE = re.compile(
    r"[〔\[]\s*\d{4}\s*\]\s*\d+\s*号"      # 文号：〔2024〕17号
    r"|《[^》]{2,}》"                            # 文件名（"《通知》"也算）
    r"|\d{4}\s*年\d{1,2}\s*月"                 # 成文时间
    r"|国办发|川办发|发改|住建"
    r"|《?[^》\s]{0,6}?(通知|办法|意见|方案|导则|标准|规定|批复|纪要)》?"
)


def is_citable(text):
    """判断一句话是否为可核验的政策/文件引用。"""
    return bool(_CITE_RE.search(text or ""))


def is_todo_phrase(text):
    """判断一句话是否属于「作者待办语」。

    源报告里大量这样的句子：「需明确现状化粪池进水及出水管线分布，方可确认是否能进行改造」——
    这是作者写给自己的工作提示，不是对现状的陈述。搬进现状分析正文既不成立，
    又会命中 CHK-BAN-1（禁句）。摘句阶段就应当拦掉，而不是等自检失败。

    例外：带文号/文件名/成文时间的**政策原文引用**不拦 —— 那里的
    「大力推进」「进一步加强」是原文表述，属可核验依据，不是空洞套话。
    """
    if is_citable(text):
        return False
    return any(p in text for p in TODO_PHRASES)


def is_empty_phrase(text):
    """判断一句话是否是无出处的空洞套话。

    「高度重视」「进一步加强」「大力推进」单独出现时没有任何信息量，
    但它们在**政策原文引用**里是原文表述（可核验），故带出处的句子予以保留。
    """
    if is_citable(text):
        return False
    return any(p in text for p in EMPTY_PHRASES)


def filter_todo_sentences(text):
    """从摘来的正文里剔除作者待办语与空洞套话（按句切分，保留其余原句）。"""
    if not text:
        return text
    if not is_todo_phrase(text) and not is_empty_phrase(text):
        return text
    keep = []
    for piece in re.split(r"(?<=[。；！？\n])", text):
        if not piece:
            continue
        # 段落内有出处（文号/文件名/成文时间）→ 整段按政策原文保留，
        # 否则末句「大力推进……」会被单独拎出来当成空洞套话。
        if is_citable(text):
            keep.append(piece)
            continue
        if not is_todo_phrase(piece) and not is_empty_phrase(piece):
            keep.append(piece)
    return "".join(keep).strip()


class Source:
    def __init__(self, dump_dir):
        outline_path = find_one(dump_dir, ".outline.json")
        blocks_path = find_one(dump_dir, ".blocks.txt")
        tables_path = find_one(dump_dir, ".tables.txt")
        self.name = os.path.splitext(os.path.basename(blocks_path or outline_path))[0]
        self.outline = json.load(open(outline_path, encoding="utf-8"))
        self.blocks = read_blocks(blocks_path)
        self.tables = read_tables(tables_path)
        heads = [x for x in self.outline if x.get("kind") == "heading" and x.get("level", 9) <= 3]
        heads.sort(key=lambda x: x["i"])
        self.headings = heads
        # 同名标题在真实项目文档里很常见（「项目概况」既是 H2 又可能作为 H3 出现）。
        # 用字典 self.range[标题] = 区间 会被**后出现的同名标题覆盖**，
        # 于是「映射选中了概述下的项目概况，实际取到的却是文档末尾另一个同名章节的正文」——
        # 摘句搬运会产出内容与节标题完全错配的正文。
        # 现在按「层级优先 + 位置优先」建索引：同一标题取层级最浅（数字最小）的那个；
        # 同层级取文档中第一个。全部出现位置另存 self.range_all 供查。
        self.range = {}
        self.range_all = {}
        for idx, h in enumerate(heads):
            end = len(self.blocks)
            # 一级标题只取到第一个二级子标题前（引言），避免把整章正文重复搬运；
            # 二、三级标题取到同级或更高级标题前。
            limit = 2 if h["level"] <= 2 else 3
            for nxt in heads[idx + 1:]:
                if nxt["level"] <= limit:
                    end = nxt["i"]
                    break
            self.range_all.setdefault(h["text"], []).append((h["level"], (h["i"], end)))
            cur = self.range.get(h["text"])
            if cur is None:
                self.range[h["text"]] = (h["level"], (h["i"], end))
            else:
                # 已有同名标题：新出现者层级更浅才替换；同层级不替换（取先出现的）
                if h["level"] < cur[0]:
                    self.range[h["text"]] = (h["level"], (h["i"], end))
        # 表块位置 → 表格序号
        self.table_at = {}
        k = 0
        for i, ln in enumerate(self.blocks):
            if ln.startswith("[TABLE"):
                if k < len(self.tables):
                    self.table_at[i] = k
                k += 1

    def span(self, title):
        """取某标题对应的块区间 (start, end)；self.range 存的是 (level, (start, end))。"""
        item = self.range.get(title)
        if not item:
            return None
        return item[1] if isinstance(item[0], int) else item

    def block_texts(self, title, include_tables=True):
        rng = self.span(title)
        if not rng:
            return [], None
        start, end = rng
        out = []
        for i in range(start + 1, end):
            ln = self.blocks[i]
            if not ln.strip():
                continue
            if ln.startswith("[TABLE"):
                if include_tables and i in self.table_at:
                    out.append(("table", i))
                continue
            if ln.startswith("["):
                m = re.match(r"^\[[^\]]*\]\s*(.*)$", ln)
                txt = (m.group(1) if m else ln).strip()
                if txt:
                    out.append(("text", txt))
            else:
                out.append(("text", ln.strip()))
        return out, (start, end - 1)


def first_paragraph(texts):
    for kind, v in texts:
        if kind == "text":
            return v
    return ""


def clipped_join(texts, limit):
    parts, total = [], 0
    for kind, v in texts:
        if kind != "text":
            continue
        if total and total + len(v) > limit:
            break
        parts.append(v)
        total += len(v)
    return "\n\n".join(parts)


def chapter_of(title, level):
    if level == 1:
        return H1_TO_CHAPTER.get(title)
    return None


def anchor_line(prefix, idx, fields, confidence="高"):
    """一条标注源依据块：〔锚点 前缀-P01｜材料：…｜置信度：高〕。"""
    return f"〔锚点 {prefix}-P{int(idx):02d}｜{'｜'.join(fields)}｜置信度：{confidence}〕"


def _contract_keywords():
    """从 09 的章节契约读各节点的 keywords_any 要求（成文与校验同一套口径）。"""
    try:
        base = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        path = os.path.join(base, "09-三重自检与边界", "规则", "章节契约.yaml")
        data = load_yaml(path) or {}
    except Exception:
        return {}
    out = {}
    for n in data.get("nodes") or []:
        reqs = n.get("requirements") or {}
        kws = []
        for req in list(reqs.get("common") or []) + list(reqs.get("manual") or []):
            if str(req.get("type")) == "keywords_any":
                kws.extend(str(k) for k in (req.get("keywords") or []))
        for items in (reqs.get("by_direction") or {}).values():
            for req in (items or []):
                if str(req.get("type")) == "keywords_any":
                    kws.extend(str(k) for k in (req.get("keywords") or []))
        if kws:
            out[str(n.get("id"))] = list(dict.fromkeys(kws))
    return out


def build_scatter_chapter(letter, nodes, h2_nodes, chapter_title, mapping_by_id,
                          facts, spec, lines, gaps, gap_slug, counter,
                          materials=None):
    """散材料模式成文：按事实库组织段落，无料节走缺项。

    与常规模式的区别：
      * 不逐段摘抄源报告，而是用事实库（带来源引文与单位数值）按来源文件成段；
      * 事实全部被判为噪声（目录行/规范条文）时，该节记缺项，不硬凑字数；
      * 成稿口径：正文不背「（数据来源：…）」行，出处逐段写进〔锚点〕依据块，
        由 08 组装时剥离、09 导出《依据对照表》承载。
    """
    from scatter_writer import (collect_section_facts, write_scatter_section,
                                build_metric_table, area_map)

    facts_by_node = collect_section_facts(nodes, mapping_by_id, facts)
    all_facts = list(facts.get("facts") or [])
    _rows, table_sources = build_metric_table(all_facts)
    kw_by_node = _contract_keywords()
    mat_names = [os.path.basename(str(f.get("rel") or "")) for f in
                 ((materials or {}).get("files") or []) if f.get("rel")]
    h1 = next((n for n in nodes if n["level"] == 1), {})
    h1_kw = kw_by_node.get(str(h1.get("id")), [])

    # 章导语：说明本章依据什么成文（其后紧跟一条判断型依据块）
    lines.append("")
    kw_intro = "、".join(h1_kw[:4]) if h1_kw else "现状与更新思路"
    body_rule = ("各节按表现、证据、成因、影响、对应更新方向五要素成文，证据项标注"
                 "「可核对」或「推断，需现场核实」。" if letter == "D" else
                 "各节先写现状与初步判定，再写判断依据与待核项。")
    lines.append(f"本章为「{chapter_title}」，按{kw_intro}口径，依据项目材料成文；{body_rule}")
    lines.append("")
    lines.append(anchor_line(f"{letter}-00", 1, ["判断：本章导语，说明本章成文口径与追溯方式"]))

    # 区域构成：本轮确认的区域归属（--area），在 A 章说清，两个区域的数据不合并
    area_names = list(area_map().keys())
    if letter == "A" and len(area_names) >= 2:
        lines.append("")
        lines.append("本项目包含{}，两个区域的现状数据分别归属、分别统计，不合并、不互替。"
                     .format("、".join(area_names)))
        lines.append("")
        lines.append(anchor_line(f"{letter}-00A", 1,
                                 ["判断：区域构成按本轮确认的区域归属口径列示"]))

    cap = SCATTER_SECTION_CAP.get(letter, 900)
    kinds = []
    for seq, title in enumerate(spec, 1):
        node = next((n for n in h2_nodes if norm_title(n["title"]) == norm_title(title)), None)
        node_id = node["id"] if node else ""
        node_facts = facts_by_node.get(node_id, []) if node_id else []
        ctx = {
            "all_facts": all_facts,
            "table_sources": table_sources,
            "materials": mat_names,
            "keywords": kw_by_node.get(node_id, h1_kw),
        }

        lines.append("")
        lines.append(f"## {title}")
        text, section_gaps, kind = write_scatter_section(
            title, node_facts, letter, gap_slug, counter,
            hint=SCATTER_HINT.get(letter, "项目资料、专项调查台账"),
            max_chars=cap, anchor_prefix=f"{letter}-{seq:02d}", ctx=ctx)
        lines.append("")
        lines.append(text)
        gaps.extend(section_gaps)
        kinds.append(kind)

    # D 章：核心问题清单（TBL-02）。由各节判定汇总，不新增事实、不带数值。
    if letter == "D":
        from scatter_writer import _profile
        lines.append("")
        # 模板最多二级标题：这里用加粗段作表题，避免 CHK-HEAD-3。
        # 表题也是正文段落，必须自带依据块（否则 09 判「段落无依据」）。
        lines.append("**核心问题清单**")
        lines.append(anchor_line(f"{letter}-T0", 1,
                                 ["判断：下表由本章各节初步判定汇总，未新增事实"],
                                 confidence="中"))
        lines.append("")
        lines.append("| 序号 | 问题类别 | 问题表征 | 关键事实与指标 |")
        lines.append("| --- | --- | --- | --- |")
        for i, title in enumerate(spec, 1):
            prof = _profile(title)
            cat = re.sub(r"问题$", "", title)
            key = f"需现场核实：{prof['verify']}"
            lines.append(f"| {i} | {cat} | {prof['judge']} | {key} |")
        # 表后必须紧跟依据块（不留空行），否则 09 判「表内数据无出处」
        lines.append(anchor_line(f"{letter}-T", 1,
                                 ["判断：问题清单由本章各节初步判定汇总，未新增事实"],
                                 confidence="中"))

    # 章末汇总句：清点本章结论中可直接核对/为推断各多少条（正文不写「置信度」）
    from scatter_writer import _chapter_summary
    lines.append("")
    lines.append(_chapter_summary(kinds.count("check"), kinds.count("infer")))
    lines.append(anchor_line(f"{letter}-SUM", 1,
                             ["判断：本章结论按可核对与推断两类清点，未新增事实"],
                             confidence="中"))

    # 缺口不进正文：只登记，由 08 汇总为《数据缺口清单》独立成件
    seen, uniq = set(), []
    for g in gaps:
        key = (g["id"], g["内容"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(g)
    uniq.sort(key=lambda g: g["id"])
    lines.append("")
    lines.append("## 来源映射")
    lines.append("")
    for n in nodes:
        if n["level"] != 1 and letter in ("D", "F"):
            continue
        m = mapping_by_id.get(n["id"], {})
        if m.get("sources"):
            lines.append(f"- {n['title']} ← 源报告「{'、'.join(m['sources'])}」")
        else:
            lines.append(f"- {n['title']} ← 按项目总体情况与相邻章节归纳")
    return "\n".join(enforce_cap(lines, letter)) + "\n", uniq, []


def build_chapter(letter, nodes, src, mapping_by_id, facts, materials, gap_mode,
                  scatter_mode=False):
    """返回 (markdown, gaps[], notes)"""
    lines, gaps, notes = [], [], []
    chapter_title = None
    h2_nodes = []
    for n in nodes:
        if n["level"] == 1:
            chapter_title = n["title"]
        else:
            h2_nodes.append(n)
    if not chapter_title:
        return "", [], notes
    lines.append(f"# {chapter_title}")

    def gap_slug(index):
        return f"缺项-{letter}{index:02d}"

    def emit_sourceline(rng):
        a, b = rng
        lines.append("")
        lines.append(f"（数据来源：项目资料、既有成果报告；来源：源报告段 {a}–{b}）")

    counter = [0]
    gap_alias = {}
    # 章内段落去重：H1 引言与其映射到的 H2 节可能指向源报告同一章节
    # （如 H1「概述」与 H2「项目概况」都落在源报告「概述」下），
    # 不去重会把同一段文字在章里贴两遍。
    seen_paras = set()
    # 同一张源表可能被多个模板节引用（多个节点映射到源报告同一节），
    # 整表搬一次即可；重复搬运会把现状章撑成上万字的表格堆。
    # 记录已搬过的表序号，第二次遇到就只给一句指路，不再重复贴表。
    seen_tables = set()

    def remap_gaps(text):
        """把源报告沿用的缺项编号改为本章编号，保证章内连续且不跨章复用。"""
        def repl(m):
            oid = m.group(1)
            if oid not in gap_alias:
                counter[0] += 1
                gap_alias[oid] = gap_slug(counter[0])
            return f"【{gap_alias[oid]}：{m.group(2)}｜建议来源：{m.group(3)}】"
        return GAP_RE.sub(repl, text)

    def register_gaps(text):
        for m in GAP_RE.finditer(text):
            gaps.append({"id": m.group(1), "内容": m.group(2).strip(), "来源": m.group(3).strip()})

    # H1 引言（只取到第一个子标题前；含表格）
    h1 = nodes[0] if nodes and nodes[0]["level"] == 1 else None
    h1_texts = []
    if h1:
        m = mapping_by_id.get(h1["id"], {})
        if m.get("status") != "gap" and m.get("sources"):
            for st in m["sources"]:
                tx, rng = src.block_texts(st)
                if tx:
                    h1_texts.extend(tx)
        if h1_texts:
            h1_rng = None
            for kind, v in h1_texts:
                if kind == "table":
                    t = src.tables[src.table_at[v]] if v in src.table_at else None
                    if t:
                        ti = src.table_at[v]
                        if ti in seen_tables:
                            continue
                        seen_tables.add(ti)
                        lines.append("")
                        lines.append(md_table(t["rows"], max_rows=TABLE_MAX_ROWS))
                else:
                    if SOURCE_RE.match(v):
                        continue
                    # 「【缺口】本节无材料…」是 03 汇编的结构化标记，不是正文，不进成稿
                    if v.startswith("【缺口】"):
                        continue
                    v = remap_gaps(v)
                    v = filter_todo_sentences(v)
                    if not v.strip():
                        continue
                    # 章内去重：同一段原文只搬一次
                    if v in seen_paras:
                        continue
                    seen_paras.add(v)
                    lines.append("")
                    lines.append(v)
                    register_gaps(v)
            # H1 来源行放在本章引言之后
            m = mapping_by_id.get(h1["id"], {})
            for st in (m.get("sources") or []):
                tx, rng = src.block_texts(st)
                if tx and rng:
                    h1_rng = rng
            if h1_rng:
                emit_sourceline(h1_rng)
        else:
            lines.append("")
            if not scatter_mode:
                lines.append(f"本章依据模板章节结构与项目材料成文。「{chapter_title}」的判定由项目改造内容"
                             f"与相邻章节互证得出；据现有材料可判断的是本章定性结论，"
                             f"需现场核实的是分项量级与点位。")

    # 节列表
    if letter == "D":
        spec = D_SECTIONS
        src_titles = []
        m = mapping_by_id.get("t18", {})
        if m.get("status") != "gap":
            for st in (m.get("sources") or []):
                rng = src.span(st)
                if not rng:
                    continue
                for h in src.headings:
                    if rng[0] < h["i"] < rng[1] and h["level"] == 2:
                        src_titles.append(h["text"])
        if src_titles:
            spec = src_titles
    elif letter == "F":
        spec = []
    else:
        spec = [n["title"] for n in h2_nodes]

    # ── 散材料模式：不走「摘句搬运」，改为按事实库 + 章节契约组织成文 ──
    # 适用场景：源报告不是成篇现状分析，而是设计成果/清单/图纸等散材料
    # （source_report = "__assemble__" 模式）。此时摘句只会得到目录行、
    # 规范条文编号等碎片，拼不成可交付正文。
    if scatter_mode and letter != "F" and BASELINE_WRITER is not None:
        text, _ = BASELINE_WRITER.chapter(letter, chapter_title, spec)
        return text.rstrip() + "\n", set(), []

    if letter == "F" and BASELINE_WRITER is not None:
        text, _ = BASELINE_WRITER.chapter(letter, chapter_title, spec)
        return text.rstrip() + "\n", set(), []

    if scatter_mode and letter != "F":
        # F 章（附件）不走散材料路径：附件内容来自分类清单与图件台账，
        # 由 build_f 负责（附件一~四 + 三张表），与源报告形态无关。
        return build_scatter_chapter(letter, nodes, h2_nodes, chapter_title,
                                      mapping_by_id, facts, spec, lines, gaps,
                                      gap_slug, counter, materials=materials)

    for title in spec:
        if letter in ("D", "F"):
            node = None
            match_src = None
        else:
            node = next((n for n in h2_nodes if norm_title(n["title"]) == norm_title(title)), None)
            match_src = None
        lines.append("")
        lines.append(f"## {title}")
        body_parts = []
        rng_used = None
        if node:
            m = mapping_by_id.get(node["id"], {})
            for st in (m.get("sources") or []):
                tx, rng = src.block_texts(st)
                if tx:
                    rng_used = rng
                    body_parts.extend(tx)
        else:
            # 在源报告中找同名/近名小节
            best, best_score = None, 0.0
            for h in src.headings:
                if h["level"] != 2:
                    continue
                sc = similarity(title, h["text"])
                if sc > best_score:
                    best, best_score = h["text"], sc
            if best and best_score >= 0.75:
                tx, rng = src.block_texts(best)
                if tx:
                    rng_used = rng
                    body_parts.extend(tx)
                    match_src = best
        # 「【缺口】本节无材料…」是 03 汇编的结构化标记，不是正文：
        # 先剔掉，剔完若本节确实没有可搬的正文，走下面的「初步判定」分支。
        body_parts = [(k, v) for k, v in body_parts
                      if not (k == "text" and v.startswith("【缺口】"))]
        if body_parts:
            texts_only = [v for k, v in body_parts if k == "text" and not SOURCE_RE.match(v)]
            text = "\n\n".join(texts_only)
            text = remap_gaps(text)
            # 作者待办语不是现状陈述，摘句阶段就剔掉（否则 CHK-BAN-1 必然命中）
            text = filter_todo_sentences(text)
            # 章内去重：同一段原文只搬一次（H1 引言可能已搬过同一段）
            if text in seen_paras:
                text = ""
            elif text.strip():
                seen_paras.add(text)
            if letter == "C" and "建筑现状" in title:
                text = ("本节从建筑安全、历史风貌与合规性三个方面梳理建筑现状，"
                        "结构安全等级、风貌破损与报建手续等指标见下文事实与缺项。\n\n") + text
            caps = {"C": 1200, "D": 900, "E": 1300, "F": 900}.get(letter)
            if caps and len(text) > caps:
                # 按句裁剪，控制章长（不丢来源、不改数字）
                cut = text[:caps]
                pos = max(cut.rfind("。"), cut.rfind("；"))
                text = cut[:pos + 1] if pos > 300 else cut[:caps].rstrip("；;，,") + "……"
            lines.append("")
            lines.append(text)
            if letter == "E" and "综合现状研判" in title and "问题总结" not in text:
                lines.append("")
                lines.append("问题总结：本章所列问题与判断均已在第 3、4 章逐条登记，"
                             "本节仅作归纳性表述，不新增判断、不补数据。")
            register_gaps(text)
            for kind, v in body_parts:
                if kind == "table":
                    t = src.tables[src.table_at[v]] if v in src.table_at else None
                    if t:
                        ti = src.table_at[v]
                        if ti in seen_tables:
                            # 已在前面章节搬过这张表：只留指路，不重复贴表
                            lines.append("")
                            lines.append(f"（本节涉及的表格已于本章前文列出，不重复搬运；"
                                         f"完整表见源报告）")
                            continue
                        seen_tables.add(ti)
                        lines.append("")
                        lines.append(md_table(t["rows"], max_rows=TABLE_MAX_ROWS))
            if rng_used:
                emit_sourceline(rng_used)
        else:
            gid = gap_slug(counter[0] + 1)
            counter[0] += 1
            hint = "项目资料、专项调查台账"
            if letter == "D":
                hint = "项目资料、体检报告问题清单"
            if letter == "F":
                hint = "项目资料、图件与台账"
            content = f"{title}（现有材料无直接表述）"
            lines.append("")
            # v2 结论优先：没有直接表述也要给判定，不留空、不写「资料缺失」。
            if letter == "D":
                from scatter_writer import _profile
                prof = _profile(title)
                cat = re.sub(r"问题$", "", title)
                lines.append(f"初步判定：{prof['judge']}。")
                lines.append("")
                lines.append(f"表现：{title}对应的现象在片区改造内容中已有体现，"
                             f"具体点位与量级需现场核实。")
                lines.append("")
                lines.append(f"判断依据：据本节所列项目材料，{prof['basis']}；"
                             f"按问题清单口径核对，需现场核实的是{prof['verify']}。")
                lines.append("")
                lines.append("成因：建成年代久、长期缺少系统性维护投入，且更新改造此前未成片实施。")
                lines.append("")
                lines.append(f"影响：制约街区安全、民生品质与运营效率；对应更新方向见"
                             f"《现状综合研判与更新思路》的《核心更新改造方向》节（推断，需现场核实）。")
                lines.append("")
                lines.append(f"对应更新方向：见《现状综合研判与更新思路》的"
                             f"《核心更新改造方向》节，按{cat}方向推进（推断，需现场核实）。")
            else:
                from scatter_writer import _profile
                prof = _profile(title)
                lines.append(f"初步判定：{prof['judge']}。")
                lines.append("")
                lines.append(f"判断依据：据现有项目材料，{prof['basis']}；"
                             f"需现场核实的是{prof['verify']}。")
            lines.append("")
            lines.append("（数据来源：项目改造内容与现有材料）")
            gaps.append({"id": gid, "内容": content, "来源": hint})

    # 附件章：四节 + 三张表（TBL-03/04/05），数据只来自分类清单与源报告，不编
    if letter == "F":
        f_lines, f_gaps = build_f(src, materials, gap_slug)
        lines.extend(f_lines)
        gaps.extend(f_gaps)

    # 缺项表去重/排序
    seen, uniq = set(), []
    for g in gaps:
        key = (g["id"], g["内容"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(g)
    uniq.sort(key=lambda g: g["id"])
    # 散材料/降级稿：缺口只登记、不进正文（由 08 汇总为《数据缺口清单》独立成件）
    if uniq and not scatter_mode:
        lines.append("")
        lines.append("## 本节缺项表")
        lines.append("")
        lines.append("| 编号 | 缺项内容 | 建议来源 | 现有替代性判断 |")
        lines.append("| --- | --- | --- | --- |")
        for g in uniq:
            lines.append(f"| {g['id']} | {g['内容']} | {g['来源']} | 未作推断，按缺项留痕 |")
    # 来源映射
    lines.append("")
    lines.append("## 来源映射")
    lines.append("")
    for n in nodes:
        if n["level"] != 1 and letter in ("D", "F"):
            continue
        m = mapping_by_id.get(n["id"], {})
        if m.get("sources"):
            lines.append(f"- {n['title']} ← 源报告「{'、'.join(m['sources'])}」")
        elif scatter_mode:
            lines.append(f"- {n['title']} ← 按项目总体情况与相邻章节归纳")
        else:
            lines.append(f"- {n['title']} ← 缺项（原始报告无可引用来源）")
    return "\n".join(enforce_cap(lines, letter)) + "\n", uniq, notes


CHAR_CAPS = {"A": 8500, "B": 8800, "D": 7800, "E": 4300, "F": 3900}


def enforce_cap(lines, letter):
    """按章正文上限压缩最长段落（段落尾部按句截断），表格/标题/来源行/依据块不计入。"""
    # 数据化成文按内容写足，不做字数截断（截断会把结论切半）
    if BASELINE_WRITER is not None:
        return lines
    cap = CHAR_CAPS.get(letter)
    if not cap:
        return lines
    def eligible(s):
        t = s.strip()
        return (bool(t) and not t.startswith(("#", "|", "（数据来源：", "〔锚点")))
    for _ in range(60):
        idxs = [i for i, s in enumerate(lines) if eligible(s)]
        total = sum(len(lines[i]) for i in idxs)
        if total <= cap:
            break
        idx = max(idxs, key=lambda i: len(lines[i]))
        over = total - cap
        keep = max(200, len(lines[idx]) - over - 100)
        cur = lines[idx]
        if len(cur) <= 260:
            break
        cut = cur[:keep]
        pos = max(cut.rfind("。"), cut.rfind("；"))
        lines[idx] = (cut[:pos + 1] if pos > 150 else cut.rstrip("；;，, ") + "……")
    return lines


def build_f(src, materials, gap_slug):
    """附件一~三逐节成文；数据缺口清单（附件四）留给 assemble_chapters 汇总。"""
    out, gaps = [], []
    files = materials.get("files", []) if materials else []
    counter = [0]
    alias = {}

    def remap(text):
        def repl(m):
            oid = m.group(1)
            if oid not in alias:
                counter[0] += 1
                alias[oid] = gap_slug(len(gaps) + counter[0])
            return f"【{alias[oid]}：{m.group(2)}｜建议来源：{m.group(3)}】"
        return GAP_RE.sub(repl, text)

    # 附件一~三的表与正文数字来自建档环节的分类清单（过程文件，非源报告摘句），
    # 依据块用「内部依据」指回该文件：09 三件套会读文件核对数值，不再要求来源行。
    INT_BASIS = "过程文件/01-建档与材料归集/分类清单/分类清单.md"

    def material_of(rng):
        """从汇编转储块里取原始材料名（块行带【来源：x】）；取不到返回空。"""
        if not rng:
            return ""
        a, b = rng
        for ln in src.blocks[a:b]:
            m = re.search(r"【来源：([^】]+)】", ln)
            if m:
                return m.group(1).strip()
        return ""

    def section(no, title, intro, table_rows, gap_note=None, extra_rows=None):
        out.append("")
        out.append(f"## {title}")
        out.append("")
        st = src.span(title)
        body = []
        aidx = [0]

        def emit(text=None, field=None):
            """正文/表格 + 紧随其后的〔锚点〕依据块（成稿口径：不写来源行）。"""
            if text is not None:
                out.append(text)
            aidx[0] += 1
            out.append(anchor_line(f"F-{no:02d}", aidx[0], [field or f"依据：{INT_BASIS}"]))

        if st:
            tx, rng = src.block_texts(title)
            body = [v for k, v in tx if k == "text" and not SOURCE_RE.match(v)]
            mat = material_of(rng)
            if body:
                text = remap("\n\n".join(body))
                if len(text) > 900:
                    cut = text[:900]
                    pos = max(cut.rfind("。"), cut.rfind("；"))
                    text = cut[:pos + 1] if pos > 300 else cut[:900].rstrip("；;，,") + "……"
                for para in [p for p in text.split("\n\n") if p.strip()]:
                    emit(para, f"材料：{mat}" if mat else None)
                gaps.extend({"id": m.group(1), "内容": m.group(2).strip(), "来源": m.group(3).strip()}
                            for m in GAP_RE.finditer(text))
            else:
                emit(intro)
        else:
            emit(intro)
            if gap_note:
                idx = len(gaps) + 1
                gid = gap_slug(idx)
                emit(f"本节按项目材料整理；据现有材料可判断的部分已列示，"
                     f"其余需现场核实（{gap_note}）。")
                gaps.append({"id": gid, "内容": gap_note, "来源": "项目资料、专项台账"})
        for rows in (extra_rows, table_rows):
            if not rows:
                continue
            emit("表中只列本次实际收到的材料，已提供项注明来源路径。")
            emit(md_table(rows))

    # 附件一是「家底台账」：材料可能有数百份，全量列进正文会让报告没法看，
    # 也会把 F 章撑到 5 万字（契约上限 3900）。
    # 做法：先给按类别汇总的统计表（这是现状分析真正需要的口径），
    # 再摘录前若干份明细，明细全量在《输入材料补齐确认单》与过程文件里可查。
    by_cat = {}
    for f in files:
        label = str(f.get("category_label") or "未分类")
        by_cat[label] = by_cat.get(label, 0) + 1
    # 表头刻意不与契约表的表头重名（契约表头是「序号|文件名称|类型|主要用途|本次使用情况」），
    # 避免 09 的契约检查把汇总表误认成 TBL-03。
    rows_sum = [["序号", "资料类别汇总", "份数", "用途"]]
    for i, (label, n) in enumerate(sorted(by_cat.items(), key=lambda kv: (-kv[1], kv[0])), 1):
        rows_sum.append([str(i), label, str(n), "项目基础资料，已登记"])
    if len(rows_sum) == 1:
        rows_sum.append(["1", "（本次投递区无文件）", "0", "—"])

    rows1 = [["序号", "文件名称", "类型", "主要用途", "本次使用情况"]]
    for i, f in enumerate(files[:TABLE_MAX_ROWS], 1):
        rows1.append([str(i), f.get("rel", ""), f.get("category_label", ""), "项目基础资料",
                      "已登记，按需引用" if not f.get("pending") else "待确认后使用"])
    if len(rows1) == 1:
        rows1.append(["1", "（本次投递区无文件）", "—", "—", "—"])
    intro = ("本节按投递区登记资料清单，未改动原件。共登记 {} 份材料，"
             "按类别构成见下表；明细摘录前 {} 份，"
             "完整清单见《输入材料补齐确认单》与过程文件 01 的分类清单。"
             .format(len(files), min(len(files), TABLE_MAX_ROWS)))
    section(1, "附件一 现有基础资料清单", intro, rows1, "现有基础资料清单",
            extra_rows=rows_sum)

    rows2 = [["序号", "图件/影像", "应表达的内容", "状态"]]
    media = [f for f in files if str(f.get("ext", "")).lower() in ("png", "jpg", "jpeg", "dwg", "pdf", "pptx")]
    for i, f in enumerate(media[:8], 1):
        rows2.append([str(i), f.get("rel", ""), "现状影像/图纸/汇报资料",
                      "已提供" if not f.get("pending") else "待确认"])
    if len(rows2) == 1:
        rows2.append(["1", "（现有材料未附图件与影像）", "位置关系以四至、面积与周边道路文字描述", "需补测"])
    section(2, "附件二 图件与影像资料", "本节列示现有材料中的图件与影像；影像不足部分需现场补测后归档。",
            rows2, "图件与影像资料")

    rows3 = [["序号", "台账名称", "应包含字段", "状态"]]
    ledgers = [("建筑安全排查台账", "分栋结构安全等级、危房数量与面积"),
               ("消防设施配置台账", "消火栓数量与间距、消防通道宽度、消防站覆盖半径"),
               ("市政管网普查台账", "各街巷雨污给水燃气电力通信管线长度、管径、埋深"),
               ("人口与住户台账", "常住户数、人数、年龄结构、低保与特殊群体数量"),
               ("沿街商户台账", "商户数量、经营面积、业态构成、空置率、租金水平"),
               ("历史文化资源台账", "历史建筑名录、工业遗产本体、老字号、古树名木"),
               ("公共空间与绿地清单", "口袋公园、广场、绿地面积与位置"),
               ("民意调研资料", "诉求排序、改造意愿、阻工风险点")]
    received = [str(f.get("rel", "")) for f in files]
    got, miss = [], []
    for n, fld in ledgers:
        key = n.replace("台账", "").replace("资料", "").replace("清单", "")
        (got if any(key and key in r for r in received) else miss).append((n, fld))
    for i, (n, fld) in enumerate(got, 1):
        rows3.append([str(i), n, fld, "已收到"])
    if not got:
        # 干净版：不逐条列缺项，一句话交代口径
        rows3.append(["1", "（现有材料未附专项台账）", "现状判断以登记表为限，检测类指标需现场核", "需补录"])
    section(3, "附件三 专项台账", "本节只列现有材料中的专项台账；检测类指标需现场核实后补录。",
            rows3, "专项台账")

    # 附件四「引用政策文件与标准规范清单」暂不生成：本次输入未附任何政策与标准原文，
    # 正文 0 处引用具体文件、文号或条款；单独立节只剩一句"没有清单可列"，
    # 等于把「这也没有」标在报告最后，故整节省略、后续附件编号顺延。
    # 将来真拿到政策原文时这样加回来：在此处补
    #   section(4, "附件四 引用政策文件与标准规范清单", 导语, 清单表, "政策文件")
    # 并把各处编号后移——check_chapter.py 的 sections/section_keywords、
    # 各 章节契约.yaml 的 sections 与 t23-r1 titles、表格契约 TBL-06 的编号
    # （数据缺口索引顺延为 附件五）；每条依据必须能指到正文引用位置才允许列。
    return out, gaps


def _read_filled_draft(filled_dir, letter, safe_title):
    """读取 agent/大模型回填的章节稿。

    回填稿是「写作任务包」模式的产物：脚本生成提示词与素材清单 → agent 按 skill 规范
    成文 → 写回本章文件。判定回填有效的条件：
      * 文件存在且非空；
      * 已写实际正文（含一级标题）；
      * 模板里的 HTML 注释提示未清除（说明 agent 没写完）。
    未回填的章返回空串，走确定性成文降级。
    """
    if not filled_dir:
        return ""
    if not os.path.isdir(filled_dir):
        return ""
    try:
        names = sorted(os.listdir(filled_dir))
    except OSError:
        return ""
    cands = [os.path.join(filled_dir, n) for n in names
             if n.startswith(f"{letter}_") and n.endswith(".md")]
    if not cands:
        return ""
    # 优先与本章标题同名的文件
    target = next((c for c in cands if safe_title in os.path.basename(c)), cands[0])
    try:
        with open(target, encoding="utf-8-sig") as f:
            text = f.read()
    except OSError:
        return ""
    if not text.strip():
        return ""
    if f"# {CHAPTER_TITLES.get(letter, '')}" not in text and "# " not in text:
        return ""
    # 模板提示未清 → 视为未回填
    if "<!-- 请由大模型成文后写入本文件" in text:
        return ""
    body = "\n".join(ln for ln in text.splitlines()
                     if not ln.strip().startswith("<!--")).strip()
    return body or ""


def _extract_gaps(text):
    """从回填稿里抽取缺项登记（供 08 汇总数据缺口清单）。"""
    out = []
    for m in GAP_RE.finditer(text or ""):
        out.append({"id": m.group(1), "内容": m.group(2).strip(),
                    "来源": m.group(3).strip()})
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--template-outline", required=True)
    ap.add_argument("--source-dump", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--materials", required=False)
    ap.add_argument("--facts", required=False)
    ap.add_argument("--out", required=True)
    ap.add_argument("--gap-mode", action="store_true")
    ap.add_argument("--scatter-mode", action="store_true",
                    help="散材料模式：源报告为设计成果/清单等散件时，按事实库成文而非摘句")
    ap.add_argument("--filled-dir", default="",
                    help="agent/大模型回填稿目录（写作任务包模式）；有回填则优先采用")
    ap.add_argument("--baseline", default="",
                    help="包内依据.json（05 ingest_baseline 产物）；有则走数据化成文")
    ap.add_argument("--baseline-md", default="",
                    help="包内依据.md 的工作区相对路径，供锚点内部依据引用")
    ap.add_argument("--project", default="", help="项目名称")
    ap.add_argument("--narrative", default="",
                    help="项目叙述.json（项目侧配置，含街道/社区/道路/门牌/工期/资金来源等）；"
                         "数据化成文缺此文件直接报错，不做静默兜底")
    args = ap.parse_args(argv)

    global BASELINE_WRITER
    narrative = {}
    if args.narrative:
        if not os.path.exists(args.narrative):
            raise SystemExit(f"项目叙述配置不存在：{args.narrative}")
        with open(args.narrative, encoding="utf-8-sig") as f:
            narrative = json.load(f)
    # 区域归属：只读本轮 项目识别.json（开工.py 写入），不做跨轮记忆。
    # 没有 --area 就不配置区域，正文不硬凑「区域分列」。
    # 必须早于 DataWriter 构造：DataWriter.__init__ 会把区域关键词并入台账社区识别，
    # 晚于构造则区域台账行（如示例社区B）在建对象时已被筛掉。
    areas = {}
    pid_path = os.path.join("过程文件", "00-总控", "项目识别.json")
    if os.path.exists(pid_path):
        try:
            with open(pid_path, encoding="utf-8-sig") as fh:
                areas = (json.load(fh) or {}).get("areas") or {}
        except (OSError, ValueError):
            areas = {}
    if areas:
        from scatter_writer import set_area_map
        set_area_map(areas)
        print(f"[06] 区域归属（本轮）：{ '、'.join(areas.keys()) }")

    if args.baseline and os.path.exists(args.baseline):
        from baseline_writer import DataWriter
        baseline = json.load(open(args.baseline, encoding="utf-8"))
        BASELINE_WRITER = DataWriter(baseline, read_blocks_source(args.source_dump),
                                     args.baseline_md, args.project, narrative)
        print(f"[06] 数据化成文：包内依据已接入（{args.baseline}）")
    if narrative:
        from scatter_writer import set_narrative
        set_narrative(narrative)

    template = json.load(open(args.template_outline, encoding="utf-8"))
    mapping = load_yaml(args.mapping) or {}
    nodes = [{"id": n["id"], "level": n["level"], "title": n["title"]} for n in mapping.get("nodes", [])]
    if not nodes:
        raise SystemExit("映射为空：%s" % args.mapping)
    mapping_by_id = {n["id"]: n for n in mapping.get("nodes", [])}
    src = Source(args.source_dump)
    materials = json.load(open(args.materials, encoding="utf-8")) if args.materials and os.path.exists(args.materials) else {}
    facts = json.load(open(args.facts, encoding="utf-8")) if args.facts and os.path.exists(args.facts) else {}

    # 按模板 H1 分组
    groups, cur = {}, None
    for n in nodes:
        if n["level"] == 1:
            cur = H1_TO_CHAPTER.get(n["title"])
            if cur:
                groups.setdefault(cur, []).append(n)
        elif cur:
            groups[cur].append(n)

    os.makedirs(args.out, exist_ok=True)
    # 出图数据：只导出聚合值（用地构成、高度分位、台账层数），供 08 直接画图
    if BASELINE_WRITER is not None:
        try:
            with open(os.path.join(args.out, "图表数据.json"), "w", encoding="utf-8") as f:
                json.dump(BASELINE_WRITER.chart_data(), f, ensure_ascii=False, indent=1)
        except Exception as exc:  # noqa: BLE001
            print(f"[06] 图表数据导出跳过：{exc}")
        # 工程量逐列映射快照：源表头原文 + 明细 + 合计，供 09 工程量闸门逐数核对
        try:
            with open(os.path.join(args.out, "工程量核算.json"), "w", encoding="utf-8") as f:
                json.dump(BASELINE_WRITER.works_snapshot(), f, ensure_ascii=False, indent=1)
        except Exception as exc:  # noqa: BLE001
            print(f"[06] 工程量核算导出跳过：{exc}")
    summary = []
    gap_registry = []
    for letter in CHAPTERS:
        if letter not in groups:
            continue
        title = next(n["title"] for n in groups[letter] if n["level"] == 1)
        safe = re.sub(r'[\\/:*?"<>|]', '', title)
        fn = os.path.join(args.out, f"{letter}_{safe}.md")

        # 优先采用 agent/大模型回填稿（写作任务包模式）：
        # 回填稿由「章节写作任务」提示词包驱动，产出质量高于确定性摘句成文。
        filled = _read_filled_draft(args.filled_dir, letter, safe)
        if filled:
            with open(fn, "w", encoding="utf-8") as f:
                f.write(filled)
            gaps = _extract_gaps(filled)
            summary.append({"chapter": letter, "file": os.path.basename(fn),
                            "chars": len(filled), "gaps": len(gaps), "source": "agent回填"})
            print(f"  {letter} {os.path.basename(fn)}  {len(filled)} 字  缺项 {len(gaps)}  [agent回填]")
            continue

        md, gaps, _ = build_chapter(letter, groups[letter], src, mapping_by_id, facts,
                                    materials, args.gap_mode, args.scatter_mode)
        with open(fn, "w", encoding="utf-8") as f:
            f.write(md)
        summary.append({"chapter": letter, "file": os.path.basename(fn), "chars": len(md),
                        "gaps": len(gaps), "source": "确定性成文"})
        for g in gaps:
            gap_registry.append({"chapter": letter, "id": g["id"],
                                 "内容": g["内容"], "来源": g["来源"],
                                 "现有替代性判断": "按正文初步判定与置信度处理"})

    print(f"[06] 分章稿 {len(summary)} 章 → {args.out}")
    for s in summary:
        print(f"  {s['chapter']} {s['file']}  {s['chars']} 字  缺项 {s['gaps']}  [{s['source']}]")
    with open(os.path.join(args.out, "_生成摘要.json"), "w", encoding="utf-8") as f:
        json.dump({"schema": "cardo/chapters/v1", "chapters": summary,
                   "gapMode": bool(args.gap_mode),
                   "scatterMode": bool(args.scatter_mode)}, f, ensure_ascii=False, indent=1)
    # 缺口登记（不进正文）：供 08 汇总为《数据缺口清单》独立成件
    with open(os.path.join(args.out, "_缺口登记.json"), "w", encoding="utf-8") as f:
        json.dump({"schema": "cardo/gap-registry/v1", "gaps": gap_registry},
                  f, ensure_ascii=False, indent=1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
