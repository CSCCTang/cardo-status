#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""build_write_tasks.py -- 生成「章节写作任务包」（cardo/write-tasks/v1）。

为什么要有这个环节：
  流水线的脚本必须是**确定性、可审计、零 LLM 调用**的——同一输入必须产出同一结果，
  否则 09 的回溯校验与基线比对都失去意义。但「把散材料组织成可交付的现状分析正文」
  恰恰是需要语言模型的活：摘句搬运得到的只是碎片。

  因此采用与「联网检索补料」同一套模式（见 开工.py --research）：
      脚本生成任务包（输入） → 人/agent 用大模型成文并回填（输出） → 流水线消费回填稿。
  这样既有大模型的成文质量，又不破坏流水线的可审计性与离线可跑性。

产物（默认 过程文件/06-分章成文/写作任务/）：
  * 写作任务.md      —— 人可读的提示词与素材清单（含各章输入、事实、缺项、契约要求）
  * 写作任务.json    —— 机器可读的任务包（供 agent 批量处理）
  * 分章初稿_待回填/  —— 每章一个空的 .md 模板，agent 直接写入即可

用法：
    python build_write_tasks.py --project-dir <工作区> [--out <目录>] [--chapters A,B,C]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _banned import TODO_PHRASES  # noqa: E402

CHAPTERS = ["A", "B", "C", "D", "E", "F"]
CHAPTER_NAMES = {
    "A": "项目概述",
    "B": "片区基础概况分析",
    "C": "现状详细分析",
    "D": "核心问题汇总",
    "E": "现状综合研判与更新思路",
    "F": "附件",
}

# 表位 → 锚点节点：材料没覆盖时，正文给「估算值 + 反算依据」并标注估算，
# 不能留一节空白，也不能写「资料缺失」（v2 结论优先）。
# 与 表格契约.yaml 的 anchor_node 对齐。
TABLE_SLOT_NODES = {
    "t10": "TBL-07（表 3.1 用地与空间利用统计）",
    "t02": "TBL-01（投资估算费用构成表）",
}

# D 章（核心问题汇总）与 F 章（附件）在模板里只有 H1，节标题由写作规范固定
# （逐字与 06-分章成文/scripts/gen_chapters.py 的 D_SECTIONS / F_SECTIONS 一致）
D_SECTIONS = ["街巷空间与肌理问题", "建筑安全与风貌破损问题", "文脉保护与活化不足问题",
              "业态老旧与活力不足问题", "交通通行与消防隐患突出问题", "基础设施老旧滞后问题",
              "民生配套与人居品质偏低", "街区治理与运营短板"]
F_SECTIONS = ["附件一 现有基础资料清单", "附件二 图件与影像资料", "附件三 专项台账"]

# 各章在写作规范下的额外要求（与 06-分章成文/规则/写作规范.yaml 对齐）
CHAPTER_BRIEF = {
    "A": "写清项目是什么：名称、建设单位、建设性质、建设内容与规模、投资与资金来源、"
         "编制依据、调研范围与期限。材料不足时按现有材料给初步判定并标注推断依据与置信度，"
         "不要写「资料缺失」，缺口只登记进数据缺口清单。",
    "B": "写清片区是什么：区位与周边、规划要求与管控边界、历史沿革与文脉。"
         "每节首段先给初步判定（固定句式：初步判定：…（推断依据：…；置信度：中/低；区间：a～b））；"
         "老旧小区方向不写「历史沿革」实质内容时，明确声明该节不涉及并说明理由。",
    "C": "写现状八个方面：土地与空间、建筑与风貌、人口与人居、业态与经营、交通与通行、"
         "市政基础设施、公共服务配套、生态与人居环境。每方面都要走「判定→描述→判断→归因」，"
         "数据不足时按同类项目公开区间给推断并标置信度，不得整节只写缺口。",
    "D": "按八类问题（街巷空间与肌理、建筑安全与风貌破损、文脉保护与活化不足、业态老旧、"
         "交通通行与消防隐患、基础设施老旧滞后、民生配套偏低、治理与运营短板）逐类归纳，"
         "每条问题用 P-NN 编号，并写全五要素：表现、证据或推断依据、成因、影响、对应更新方向；"
         "数据不足也要给出初步判定（推断依据＋置信度），不得整列「资料缺失」。",
    "E": "写研判与思路：综合现状研判（评价、潜力、问题总结、差距）、总体更新思路"
         "（原则）、核心更新改造方向。每条对策用 S-NN 编号并对应 P-NN 问题。",
    "F": "附件：现有基础资料清单、图件与影像资料、专项台账、引用政策文件与标准规范清单。"
         "只列本次实际收到的材料；未收到的图件不列「待补」、不出缺项，也不影响结论。",
}


def load_json(path, default=None):
    try:
        with open(path, encoding="utf-8-sig") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default if default is not None else {}


def load_yaml(path):
    """读 yaml。优先 PyYAML；没有则用内置的 _yaml_lite 子集解析器。

    为什么要有回退：便携环境里 PyYAML **可能没装**（实测缺 yaml），
    而此前 `except: return {}` 是**静默失败**——映射读成空 dict，
    06 只生成 D/F 两章（它们的节标题是硬编码的，不依赖 yaml），
    A/B/C 三章直接消失，且不报任何错。
    """
    try:
        import yaml
        with open(path, encoding="utf-8-sig") as f:
            return yaml.safe_load(f) or {}
    except ImportError:
        pass
    except Exception:
        pass
    try:
        import _yaml_lite
        with open(path, encoding="utf-8-sig") as f:
            return _yaml_lite.parse(f.read()) or {}
    except Exception as exc:
        print(f"[写作任务] 读不了 {path}：{exc}", file=sys.stderr)
        return {}


def norm(t):
    return re.sub(r"\s+", "", str(t or ""))


def load_ledger(project_dir: Path) -> dict:
    """读《检索台账》有效条目。校验规则唯一实现在 09/scripts/_ledger.py，这里只调用。"""
    scripts = str(Path(__file__).resolve().parents[2] / "09-三重自检与边界" / "scripts")
    if scripts not in sys.path:
        sys.path.insert(0, scripts)
    try:
        import _ledger
    except ImportError:
        print("[写作任务] 未找到 09/scripts/_ledger.py，本章不注入外部条目", file=sys.stderr)
        return {"path": "", "valid": {}}
    return _ledger.load(str(project_dir))


# scope 写这些词的条目视为全章可用（其余按 scope 与章/节标题互含匹配）
LEDGER_WILDCARD = ("全文", "通用", "全部", "不限", "全篇")


def ledger_for_chapter(led: dict, chapter_title: str,
                       section_titles: list[str]) -> list[dict]:
    """挑出本章可引用的台账条目：只给 scope 命中的**有效**条目。

    不给全量台账是有意的——把整本台账摊在写作任务里，等于鼓励"先写结论再找依据"。
    此处只列"这一章用得上、且已经落盘可核"的条目，引用时写 `[S-012]`。
    """
    keys = [norm(chapter_title)] + [norm(t) for t in section_titles]
    keys = [k for k in keys if k]
    out: list[dict] = []
    for eid, e in (led.get("valid") or {}).items():
        scope = str(e.get("scope") or "")
        if not scope:
            continue
        hit = any(w in scope for w in LEDGER_WILDCARD)
        if not hit:
            nscope = norm(scope)
            hit = any(k in nscope or nscope in k for k in keys)
        if hit:
            out.append(dict(e, id=eid))
    return out


def chapter_sections(nodes, letter):
    """按 H1 归组出该章的节标题（与 gen_chapters 的分组口径一致）。

    必须按「章 H1 → 其后到下一个 H1 之前的 H2」切分；
    不能把所有 H2 挂到最后一个 H1 上。
    """
    h1_title = CHAPTER_NAMES.get(letter)
    out, cur = [], None
    for n in nodes:
        lvl = n.get("level")
        title = n.get("title")
        if lvl == 1:
            cur = title if norm(title) == norm(h1_title) else None
        elif cur is not None:
            out.append(n)
    return out


def chapter_h1(nodes, letter):
    """取该章的 H1 标题。"""
    h1_title = CHAPTER_NAMES.get(letter)
    for n in nodes:
        if n.get("level") == 1 and norm(n.get("title")) == norm(h1_title):
            return n.get("title")
    return h1_title


def build_tasks(project_dir: Path, out_dir: Path, chapters=None) -> dict:
    P = project_dir / "过程文件"
    D03 = P / "03-模板与源报告转储"
    D04 = P / "04-章节映射"
    D05 = P / "05-事实抽取与事实库"
    D06 = P / "06-分章成文"
    rules = Path(__file__).resolve().parents[1] / "规则"

    mapping = load_yaml(D04 / "章节映射.yaml")
    facts = load_json(D05 / "facts" / "事实库.json", {})
    pdf_evidence = [e for e in (facts.get("pdfEvidence") or []) if isinstance(e, dict)][:30]
    contract = load_yaml(rules / "章节契约.yaml")
    spec = load_yaml(rules / "写作规范.yaml")
    ledger = load_ledger(project_dir)

    nodes = mapping.get("nodes") or []
    facts_by_section = {}
    for f in facts.get("facts") or []:
        facts_by_section.setdefault(str(f.get("section") or ""), []).append(f)
    # 定性材料（03 的源报告转储）：有文字但未抽成数值事实的内容
    qualitative = collect_qualitative(project_dir / "过程文件" / "03-模板与源报告转储" / "源报告转储")

    out_dir.mkdir(parents=True, exist_ok=True)
    draft_dir = out_dir / "分章初稿_待回填"
    draft_dir.mkdir(parents=True, exist_ok=True)

    tasks = []
    for letter in (chapters or CHAPTERS):
        nodes_c = chapter_sections(nodes, letter)
        # D/F 两章在模板里只有 H1，节标题由写作规范固定（与 gen_chapters 一致）
        if letter == "D":
            nodes_c = [{"id": "t18", "level": 2, "title": t} for t in D_SECTIONS]
        elif letter == "F":
            nodes_c = [{"id": "t23", "level": 2, "title": t} for t in F_SECTIONS]
        if not nodes_c:
            continue
        title = chapter_h1(nodes, letter)
        # 本章涉及的契约节点：H1 节点 + 其下 H2 节点
        node_ids = {str(n.get("id")) for n in nodes_c}
        h1_id = next((str(n.get("id")) for n in nodes
                      if n.get("level") == 1 and norm(n.get("title")) == norm(title)), None)
        if h1_id:
            node_ids.add(h1_id)
        # 契约要求：按节点在模板中的**章节归属**取，不按 writing_unit 取。
        # 契约里 t05（片区基础概况分析 H1）标 writing_unit: A，但它的要求是
        # 「必须包含节：区位条件、规划要求、历史沿革分析」——内容属于 B 章。
        # 也就是说 writing_unit 与章节字母并非一一对应，据它取要求会整体错位一位。
        reqs = []
        for node in contract.get("nodes") or []:
            if str(node.get("id") or "") not in node_ids:
                continue
            common = (node.get("requirements") or {}).get("common") or []
            for r in common:
                if r.get("check") != "auto":
                    continue
                if r.get("type") == "section_exists":
                    reqs.append(f"必须包含节：{'、'.join(r.get('titles') or [])}")
                elif r.get("type") == "keywords_any":
                    reqs.append(f"须出现关键词（至少一）：{'、'.join(r.get('keywords') or [])}")
                elif r.get("type") == "min_chars":
                    reqs.append(f"本节正文不少于 {r.get('min')} 字")
                elif r.get("type") == "table_required":
                    reqs.append(f"须含表格：{r.get('table')}")
        # 表位缺失的**处理方式**（v2 结论优先）：
        # 契约里 TBL-07（表 3.1 用地与空间利用统计）锚在 t10，status=planned。
        # 没有实测数据时**不再写「资料缺失」**，改为给「估算值 + 反算依据」并标注估算，
        # 让读者拿到可用量级；缺口只登记进数据缺口清单。
        for n in nodes_c:
            nid = str(n.get("id"))
            if nid not in TABLE_SLOT_NODES:
                continue
            fs = facts_by_section.get(nid, [])
            if fs:
                continue
            reqs.append(f"本节对应表格 {TABLE_SLOT_NODES[nid]}：现有材料无实测数据，"
                        f"**按可反算的口径给「估算值」并同行写明反算依据与「估算」字样**"
                        f"（如：户均面积×户数）；不得编造具体数字，也不得留空节")
        # 各节素材
        sections = []
        for n in nodes_c:
            nid = n.get("id")
            fs = facts_by_section.get(str(nid), [])
            materials = []
            for f in fs[:40]:
                src = f.get("source") or {}
                materials.append({
                    "value": f.get("value"),
                    "unit": f.get("unit"),
                    "label": str(f.get("label") or "")[:120],
                    # context = 这个数属于谁（小区名/地址）——没有它，
                    # 成文只能写"122 户"这种裸数，说不清是哪个小区
                    "context": str(f.get("context") or "")[:120],
                    "quote": str(src.get("quote") or "")[:300],
                    "file": str(src.get("file") or "").replace("\\", "/").split("/")[-1],
                    "locator": src.get("locator"),
                })
            status = "gap" if not materials else "ok"
            # 定性材料：这一节**有文字**（PPT 段落、图片 OCR、文档段落），
            # 但没有可抽成事实的结构化数字。判成 gap 会让最该有内容的节留空——
            # 实测「项目概况」有整段项目信息卡文字，却因事实库里没有它的数而被判无材料。
            # 故单列一态「qual」：任务包里列出原文摘录（低置信），
            # 让写作能用上这些**定性事实**，同时明确「不要从中引数值」。
            qual = []
            if status == "gap":
                for q in (qualitative.get(n.get("title")) or [])[:8]:
                    qual.append({"text": q["text"], "file": q["file"]})
                if qual:
                    status = "qual"
            sections.append({
                "nodeId": nid, "title": n.get("title"), "status": status,
                "materials": materials, "qualitative": qual,
            })

        banned = list(TODO_PHRASES)
        led_items = ledger_for_chapter(ledger, title, [s["title"] for s in sections])
        task = {
            "chapter": letter,
            "chapterTitle": title,
            "brief": CHAPTER_BRIEF.get(letter, ""),
            "requirements": reqs,
            "sections": sections,
            "ledgerItems": led_items,
            "bannedPhrases": banned,
            "sourceLineFormat": ((spec.get("source_line") or {}).get("format")
                                  or "（数据来源：项目资料，来源：<来源>，第 N 页/段 N）"),
            "analysisChain": [s.get("step") for s in (spec.get("analysis_chain") or [])],
            "pdfEvidence": _pdf_evidence_for_chapter(pdf_evidence, nodes_c),
            "outputFile": f"{letter}_{title}.md",
        }
        tasks.append(task)
        # 空的回填模板
        tmpl = [f"# {title}", "",
                f"<!-- 请由大模型成文后写入本文件；完成后删除本注释。 -->", "",
                f"<!-- 章节：{letter} · 标题：{title} -->", "",
                "<!-- 成文要求（v2 结论优先）：",
                "     1) 标题层级：章用 #，各节用 ##，不要出现三级及以下标题；",
                "     2) 每节**第一段就是初步判定**，固定句式：",
                "        初步判定：<结论>（推断依据：<材料/公开区间>；置信度：高/中/低；区间：a～b）；",
                "     3) 正文**不得出现**「资料缺失/未提供/待补/缺项-」等口语文句；",
                "        数据不足时按现有材料或同类项目公开区间给推断，标注推断与置信度；",
                "     4) 允许使用外部公开资料（政策、片区公开统计、体检/造价指标、同类案例），",
                "        但**必须带台账编号**，如「来源：…；编号 [S-012]；检索日期：YYYY-MM-DD」；",
                "        写「据公开资料/按同类项目公开区间」却不带编号，会被自检判「空头依据」；",
                "     5) 不得编造具体数字与文号；估算值须同行标注「估算」并写反算依据；",
                "     6) 每段（含每个表格）**另起一行**写一条依据块：",
                "        〔锚点 <章>-<节>-P<段序两位>｜材料：<文件名>，<表号/页码>｜"
                "推导：a+b=c｜定稿值：k=v｜公开：S-001｜置信度：高〕；",
                "        正文本身不写编号、不写「（数据来源：…）」、不写置信度；"
                "平台据此生成干净版与《依据对照表》；",
                "     7) 正文不是表格：同一段出现 ≥3 组「数值+栋/户」即判成段罗列，"
                "多对象参数一律进表，正文只留判断与一句总量；",
                "     8) 第四章八类问题每条都要写全：表现、证据或推断依据、成因、影响、对应更新方向。 -->", ""]
        if led_items:
            tmpl += ["<!-- 本章可引用的外部条目（正文引用须写编号，如 [S-012]；编号即台账 file 指向的落盘文件）："]
            for e in led_items:
                val = f"{e.get('value', '')}{e.get('unit', '') or ''}".strip()
                tmpl.append(f"     [{e['id']}] {e.get('metric', '')}｜{val or '（已查无果）'}｜"
                            f"{e.get('publisher', '')}·{e.get('title', '')}｜{e.get('date', '')}")
            tmpl += ["-->", ""]
        for s in task["sections"]:
            # 节标题用二级：模板骨架最多到二级，写三级会被自检判 FAIL
            tmpl.append(f"## {s['title']}")
            tmpl.append("")
            if s["status"] == "gap":
                # 结论优先：没有结构化数值不等于不能判断。
                # 给**可直接照抄的推断段样例**，避免模型写成「本节资料缺失」。
                tmpl.append(f"<!-- 本节无结构化数值：**必须给初步判定**，不得写「资料缺失/未提供」。")
                tmpl.append(f"     下面这段是**格式样例**，照抄改内容即可：")
                tmpl.append(f"")
                tmpl.append(f"     初步判定：<一句话结论>（推断依据：<现有材料里的间接证据/同类项目公开区间>；"
                            f"置信度：中；区间：<a～b>）。")
                tmpl.append(f"")
                tmpl.append(f"     <展开 2～4 句：现象是什么、与什么比、差在哪量级>。")
                tmpl.append(f"")
                tmpl.append(f"     <归因：机制/历史/主体/资金/治理>。")
                tmpl.append(f"")
                tmpl.append(f"     〔锚点 <章>-<节>-P01｜判断：<一句话结论>｜"
                            f"依据：<现有材料里的间接证据>｜置信度：中〕")
                tmpl.append(f"-->")
            else:
                if s["status"] == "qual":
                    tmpl.append("<!-- 本节**有定性材料**（有文字、无结构化数值），请据以下原文摘录成文；")
                    tmpl.append("     摘录中标〔…·口径不明不得引用〕的数字**不得写进正文**；")
                    tmpl.append("     首段给初步判定（推断依据＋置信度），不得写「资料缺失」。 -->")
                    for q in s.get("qualitative") or []:
                        tmpl.append(f"     · {q['text'][:300]}｜来源：{q['file']}")
                else:
                    tmpl.append("<!-- 材料（只可用以下内容，不得新增数字）：")
                    for m in s["materials"][:20]:
                        val = f"{m['value']}{m['unit']}" if m["value"] is not None else ""
                        tmpl.append(f"     - {val}｜{m['quote'][:120]}｜来源：{m['file']}")
                tmpl.append("-->")
            tmpl.append("")
        (draft_dir / task["outputFile"]).write_text("\n".join(tmpl), encoding="utf-8")

    payload = {
        "schema": "cardo/write-tasks/v1",
        "project": (load_json(project_dir / "工作区参数.json", {}) or {}).get("project", ""),
        "chapters": tasks,
        "_说明": "每章的成文要求见 requirements/brief；材料见 sections.materials（只能引用这些内容）；"
                 "成文结果写入 分章初稿_待回填/<章>_<标题>.md 后重跑流水线 06。",
    }
    with open(out_dir / "写作任务.json", "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    _write_md(out_dir / "写作任务.md", payload, spec)
    return payload


def _pdf_evidence_for_chapter(items: list[dict], nodes_c: list[dict]) -> list[dict]:
    """从事实库的 PDF 证据里挑本章可引用的条目（弱匹配标题，匹配不到则共享前 12 条）。

    这些是**证据**，不是叙事来源：不得据此改写章节结构，也不得据此生成整章正文。
    """
    keys = [norm(str(n.get("title") or "")) for n in nodes_c]
    keys = [k for k in keys if k]
    hit = []
    for e in items:
        quote = norm(str((e.get("source") or {}).get("quote") or ""))
        label = norm(str(e.get("label") or ""))
        if any(k in quote or k in label for k in keys):
            hit.append(e)
    return (hit or items)[:12]


def _write_md(path: Path, payload: dict, spec: dict):
    lines = [f"# 章节写作任务 · {payload.get('project') or '本项目'}", "",
             "> 本文件是**给大模型的输入**。请按每章的「成文要求 + 可用材料」生成章节正文，",
             "> 写入 `分章初稿_待回填/<章>_<标题>.md`，然后重跑流水线（06 会自动采用回填稿）。", "",
             "> 硬约束（v2 结论优先）：",
             f"> - **结论优先**：每节第一段必须是「初步判定：…（推断依据：…；置信度：中/低；区间：a～b）」；",
             f"> - **正文不许写缺口**：「资料缺失/未提供/待补/缺项-」一律不出现（缺口只进 数据缺口清单.xlsx）；",
             f"> - **材料少也要判断**：按现有材料的间接证据、同类项目公开区间或公开政策给推断，标注推断与置信度；",
             f"> - **允许使用外部公开资料**：政策、片区公开统计、体检/造价指标、同类案例；"
             f"引用必须带台账编号（如 `[S-012]`）与检索日期：`来源：…；编号 [S-012]；检索日期：YYYY-MM-DD`；",
             f"> - **不许空头依据**：写「据公开资料」「按同类项目公开区间」却不带编号，自检直接判不合格；",
             f"> - **不编数字/文号**：估算值同行标注「估算」并写反算依据；按台账加总的数在所在行写「加总」「测算」；",
             f"> - **第四章五要素**：表现、证据或推断依据、成因、影响、对应更新方向，缺一不可；",
             f"> - **每段（含每个表格）末尾另起一行写一条依据块**："
             f"`〔锚点 <章>-<节>-P<段序两位>｜材料：<文件名>，<表号/页码>｜推导：a+b=c｜定稿值：k=v｜公开：S-001｜置信度：高〕`；",
             f"> - **正文自己不带标注**：不写「（数据来源：…）」、不写编号、不写置信度；"
             f"平台据依据块生成干净版与《依据对照表》；",
             f"> - **正文不是表格**：同段出现 ≥3 组「数值+栋/户」判成段罗列，多对象参数进表、"
             f"正文只留判断与一句总量；",
             f"> - 禁用词（出现即判不合格）：{'、'.join(TODO_PHRASES)}；",
             "> - 标题层级：章用一级标题（#），节用二级标题（##），**不要出现三级及以下标题**；",
             "> - 分析动作四步：描述 → 判断 → 归因 → 衔接。", ""]
    for task in payload["chapters"]:
        lines += [f"## {task['chapter']} {task['chapterTitle']}", "",
                  f"**本章任务**：{task['brief']}", ""]
        if task["requirements"]:
            lines += ["**成文要求**：", ""]
            lines += [f"- {r}" for r in task["requirements"]]
            lines += [""]
        if task.get("pdfEvidence"):
            lines += ["**可引用的 PDF 证据**（证据，不是叙事来源；引用须写「来源：<文件名> 第 N 页」，"
                      "不得据此改写章节结构或生成整章正文）：", ""]
            for e in task["pdfEvidence"]:
                src = e.get("source") or {}
                val = f"{e.get('value')}{e.get('unit') or ''}"
                lines.append(f"- {val}｜{src.get('quote', '')[:120]}｜"
                             f"来源：{str(src.get('file') or '').split('/')[-1]} · {src.get('locator')}")
            lines += [""]
        if task.get("ledgerItems"):
            lines += ["**本章可引用的外部条目**（只列已落盘、可核的台账条目；正文引用须写编号 `[S-012]`，"
                      "编号即台账 `file` 指向的过程文件）：", "",
                      "| 编号 | 指标 | 数值或区间 | 来源机构·标题 | 检索日期 | 适用范围 |",
                      "|---|---|---|---|---|---|"]
            for e in task["ledgerItems"]:
                val = f"{e.get('value', '')}{e.get('unit', '') or ''}".strip()
                lines.append(f"| [{e['id']}] | {e.get('metric', '')} | {val or '（已查无果）'} | "
                             f"{e.get('publisher', '')}·{e.get('title', '')} | "
                             f"{e.get('date', '')} | {e.get('scope', '')} |")
            lines += [""]
        lines += ["**各节与可用材料**：", ""]
        for s in task["sections"]:
            tag = ""
            if s["status"] == "gap":
                tag = "（无结构化数值，须给初步判定）"
            elif s["status"] == "qual":
                tag = "（有定性材料，无结构化数值）"
            lines.append(f"### {s['title']} {tag}")
            lines.append("")
            if s["status"] == "gap":
                lines.append("- 本节没有可用数值：**给初步判定**（推断依据＋置信度＋区间），"
                             "按同类项目公开区间或现有材料的间接证据推断；不得写「资料缺失」。")
            elif s["status"] == "qual":
                lines.append("- 本节有**定性材料**（文字），但没有可引用的结构化数值。"
                             "请据下列原文摘录成文，并给初步判定；")
                lines.append("  标〔…·口径不明不得引用〕的数字**不得写进正文**。")
                for q in s.get("qualitative") or []:
                    lines.append(f"  - {q['text'][:300]}｜来源：{q['file']}")
            else:
                for m in s["materials"][:20]:
                    val = f"{m['value']}{m['unit']}" if m["value"] is not None else ""
                    # 写成「归属｜指标｜值」——没有归属（context）时退化成「指标｜值」
                    label = m.get("label") or m.get("quote") or ""
                    ctx = m.get("context") or ""
                    if ctx:
                        lines.append(f"- {ctx}｜{label}｜{val}｜来源：{m['file']}")
                    else:
                        lines.append(f"- {label}｜{val}｜来源：{m['file']}")
            lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def collect_qualitative(dump_dir: Path) -> dict:
    """从 03 的源报告转储里取**定性文本**（有文字、但未必有结构化数字）。

    为什么需要：06 此前只读事实库，而事实库只抽**带单位的数值**——
    于是「项目概况」这类**有大量文字材料**的节被判成「无材料，需出缺项」，
    报告里最该有内容的节反而是空的（实测事实覆盖 1/23 节）。

    转储文件（*.blocks.txt）按 `[Heading 2] <节标题>` 分节，
    节标题与模板节标题**逐字一致**，故按标题匹配即可。
    """
    out: dict[str, list[dict]] = {}
    if not dump_dir.is_dir():
        return out
    for path in sorted(dump_dir.glob("*.blocks.txt")):
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except OSError:
            continue
        cur = ""
        for ln in lines:
            m = re.match(r"^\[Heading (\d)\]\s*(.*?)\s*$", ln.strip())
            if m:
                cur = m.group(2).strip()
                continue
            if not cur:
                continue
            s = ln.strip()
            if not s:
                continue
            # 跳过结构标记与表头/数据行（那些进事实库）
            if s.startswith("[") and "]" in s[:20]:
                style = s[1:s.index("]")] if "]" in s else ""
                text = s[s.index("]") + 1:].strip() if "]" in s else ""
            else:
                style, text = "", s
            if not text or text.startswith("【缺口】"):
                continue
            if text.startswith("表（") and ("表头：" in text or "数据行：" in text):
                continue
            # OCR 来源的**裸数字**：口径不明，不得当数值引用。
            # 实测项目信息.png 的 OCR 文本里出现「20000 12000」，
            # 它们在图片上是工程量/面积一类的数，**没有表头说明是投资额**——
            # 若当成投资额写进报告，就是凭空造数。
            # 故凡来自图片 OCR 的数字，一律标「口径不明、不得引用」。
            # 注意判断要看**整行**：行首有「【来源：xxx.png】」前缀，
            # 不能用 startswith("图片文字（") 判断。
            if "图片文字（" in text:
                text = _mark_ocr_numbers(text)
            out.setdefault(cur, []).append({
                "text": text[:400],
                "file": path.name.replace(".blocks.txt", ""),
            })
    return out


# OCR 裸数字（无单位、无表头）→ 标注「口径不明、不得引用」
_OCR_NUM_RE = re.compile(r"(?<![\d.])\d+(?:[.,]\d+)?(?![\d.])")


def _mark_ocr_numbers(text: str) -> str:
    """把 OCR 文本里的数字标出口径不明。

    为什么全标：图片 OCR **丢掉了表格结构**——列名/表头与数据的对应关系没了，
    数字在版面上挨着什么就是什么。实测项目信息.png 的 OCR 文本里
    「20000 12000」与「一、中央预算内投资」相邻，极易被误读成投资额，
    而它们在图片上其实是工程量/面积一类的数，**没有任何表头说明是投资额**。
    故凡来自图片 OCR 的数字一律标「口径不明、不得引用」。

    保留原文摘录的价值（定性描述仍可用），只是把数值圈出来。
    """
    def repl(m):
        return f"〔{m.group(0)}·口径不明不得引用〕"
    return _OCR_NUM_RE.sub(repl, text)


def main(argv=None):
    ap = argparse.ArgumentParser(description="生成章节写作任务包（大模型成文输入）")
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--out", default="")
    ap.add_argument("--chapters", default="", help="只生成指定章，如 A,B,C")
    args = ap.parse_args(argv)

    project_dir = Path(args.project_dir).resolve()
    out_dir = Path(args.out) if args.out else project_dir / "过程文件" / "06-分章成文" / "写作任务"
    chapters = [c.strip() for c in args.chapters.split(",") if c.strip()] or None
    payload = build_tasks(project_dir, out_dir, chapters)
    print(f"[写作任务] {len(payload['chapters'])} 章 → {out_dir}")
    print(f"[写作任务] 提示词：{out_dir / '写作任务.md'}")
    print(f"[写作任务] 回填：{out_dir / '分章初稿_待回填'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
