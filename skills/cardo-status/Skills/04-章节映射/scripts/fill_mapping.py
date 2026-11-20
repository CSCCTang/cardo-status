#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第 04 步：在映射骨架上填「模板节点 → 原始报告标题」，并输出置信度。

规则（fail-closed）：
* 标题归一化后完全一致 / 互相包含 → high
* 命中领域同义词表（可研常见章节名）→ medium，是「提案」
* 其余 → gap（不猜；gap 是合法结论，不阻塞）
* 人工确认文件（人工确认.yaml）优先：
    nodes[].sources 非空 → high（人工确认落位）
    nodes[].sources 为空 → gap（人工复核判定不作映射）

medium 提案的处置（默认自动采纳）：
  同义词表给出的映射命中面窄、可核对，为它逐条问用户「保留还是判缺项」
  是纯粹的往返消耗。因此默认**直接采纳**并写进 映射统计.json，
  用户想改判随时写 人工确认.yaml 后 --from 04 重跑。
  需要旧行为（停下来逐条复核）时加 --strict-review。

产物：过程文件/04-章节映射/章节映射.yaml + 映射统计.json
退出码：0 全部可落位（含自动采纳的 medium 提案）；3 仅在 --strict-review 下有 medium 提案待复核
"""
import argparse
import json
import os
import re
import sys

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None

ALIASES = {
    "项目调研范围与期限": ["项目概况"],
    "区位条件": ["项目选址", "项目建设条件"],
    "规划要求": ["规划政策符合性"],
    "历史沿革分析": ["项目建设背景"],
    "土地与空间利用现状": ["项目选址", "建设内容和规模"],
    "建筑现状与风貌保护现状": ["老旧街区改造方案", "工程方案", "建设内容和规模"],
    "人口与社会人居现状": ["需求分析"],
    "街区业态与经营现状": ["需求分析", "建设内容和规模"],
    "交通与通行现状": ["建设方案", "技术方案", "工程方案"],
    "市政基础设施现状": ["建设方案", "技术方案", "工程方案"],
    "公共服务配套现状": ["需求分析", "建设内容和规模"],
    "生态与人居环境现状": ["项目建设条件", "生态环境影响分析"],
}

# 汇编产物里的缺口标记（assemble_materials.py 写入）
GAP_MARK = "【缺口】"


def load_coverage(path: str) -> dict:
    """读汇编覆盖度清单：{标题: {hasMaterial, blockCount, chars}}。

    没有这份清单时（例如源报告是用户自备的成篇报告），返回空 dict，
    判定退化为「按标题匹配」的原行为。
    """
    if not path or not os.path.exists(path):
        return {}
    try:
        data = json.load(open(path, encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    out = {}
    for sec in data.get("sections") or []:
        title = str(sec.get("title") or "").strip()
        if title:
            out[norm(title)] = sec
    return out


def norm(t):
    t = re.sub(r"\s+", "", str(t or ""))
    t = re.sub(r"^(第[一二三四五六七八九十0-9]+[章节部分]|[0-9.]+[、.]?|"
               r"[（(][0-9一二三四五六七八九十]+[)）])", "", t)
    return re.sub(r"[：:、，,。.（）()《》“”\"'—\-]", "", t)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--skeleton", required=True)
    ap.add_argument("--source-outline", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--confirm", required=False, default="")
    ap.add_argument("--direction", default="老旧街区项目")
    ap.add_argument("--coverage", default="",
                    help="汇编覆盖度清单（assemble_materials 产出）；用于判定无料节为 gap")
    ap.add_argument("--strict-review", action="store_true",
                    help="中等置信映射改为停下来等人工复核（默认直接采纳，不打断）")
    args = ap.parse_args(argv)

    skeleton = yaml.safe_load(open(args.skeleton, encoding="utf-8"))
    outline = json.load(open(args.source_outline, encoding="utf-8"))
    headings = [x["text"] for x in outline
                if x.get("kind") == "heading" and x.get("level", 9) <= 2]
    hnorm = [(h, norm(h)) for h in headings]
    coverage = load_coverage(args.coverage)
    confirm = {}
    stale = []
    if args.confirm and os.path.exists(args.confirm):
        data = yaml.safe_load(open(args.confirm, encoding="utf-8")) or {}
        # 人工确认可能来自上一轮（换过源报告形态：成篇源报告 ↔ 散材料汇编）。
        # 若某条确认指向的源标题在**当前** source 大纲里已不存在，
        # 该确认已失效——应丢弃并提示，而不是让 04 的自检笼统报 fail。
        src_titles = {norm(h) for h in headings}
        for d in data.get("nodes", []) or []:
            nid = d.get("id")
            srcs = [str(x) for x in (d.get("sources") or []) if str(x).strip()]
            if srcs and not any(norm(x) in src_titles for x in srcs):
                stale.append((nid, srcs))
                continue
            confirm[nid] = d
    if stale:
        for nid, srcs in stale:
            print(f"[04][提示] 人工确认 {nid}（来源：{'、'.join(srcs)}）在当前源报告中已不存在，"
                  f"该条确认已失效，按自动映射处理")

    stats = {"exact": 0, "alias": 0, "gap": 0, "pending_confirm": []}
    gap_nodes = []          # 真实材料缺口（区别于「结构上不需要材料」的节点）
    out_nodes = []
    adopted_medium = []     # 默认采纳的中等置信提案（仅记录，不打断）
    for n in skeleton.get("nodes", []):
        nid, title, level = n["id"], n["title"], n["level"]
        nt = norm(title)
        sources, confidence, note, status = [], "", "", "gap"
        decided = False
        if nid in confirm:
            c = confirm[nid]
            if c.get("sources"):
                sources = [str(s) for s in c["sources"] if str(s).strip()]
                status, confidence = "done", "high"
                note = "人工确认落位" + (f"：{c['note']}" if c.get("note") else "")
                stats["exact"] += 1
                decided = True
            elif "sources" in c:
                status, confidence = "gap", "low"
                note = "人工复核判为 gap" + (f"：{c['note']}" if c.get("note") else "")
                stats["gap"] += 1
                gap_nodes.append({"id": nid, "level": level, "title": title,
                                  "reason": "人工复核判为缺项"})
                decided = True
        if not decided:
            for h, hn in hnorm:
                if hn == nt:
                    sources, confidence, note = [h], "high", f"标题逐字一致（源报告「{h}」）"
                    break
            if not sources:
                cand = [h for h, hn in hnorm if hn and (hn in nt or nt in hn) and abs(len(hn) - len(nt)) <= 4]
                if len(cand) == 1:
                    sources, confidence, note = cand, "high", f"标题互相包含（源报告「{cand[0]}」）"
            if not sources:
                for alias in ALIASES.get(title, []):
                    hit = [h for h, hn in hnorm if norm(alias) in hn]
                    if len(hit) == 1:
                        sources, confidence = hit, "medium"
                        note = f"按领域同义词提议：{title} → 源「{hit[0]}」，待人工复核"
                        break
            if sources:
                # 覆盖度复核：汇编模式下，标题能对上不代表有材料。
                # 若该源章节在覆盖度清单里标记为「无材料」，改判 gap，
                # 否则真实缺口会被标题匹配洗掉，缺口版永远不启用。
                empty = [s for s in sources
                         if coverage.get(norm(s), {}).get("hasMaterial") is False]
                if empty and len(empty) == len(sources):
                    status = "gap"
                    sources = []
                    note = "汇编时该节无材料命中（按缺项处理，不作推断）"
                    stats["gap"] += 1
                    # 结构节点（如「核心问题汇总」「附件」）：内容由流程段/模板外附加节
                    # 生成，不是材料缺口，不计入真实缺口。
                    structural = all(coverage.get(norm(s), {}).get("structural")
                                     for s in empty)
                    if not structural:
                        gap_nodes.append({"id": nid, "level": level, "title": title,
                                          "reason": "材料未命中"})
                else:
                    status = "done"
                    stats["exact" if confidence == "high" else "alias"] += 1
            else:
                status = "gap"
                note = "原始报告无对应章节（按缺项处理，不作推断）"
                stats["gap"] += 1
                gap_nodes.append({"id": nid, "level": level, "title": title,
                                  "reason": "源报告无对应章节"})
        if confidence == "medium":
            stats["pending_confirm"].append(nid)
            adopted_medium.append(nid)
        out_nodes.append({"id": nid, "level": level, "title": title,
                          "sources": sources, "status": status,
                          "confidence": confidence or "low", "notes": note})

    out = dict(skeleton)
    out["nodes"] = out_nodes
    out["direction"] = args.direction
    out["fillTool"] = "fill_mapping.py 1.0.0"

    class IndentDumper(yaml.SafeDumper):
        # 与 make_mapping_skeleton 的输出风格一致：块序列缩进 2 空格，
        # 便于 extract_facts 内置解析器（只认缩进列表）读取。
        def increase_indent(self, flow=False, indentless=False):
            return super().increase_indent(flow, False)

    with open(args.out, "w", encoding="utf-8") as f:
        yaml.dump(out, f, Dumper=IndentDumper, allow_unicode=True,
                  sort_keys=False, width=200, default_flow_style=False)

    print(f"[04] 映射已填：{len(out_nodes)} 节点；high={stats['exact']} "
          f"medium={stats['alias']} gap={stats['gap']}")
    for n in out_nodes:
        if n["status"] == "gap":
            print(f"   [gap] {n['id']} {n['title']}")
    stats_path = os.path.join(os.path.dirname(args.out), "映射统计.json")
    with open(stats_path, "w", encoding="utf-8") as f:
        json.dump({"schema": "cardo/mapping-stats/v1", **stats,
                   "coverageApplied": bool(coverage),
                   "coverageEmptySections": [t for t, s in coverage.items()
                                             if s.get("hasMaterial") is False],
                   # 真实材料缺口（已剔除结构节点）：缺口版启用与否的依据
                   "materialGaps": gap_nodes,
                   "materialGapCount": len(gap_nodes),
                   "mediumNodes": [n["id"] for n in out_nodes if n["confidence"] == "medium"],
                   # 默认采纳的中等置信提案：已在映射里落位，不阻断流程；
                   # 用户要复核可写 人工确认.yaml 后 --from 04 重跑。
                   "adoptedMedium": adopted_medium},
                  f, ensure_ascii=False, indent=1)

    if stats["pending_confirm"] and args.strict_review:
        print("[04][GATE] 中等置信映射仅为提案，需人工复核：")
        for n in out_nodes:
            if n["id"] in stats["pending_confirm"]:
                print(f"   [review] {n['id']} {n['title']} → 源「{'、'.join(n['sources'])}」"
                      f"（{n['notes']}）")
        print("   在 过程文件/04-章节映射/人工确认.yaml 写明 nodes[].id + sources："
              "保留该映射则填原 sources，判为缺项则写 sources: []")
        return 3

    if stats["pending_confirm"]:
        # 默认路径：中等置信提案直接采纳，不打断。
        # 理由：同义词表（ALIASES）给出的映射命中面窄、可核对，
        # 而为此让用户逐条答「保留还是判缺项」是纯粹的往返消耗；
        # 提案本身已写进 映射统计.json，用户想改随时能改。
        print(f"[04] 中等置信提案 {len(adopted_medium)} 条已自动采纳"
              "（可在 人工确认.yaml 里改判后 --from 04 重跑）：")
        for n in out_nodes:
            if n["id"] in stats["pending_confirm"]:
                print(f"   [adopted] {n['id']} {n['title']} → 源「{'、'.join(n['sources'])}」")
    return 0


if __name__ == "__main__":
    sys.exit(main())
