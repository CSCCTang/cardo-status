#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_research_tasks.py -- 撰写前生成《检索任务清单》（程序只定契约，不联网）。

它解决的是"到底该去查什么"：把事实库里 missing/conflict 的指标、外加同类项目
必查的高频指标，整理成一份可执行清单——查什么、为什么查、影响哪一章、建议什么
口径、用什么关键词。执行者是 agent；程序不发网络请求，也不生成任何数值。

产物（默认落在 <工作区>/过程文件/00-总控/网络补料/）：
    检索任务清单.md   人读，按 P0→P2 排序
    检索任务.json     机读，供 check_ledger.py 检查 P0 是否闭环

用法：
    python build_research_tasks.py --facts <事实库.json> [--out-dir <网络补料目录>]
                                  [--workspace <工作区>] [--contracts <规则目录>]
"""

from __future__ import annotations

import argparse
import json
import os
import sys

__version__ = "0.1.0"
SCHEMA_TASKS = "cardo/research-tasks/v1"

# 同类项目高频必查指标：不依赖本项目事实库，属于"写这类报告绕不开的外部数据"。
# 关键词一律中性（不带任何具体城市/省份），地域前缀由项目侧 工作区参数.json 的 region 给；
# 未配置 region 时按全国口径出词，不猜地名。
TASK_LIB = [
    {
        "metric": "老旧街区综合整治综合单价", "unit": "元/㎡", "priority": "P0",
        "why": "第一章投资估算与表1费用构成需要外部区间交叉校核（以委托方提供的造价指标为准）",
        "keywords": ["老旧街区 改造 中标 公告 单价",
                     "老旧小区 改造 投资 元/平方米",
                     "城市更新 片区 综合整治 概算"],
    },
    {
        "metric": "户均人口", "unit": "人/户", "priority": "P1",
        "why": "第三章人口与使用强度推算的分母，缺它只能按同类街区经验取值",
        "keywords": ["统计年鉴 户均人口 街道",
                     "七普 家庭户 规模 分区",
                     "城镇 家庭户 平均人数"],
    },
    {
        "metric": "片区容积率", "unit": "—", "priority": "P1",
        "why": "第三章用地与建筑面积口径校核（本项目用地面积已有，容积率为反算校核值）",
        "keywords": ["控制性详细规划 容积率 片区",
                     "老旧街区 容积率 现状",
                     "居住街坊 容积率 标准"],
    },
    {
        "metric": "老旧小区改造单位造价构成", "unit": "元/㎡", "priority": "P2",
        "why": "第一章表1分项占比是否合理，需要外部同类项目的构成比例做旁证",
        "keywords": ["老旧小区 改造 造价 构成 占比",
                     "老旧院落 改造 结算 单价",
                     "城市更新 立面整治 单价 元/平方米"],
    },
]


def locality_prefixes(region: dict | None) -> list[str]:
    """项目侧配置的检索地域前缀（市→省，去重）；未配置返回空列表（全国口径）。"""
    region = region or {}
    out: list[str] = []
    for key in ("city", "province"):
        value = str(region.get(key) or "").strip()
        if value and value not in out:
            out.append(value)
    return out


def localize(keywords: list[str], prefixes: list[str]) -> list[str]:
    """中性关键词 → 本地优先、再全国：每个词前挂地域前缀，末尾保留原中性词。"""
    out: list[str] = []
    for kw in keywords:
        out.extend(f"{p} {kw}" for p in prefixes)
        out.append(kw)
    return out



def _load_json(path: str) -> dict:
    with open(path, encoding="utf-8-sig") as fh:
        return json.load(fh)


def collect_tasks(facts: list[dict], prefixes: list[str] | None = None) -> list[dict]:
    """事实库缺口 → P0 任务；高频指标缺失 → 库内任务。"""
    prefixes = prefixes or []
    tasks: list[dict] = []
    seen: set[str] = set()
    for f in facts:
        status = str(f.get("status") or "")
        if status not in ("missing", "conflict"):
            continue
        name = str(f.get("name") or f.get("metric") or f.get("id") or "").strip()
        if not name or name in seen:
            continue
        seen.add(name)
        tasks.append({
            "metric": name,
            "unit": str(f.get("unit") or ""),
            "priority": "P0",
            "why": (f"事实库 {f.get('id')}（{status}）未闭合，"
                    f"影响章节：{f.get('section') or '未标注'}"),
            "keywords": [f"{name} {p}" for p in prefixes] + [f"{name} 公开 数据"],
        })
    for item in TASK_LIB:
        if any(item["metric"] in t["metric"] or t["metric"] in item["metric"] for t in tasks):
            continue
        task = dict(item)
        task["keywords"] = localize(item["keywords"], prefixes)
        tasks.append(task)
    for i, t in enumerate(tasks, 1):
        t["taskId"] = f"T-{i:03d}"
    tasks.sort(key=lambda t: (t["priority"], t["taskId"]))
    for i, t in enumerate(tasks, 1):
        t["taskId"] = f"T-{i:03d}"
    return tasks


def render_md(tasks: list[dict], workspace: str, prefixes: list[str] | None = None) -> str:
    prefixes = prefixes or []
    scope = "/".join(prefixes) + " → 全国" if prefixes else "全国口径"
    lines = [
        "# 检索任务清单", "",
        "> 本清单由程序按事实库缺口生成，**程序不联网**。检索由 agent 执行：",
        "> 有用的页面落盘为清洗后的正文文本（`.txt`），逐条回填《检索台账》（检索台账.json），",
        "> 正文写作时用 `[S-xxx]` 引用。查不到也要登记：`notFound: true` + 检索口径 + 日期。",
        "> 缺 URL / 缺检索日期 / 缺原文摘录 / 落盘文件不存在，条目一律无效。", "",
        "## 任务",
        "",
        "| 任务 | 指标 | 优先级 | 为什么（影响章节） | 建议口径 | 检索关键词 | 期望单位 |",
        "|---|---|---|---|---|---|---|",
    ]
    for t in tasks:
        kws = "；".join(t.get("keywords") or [])
        lines.append(f"| {t['taskId']} | {t['metric']} | {t['priority']} | {t.get('why','')} | "
                     f"本地优先：{scope} | {kws} | {t.get('unit') or '—'} |")
    lines += [
        "",
        "## 落盘与回填",
        "",
        "1. 页面正文另存为 `过程文件/00-总控/网络补料/材料/S-0xx_来源机构_YYYY-MM-DD.txt`"
        "（去掉导航、广告、评论与无关段落，只留正文）。",
        "2. 台账 `检索台账.json` 每条填：`id / metric / value / unit / publisher / title / url / "
        "date / quote / scope / file`；`file` 填上面那份 .txt 的相对路径。",
        "3. 查不到：`notFound: true` + `reason` + `quote`（写清检索口径与检索面）+ `date`。",
        "4. P0 任务必须闭环（有有效条目或已查无果），否则 `check_ledger.py` 判失败。",
        "",
        f"（工作区：{workspace or '未指定'}）",
    ]
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="build_research_tasks.py", description="生成检索任务清单")
    ap.add_argument("--facts", default="", help="事实库 json（缺省 <workspace>/过程文件/05-*/facts/事实库.json）")
    ap.add_argument("--workspace", default="", help="工作区根（定位默认输出目录）")
    ap.add_argument("--out-dir", default="", help="输出目录（默认 <workspace>/过程文件/00-总控/网络补料）")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args(argv)

    ws = os.path.abspath(args.workspace or os.getcwd())
    out_dir = args.out_dir or ""
    if not out_dir:
        for rel in ("过程文件/00-总控/网络补料", "05_过程/网络补料"):
            cand = os.path.join(ws, *rel.split("/"))
            if os.path.isdir(cand) or rel.startswith("过程文件"):
                out_dir = cand
                break
    facts_path = args.facts
    if not facts_path:
        for root, _dirs, names in os.walk(os.path.join(ws, "过程文件")):
            if "事实库.json" in names:
                facts_path = os.path.join(root, "事实库.json")
                break
    facts: list[dict] = []
    if facts_path and os.path.isfile(facts_path):
        try:
            facts = _load_json(facts_path).get("facts") or []
        except (OSError, ValueError) as exc:
            print(f"WARN: 事实库解析失败（{type(exc).__name__}），只生成通用任务", file=sys.stderr)
    else:
        print("WARN: 未找到事实库，只生成通用任务", file=sys.stderr)

    # 地域前缀是项目侧信息（检索是"本地优先"），工具侧不写死任何城市/省份。
    region: dict = {}
    params_path = os.path.join(ws, "工作区参数.json")
    if os.path.isfile(params_path):
        try:
            region = (_load_json(params_path) or {}).get("region") or {}
        except (OSError, ValueError) as exc:
            print(f"WARN: 工作区参数.json 解析失败（{type(exc).__name__}），检索关键词按全国口径生成",
                  file=sys.stderr)
    prefixes = locality_prefixes(region)
    if not prefixes:
        print("说明：工作区参数.json 未配置 region，检索关键词按全国口径生成。", file=sys.stderr)

    tasks = collect_tasks(facts, prefixes)
    os.makedirs(out_dir, exist_ok=True)
    md_path = os.path.join(out_dir, "检索任务清单.md")
    js_path = os.path.join(out_dir, "检索任务.json")
    with open(md_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(render_md(tasks, ws, prefixes))
    with open(js_path, "w", encoding="utf-8", newline="\n") as fh:
        json.dump({"schema": SCHEMA_TASKS, "workspace": ws, "tasks": tasks},
                  fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    p0 = sum(1 for t in tasks if t["priority"] == "P0")
    print(f"[research] 任务 {len(tasks)} 条（P0 {p0}）→ {md_path}")
    print("[research] 程序不联网：检索由 agent 按清单执行，结果落盘并回填 检索台账.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
