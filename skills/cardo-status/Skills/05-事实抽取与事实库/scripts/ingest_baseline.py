#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ingest_baseline.py -- 把包内既有成果（OSGB 现状指标 / AHP 结论）接进现状分析。

盘点的三个断点（2026-11-14）：
  1. 03 的转储白名单不收 .json，OSGB《现状指标.json》落到 10_低空指标 后无人消费；
  2. 现状分析不读 AHP 结论，跨包引用各说各话；
  3. 07 只算表 3.1，OSGB 已算好的绿地率/建筑密度/容积率/建筑高度没有入口。

本脚本只做「接入 + 核对」，不重算 OSGB 的任何指标：
  * 按 JSON 结构识别现状指标（metrics/scope/sensitivity/texture/diagnosis）
    与 AHP 权重（criteria/consistency）、AHP 评分（ranking/schemes）；
  * 原值照搬，登记来源相对路径 + sha256（同源可核对）；
  * 渲染成带单位的《包内依据.md》，供正文锚点「内部依据：…」引用；
  * 核对 AHP 权重和、CR 判定、排序单调，写入同源核对.json。

用法：
  python ingest_baseline.py --project-dir <工作区> [--out <过程文件/05…/包内依据>]
退出码：0 成功（无包内依据也返回 0，写空清单）；2 输入错。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

SCHEMA = "cardo/status-baseline/v1"

# 上游成果整包形态：成果文件/成品版/内部件 + 目录名（AHP/OSGB）+ 件名 + 内部结构。
UPSTREAM_PKG_HINTS = ("成果文件", "成品版", "内部件")
# 规定该取的件：AHP = 指标权重与评分依据；OSGB = 量化指标与体检依据。
CLAIM_BY_NAME = {
    "权重表.json": ("ahp-weights", "权重表"),
    "评分表.json": ("ahp-scores", "评分表"),
    "现状指标.json": ("osgb", "现状指标"),
    "扩展指标.json": ("osgb-extended", "扩展指标"),
}


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def load_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return None


def pct(x) -> str:
    return f"{float(x) * 100:.2f}%"


def is_osgb(d: dict) -> bool:
    return isinstance(d.get("metrics"), list) and any(
        isinstance(m, dict) and m.get("id") and "value" in m for m in d["metrics"])


def is_ahp_weights(d: dict) -> bool:
    return isinstance(d.get("criteria"), list) and isinstance(d.get("consistency"), dict)


def is_ahp_scores(d: dict) -> bool:
    return isinstance(d.get("ranking"), list) and isinstance(d.get("schemes"), list)


def scan(project_dir: Path) -> list[dict]:
    """扫描输入文件里的上游成果包，按「整包 + 件名 + 内部结构」认领。

    角色口径（2026-11-14）：
      * AHP 成果 = 指标权重与评分依据 → 只取 权重表.json、评分表.json；
      * OSGB 成果 = 量化指标与体检依据 → 只取 现状指标.json、扩展指标.json；
      * 包内其余中间件与模型资产（矩阵集、群体矩阵、拓扑清理、mesh-repair、
        pc-classify、roof-segment…）登记「已见、不入正文」，不撒进材料分类。

    返回记录：{path, data, kind, item, inBody, reason}
      kind: osgb / osgb-extended / ahp-weights / ahp-scores / seen-only
    """
    out: list[dict] = []
    root = project_dir / "输入文件"
    if not root.is_dir():
        return out
    for p in sorted(root.rglob("*.json")):
        try:
            if p.stat().st_size > 20 * 1024 * 1024:
                continue
        except OSError:
            continue
        rel = p.relative_to(root)
        in_pkg = any(h in part for part in rel.parts for h in UPSTREAM_PKG_HINTS) \
            or any(k in rel.as_posix().lower() for k in ("ahp", "osgb"))
        claim = CLAIM_BY_NAME.get(p.name)
        if not in_pkg and not claim:
            continue
        d = load_json(p)
        if not isinstance(d, dict):
            if in_pkg:
                out.append({"path": p, "data": {}, "kind": "seen-only",
                            "item": "中间件", "inBody": False,
                            "reason": "非 JSON 对象（模型资产/缓存），已见、不入正文"})
            continue
        if claim:
            kind, item = claim
            if kind == "ahp-weights" and not is_ahp_weights(d):
                out.append({"path": p, "data": d, "kind": "seen-only", "item": item,
                            "inBody": False, "reason": "结构不含 criteria/consistency，非权重表"})
                continue
            if kind == "ahp-scores" and not is_ahp_scores(d):
                out.append({"path": p, "data": d, "kind": "seen-only", "item": item,
                            "inBody": False, "reason": "结构不含 ranking/schemes，非评分表"})
                continue
            if kind in ("osgb", "osgb-extended") and not is_osgb(d):
                out.append({"path": p, "data": d, "kind": "seen-only", "item": item,
                            "inBody": False, "reason": "结构不含 metrics 列表，非指标件"})
                continue
            out.append({"path": p, "data": d, "kind": kind, "item": item,
                        "inBody": True, "reason": ""})
            continue
        out.append({"path": p, "data": d, "kind": "seen-only",
                    "item": item_name_from_path(p.name), "inBody": False,
                    "reason": "包内中间件/模型资产，已见、不入正文"})
    # 上游包内的说明/依据件（指标说明、体检报告、AHP 结论、依据对照表、权重/评分表）：
    # 认领进依据清单（只登记件与指纹，不解析正文），中间件继续留在「已见、不入正文」。
    for p in sorted(root.rglob("*")):
        if p.suffix.lower() not in (".docx", ".pdf", ".xlsx", ".xls", ".csv"):
            continue
        try:
            rel = p.relative_to(root)
        except ValueError:
            continue
        if not any(h in part for part in rel.parts for h in UPSTREAM_PKG_HINTS):
            continue
        item = item_name_from_path(p.name)
        if item == "中间件":
            if "权重表" in p.name:
                item = "权重表"
            elif "评分表" in p.name:
                item = "评分表"
            else:
                out.append({"path": p, "data": {}, "kind": "seen-only",
                            "item": "中间件", "inBody": False,
                            "reason": "包内中间件/模型资产，已见、不入正文"})
                continue
        out.append({"path": p, "data": {}, "kind": "doc", "item": item,
                    "inBody": True, "reason": ""})
    return out


def item_name_from_path(name: str) -> str:
    """给包内件一个可读件名（规定取件进正文，其余只登记）。"""
    if "依据对照表" in name:
        return "依据对照表"
    if "权重与评分结论" in name or "方案评分结论" in name:
        return "AHP结论"
    for key in ("指标说明", "体检报告", "权重表", "评分表",
                "现状指标", "扩展指标", "矩阵集", "群体矩阵",
                "场地要素", "三维要素", "拓扑清理"):
        if key in name:
            return key
    return "中间件"


def norm_metrics(d: dict) -> dict:
    metrics = []
    for m in d.get("metrics") or []:
        if not isinstance(m, dict) or not m.get("id"):
            continue
        metrics.append({k: m.get(k) for k in
                        ("id", "name", "value", "range", "unit", "confidence",
                         "method", "caliber", "limits", "sensitivity")})
    return {"metrics": metrics, "scope": d.get("scope") or {},
            "sensitivity": d.get("sensitivity") or {},
            "texture": d.get("texture") or {},
            "diagnosis": d.get("diagnosis") or {},
            # 逐建筑分量与场地高程：空间分析（基底集中度、院落围合、台地高差）用原始聚合值
            "diagnostics": d.get("diagnostics") or {},
            "generatedAt": d.get("generatedAt") or ""}


def norm_ahp(w: dict, s: dict | None) -> dict:
    out = {"weights": None, "scores": None, "checks": {}}
    if w:
        out["weights"] = {
            "criteria": [{"id": c.get("id"), "name": c.get("name"),
                          "weight": c.get("weight")} for c in (w.get("criteria") or [])],
            # 准则内指标的局部权重：现状章节引用「哪几项指标在现状准则里更重」时需要
            "indicators": [{"id": i.get("id"), "name": i.get("name"),
                            "criterionId": i.get("criterionId"),
                            "localWeight": i.get("localWeight"),
                            "globalWeight": i.get("globalWeight")}
                           for i in (w.get("indicators") or [])],
            "consistency": w.get("consistency") or {},
            "methods": w.get("methods") or {},
        }
        ws = [float(c.get("weight") or 0) for c in (w.get("criteria") or [])]
        out["checks"]["weightsSum"] = round(sum(ws), 9)
        out["checks"]["weightsSumOk"] = abs(sum(ws) - 1.0) <= 1e-6
    if s:
        ratings = {}
        for r in s.get("ranking") or []:
            ratings[r.get("name")] = (r.get("rating") or {}).get("label")
        out["scores"] = {
            "demo": bool(s.get("demo")),
            "ranking": [{"order": r.get("order"), "name": r.get("name"),
                         "composite": r.get("composite"),
                         "label": (r.get("rating") or {}).get("label")}
                        for r in (s.get("ranking") or [])],
            "schemes": [{"id": x.get("id"), "name": x.get("name"),
                         "composite": x.get("composite")} for x in (s.get("schemes") or [])],
            # 各方案在四项准则上的加权得分：现状章节要与「现状情况」准则的判断逻辑对齐
            "criterionScores": {},
        }
        for scheme in (s.get("schemes") or []):
            agg: dict[str, list[float]] = {}
            for it in (scheme.get("scores") or []):
                cid = it.get("criterionId")
                gw = float(it.get("globalWeight") or 0)
                if not cid or gw <= 0:
                    continue
                pair = agg.setdefault(cid, [0.0, 0.0])
                pair[0] += float(it.get("score") or 0) * gw
                pair[1] += gw
            out["scores"]["criterionScores"][scheme.get("name")] = {
                k: round(v[0] / v[1], 2) for k, v in agg.items() if v[1]}
        comps = [float(r.get("composite") or 0) for r in (s.get("ranking") or [])]
        out["checks"]["rankingDesc"] = all(
            comps[i] >= comps[i + 1] for i in range(len(comps) - 1))
    return out


def norm_extended(d: dict) -> dict:
    """OSGB 扩展指标：只登记件名与原值，不改写主指标口径。"""
    metrics = []
    for m in d.get("metrics") or []:
        if not isinstance(m, dict) or not m.get("id"):
            continue
        metrics.append({"id": m.get("id"), "name": m.get("name"),
                        "value": m.get("value"), "unit": m.get("unit"),
                        "caliber": m.get("caliber")})
    return {"metrics": metrics, "count": len(metrics)}


def render_md(osgb: dict, ahp: dict, sources: list[dict],
              extended: dict | None = None, seen: list[dict] | None = None) -> str:
    """渲染带单位的依据文本：正文锚点「内部依据：<本文件>」靠它提供可核数值。"""
    L = ["# 包内依据（OSGB 现状指标 / AHP 结论，原值照搬）", ""]
    if sources:
        L.append("## 认领的上游成果件（只取规定该取的件）")
        for s in sources:
            tag = f"{s.get('item')}（{s.get('kind')}）" if s.get("item") else s.get("kind")
            L.append(f"- {tag}：{s['file']}（sha256 {s['sha256'][:16]}…）")
        L.append("")

    if osgb:
        scope = osgb.get("scope") or {}
        extent = scope.get("extentM") or {}
        L.append("## OSGB 现状指标（实测值，不再重算）")
        L.append(f"- 测绘范围面积：{float(scope.get('areaM2') or 0):.2f} 平方米"
                 f"（口径 {scope.get('areaMethod') or 'bbox_footprint'}，{scope.get('crs') or ''}）")
        if extent:
            L.append(f"- 范围尺度：东西 {float(extent.get('x') or 0):.2f} 米 × "
                     f"南北 {float(extent.get('y') or 0):.2f} 米")
        if scope.get("groundZ") is not None:
            L.append(f"- 场地地面高程：{float(scope['groundZ']):.2f} 米")
        if scope.get("validCells") is not None:
            L.append(f"- 有效栅格：{scope.get('validCells')} 个（掩膜 {scope.get('maskCells')} 个）")
        for m in osgb.get("metrics") or []:
            mid = m.get("id") or ""
            unit = m.get("unit") or ""
            val = m.get("value")
            rg = m.get("range") or []
            if mid in ("building_coverage_ratio", "greenery_ratio", "road_ratio"):
                sval = pct(val)
                srg = "～".join(pct(x) for x in rg)
            elif mid == "floor_area_ratio_est":
                sval = f"{float(val):.4f}"
                srg = "～".join(f"{float(x):.4f}" for x in rg)
            elif unit == "m":
                sval = f"{float(val):.2f} 米"
                srg = "～".join(f"{float(x):.2f} 米" for x in rg)
            elif unit == "栋":
                sval = f"{float(val):.0f} 栋"
                srg = "～".join(f"{float(x):.0f} 栋" for x in rg)
            elif unit in ("", "1"):
                sval = f"{val}"
                srg = "～".join(str(x) for x in rg)
            else:
                sval = f"{val}{unit}"
                srg = "～".join(str(x) for x in rg)
            conf = m.get("confidence")
            line = f"- {m.get('name')}：{sval}"
            if srg:
                line += f"（区间 {srg}）"
            if conf is not None:
                try:
                    line += f"，置信度 {float(conf):.2f}"
                except (TypeError, ValueError):
                    line += f"，置信度 {conf}"
            if m.get("method"):
                line += f"，口径：{m['method']}"
            if m.get("caliber"):
                line += f"，{m['caliber']}"
            L.append(line)
            # limits 可能是数组，也可能是一整段字符串；按字符遍历会把句子拆碎
            lims = m.get("limits") or []
            if isinstance(lims, str):
                lims = [s.strip() for s in re.split(r"[；;\n]", lims) if s.strip()]
            for lim in list(lims)[:3]:
                L.append(f"  - 限制：{lim}")
        sens = osgb.get("sensitivity") or []
        if isinstance(sens, list) and sens:
            L.append("### 网格敏感性（同一指标在三种网格下的复算）")
            for row in sens:
                if not isinstance(row, dict):
                    continue
                cell = row.get("cellM")
                bits = [f"网格 {cell} 米"]
                if row.get("coverage") is not None:
                    bits.append(f"建筑覆盖率 {pct(row['coverage'])}")
                if row.get("buildingCount") is not None:
                    bits.append(f"建筑栋数 {row['buildingCount']} 栋")
                if row.get("heightP50") is not None:
                    bits.append(f"高度中位数 {float(row['heightP50']):.2f} 米")
                if row.get("heightP90") is not None:
                    bits.append(f"高度 P90 {float(row['heightP90']):.2f} 米")
                if row.get("floorAreaRatio") is not None:
                    bits.append(f"容积率 {float(row['floorAreaRatio']):.4f}")
                L.append("- " + "，".join(bits))
        tex = osgb.get("texture") or {}
        if tex:
            L.append("### 纹理抽样（绿化/道路占比的代理口径）")
            for k, v in tex.items():
                if not isinstance(v, (int, float, str)):
                    continue
                if isinstance(v, float) and ("ratio" in k or "share" in k or "coverage" in k):
                    L.append(f"- {k}：{pct(v)}")
                else:
                    L.append(f"- {k}：{v}")
        diag = osgb.get("diagnosis") or {}
        if diag:
            L.append("### 诊断（跨格聚合参考值）")
            for k, v in diag.items():
                if not isinstance(v, (int, float, str)):
                    continue
                if isinstance(v, float) and ("ratio" in k or "share" in k or "coverage" in k):
                    L.append(f"- {k}：{pct(v)}")
                else:
                    L.append(f"- {k}：{v}")
            g = diag.get("global") or {}
            for k, v in g.items():
                if not isinstance(v, (int, float, str)):
                    continue
                if isinstance(v, float) and ("ratio" in k or "share" in k or "coverage" in k):
                    L.append(f"- 全网格 {k}：{pct(v)}")
                else:
                    L.append(f"- 全网格 {k}：{v}")
        # 逐建筑分量与场地高程：建筑基底集中度、院落围合、台地高差的原值出处
        dg = osgb.get("diagnostics") or {}
        comp = dg.get("components") or {}
        mask = dg.get("mask") or {}
        z = dg.get("cellZ") or {}
        if comp or z:
            L.append("### 诊断：建筑分量与场地高程（1 米网格原值）")
        if comp:
            L.append(f"- 识别分量 {comp.get('total')} 个；闭运算后栅格 {mask.get('closedCells')} 个；"
                     f"最大分量 {comp.get('largestCells')} 格，占比 {pct(comp.get('largestShare'))}；"
                     f"小于 5 格的分量 {comp.get('lt5Cells')} 个")
        if z:
            L.append(f"- 场地格高程：P5 {float(z.get('p5')):.2f} 米，P50 {float(z.get('p50')):.2f} 米")
        L.append("")

    if extended and extended.get("metrics"):
        L.append("## OSGB 扩展指标（原值照搬，作补充依据）")
        for m in extended["metrics"]:
            val = m.get("value")
            unit = m.get("unit") or ""
            line = f"- {m.get('name') or m.get('id')}：{val}{unit}"
            if m.get("caliber"):
                line += f"（口径：{m['caliber']}）"
            L.append(line)
        L.append("")

    if ahp.get("weights"):
        L.append("## AHP 结论（定稿值，直接引用）")
        for c in ahp["weights"]["criteria"]:
            L.append(f"- 准则权重 {c.get('name')}：{float(c.get('weight') or 0):.6f}"
                     f"（{pct(c.get('weight') or 0)}）")
        cons = ahp["weights"].get("consistency") or {}
        if cons.get("threshold") is not None:
            L.append(f"- 一致性阈值：CR<{cons.get('threshold')}，通过 "
                     f"{cons.get('passCount')}/{cons.get('total')} 个矩阵")
        for m in cons.get("matrices") or []:
            L.append(f"- 矩阵 {m.get('id')}：λmax={m.get('lambdaMax')}，CR={m.get('cr')}")
        inds = [i for i in (ahp["weights"].get("indicators") or [])
                if i.get("criterionId") == "current"]
        if inds:
            L.append("- 现状情况准则内指标局部权重：" + "、".join(
                f"{i.get('name')} {pct(i.get('localWeight') or 0)}" for i in inds))
        L.append("")
    if ahp.get("scores"):
        L.append("### 方案评分（AHP 定稿）")
        for r in ahp["scores"]["ranking"]:
            L.append(f"- 第 {r.get('order')} 名 {r.get('name')}："
                     f"{r.get('composite')}（{r.get('label')}）")
        cs = ahp["scores"].get("criterionScores") or {}
        if cs:
            L.append("### 准则得分（方案在各准则上的加权得分）")
            for crit in (ahp["weights"] or {}).get("criteria") or []:
                parts = [f"{n} {v[crit['id']]:.2f} 分" for n, v in cs.items()
                         if crit.get("id") in v]
                if parts:
                    L.append(f"- 准则得分 {crit.get('name')}：" + "、".join(parts))
        L.append("")
    if seen:
        L.append("## 上游包内其余件（已见、不入正文）")
        L.append("")
        L.append("| 件 | 文件 | 处置 |")
        L.append("|---|---|---|")
        for s in seen:
            L.append(f"| {s.get('item')} | {s['file']} | {s.get('reason')} |")
        L.append("")
    return "\n".join(L) + "\n"


def main() -> int:
    ap = argparse.ArgumentParser(description="包内依据（OSGB/AHP）接入")
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--out", default="")
    args = ap.parse_args()
    ws = Path(args.project_dir)
    if not ws.is_dir():
        print(f"[AHP-E-INPUT] 工作区不存在：{ws}", file=sys.stderr)
        return 2
    out = Path(args.out) if args.out else (ws / "过程文件" / "05-事实抽取与事实库" / "包内依据")
    out.mkdir(parents=True, exist_ok=True)

    found = scan(ws)
    osgb, extended, ahp_w, ahp_s = None, None, None, None
    sources: list[dict] = []
    claimed: list[dict] = []
    seen: list[dict] = []
    for rec in found:
        p, d, kind = rec["path"], rec["data"], rec["kind"]
        rel = p.relative_to(ws).as_posix()
        if rec.get("inBody"):
            digest = sha256_file(p)
            claimed.append({"item": rec["item"], "kind": kind, "file": rel,
                            "sha256": digest, "inBody": True})
            sources.append({"kind": kind, "item": rec["item"], "file": rel,
                            "sha256": digest, "schema": d.get("schema") or ""})
        else:
            seen.append({"item": rec.get("item") or "中间件", "file": rel,
                         "inBody": False, "reason": rec.get("reason") or "已见、不入正文"})
        if kind == "osgb" and osgb is None:
            osgb = norm_metrics(d)
        elif kind == "osgb-extended" and extended is None:
            extended = norm_extended(d)
        elif kind == "ahp-weights" and ahp_w is None:
            ahp_w = d
        elif kind == "ahp-scores" and ahp_s is None:
            ahp_s = d
    ahp = norm_ahp(ahp_w, ahp_s) if (ahp_w or ahp_s) else {}

    upstream = {"claimed": claimed, "seenOnly": seen}
    data = {"schema": SCHEMA, "sources": sources, "upstream": upstream,
            "osgb": osgb or {}, "osgbExtended": extended or {}, "ahp": ahp,
            "mdPath": (out / "包内依据.md").as_posix()}
    (out / "包内依据.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    (out / "包内依据.md").write_text(
        render_md(osgb or {}, ahp, sources, extended or {}, seen), encoding="utf-8")
    (out / "上游成果认领.json").write_text(
        json.dumps({"schema": "cardo/upstream-claim/v1", **upstream},
                   ensure_ascii=False, indent=1) + "\n", encoding="utf-8")

    # 同源核对：跨包引用的一致性证据，不是重算
    rel_json = (out / "包内依据.json").relative_to(ws).as_posix()
    checks = {
        "schema": "cardo/status-baseline-checks/v1",
        "sources": sources,
        "osgbMetricCount": len((osgb or {}).get("metrics") or []),
        "osgbExtendedCount": len((extended or {}).get("metrics") or []),
        "ahpWeightsSum": (ahp.get("checks") or {}).get("weightsSum"),
        "ahpWeightsSumOk": (ahp.get("checks") or {}).get("weightsSumOk"),
        "ahpRankingDesc": (ahp.get("checks") or {}).get("rankingDesc"),
        "ahpDemo": ((ahp.get("scores") or {}).get("demo")),
        "upstreamClaimed": len(claimed),
        "upstreamSeenOnly": len(seen),
    }
    (out / "同源核对.json").write_text(
        json.dumps(checks, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")

    ok = bool(osgb) or bool(ahp)
    print(f"[baseline] {'接入成功' if ok else '未发现 OSGB/AHP JSON（按普通材料出稿）'}；"
          f"认领 {checks['upstreamClaimed']} 件、已见不入正文 {checks['upstreamSeenOnly']} 件；"
          f"OSGB 指标 {checks['osgbMetricCount']} 项、扩展 {checks['osgbExtendedCount']} 项；"
          f"AHP 权重和 {checks['ahpWeightsSum']}；输出 {rel_json}")
    return 0


if __name__ == "__main__":
    setup_stdout()
    sys.exit(main())
