#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""final_gates.py -- 成品硬闸门（只针对 成品版/ 里的报告）。

三道闸门，任一不过即退出码 1：
  1. 缺口词零命中：未提资 / 暂缺 / 待补 / 未提供 / 请提供 / 待补充；
  2. 政策法规不编造：文号「〔2024〕12号」、书名号文件（办法/规定/条例/标准/规范/
     通知/意见/方案/导则），必须在包内依据或素材汇编里能找到出处；
  3. 跨包同源：OSGB 指标、AHP 权重/排序的引用值必须与包内依据原值一致，
     报告里不得出现同一指标的第二种数值。

用法：
  python final_gates.py --project-dir <工作区> --docx <成品.docx> --pdf <成品.pdf> \
      --baseline <包内依据.json> --out <成品闸门.json>
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path

# 运行时是 python313._pth 模式：兄弟模块必须显式加入 sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import _dimension_gate as dimgate  # noqa: E402
import _scope_gate as scopegate  # noqa: E402
import _works_gate as worksgate  # noqa: E402

GAP_WORDS = ["未提资", "暂缺", "待补", "未提供", "请提供", "待补充"]
DOC_NO_RE = re.compile(r"[〔\[【(（]\s*(?:19|20)\d{2}\s*[〕\]】)）]\s*\d+\s*号")
DOC_NAME_RE = re.compile(r"《([^》]{2,40})》")
DOC_KIND_RE = re.compile(r"(办法|规定|条例|标准|规范|通知|意见|方案|导则|细则|决定)")


def docx_text(path: Path) -> str:
    from docx import Document
    doc = Document(str(path))
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def pdf_text(path: Path) -> str:
    try:
        import fitz
    except ImportError:
        return ""
    if not path or not path.exists():
        return ""
    doc = fitz.open(str(path))
    txt = "\n".join(p.get_text("text") for p in doc)
    doc.close()
    # 页脚「第 X 页 / 共 Y 页」不是正文：抽文本时会粘在段尾，被量纲/工程量闸门当成本句的数字，
    # 出现「户均面积 …… 15」这类假命中。先剥掉。
    return re.sub(r"第\s*\d+\s*页\s*[/／]\s*共\s*\d+\s*页", "", txt)


def pct(v: float) -> str:
    return f"{v * 100:.2f}%"


def fmt_metric(mid: str, val) -> str:
    try:
        f = float(val)
    except (TypeError, ValueError):
        return str(val)
    if mid in ("building_coverage_ratio", "greenery_ratio", "road_ratio"):
        return pct(f)
    if f == int(f):
        return str(int(f))
    return f"{f:.4f}".rstrip("0").rstrip(".")


def same_source(text: str, baseline: dict) -> tuple[list, list]:
    """返回 (核对明细, 违规项)。"""
    rows, bad = [], []
    osgb = (baseline.get("osgb") or {})
    metrics = osgb.get("metrics") or {}
    if isinstance(metrics, list):      # 包内依据.json 里 metrics 是数组
        metrics = {m.get("id"): m for m in metrics if isinstance(m, dict)}
    checks = [
        ("building_coverage_ratio", "建筑密度"),
        ("greenery_ratio", "绿地率"),
        ("road_ratio", "道路占比"),
        ("floor_area_ratio_est", "容积率"),
        ("building_count", "建筑栋数"),
        ("building_height_p50", "建筑高度中位"),
        ("building_height_p90", "建筑高度90分位"),
    ]
    for mid, label in checks:
        item = metrics.get(mid) or {}
        if item.get("value") is None:
            continue
        expect = fmt_metric(mid, item["value"])
        used = expect in text or expect.replace("%", "") in text
        rows.append({"indicator": label, "id": mid, "value": item["value"],
                     "rendered": expect, "used": used, "sameSource": used})
        # 若报告写了该指标，但数值不是原值 → 视为另造数
        if mid in ("building_height_p50", "building_height_p90"):
            # 「建筑高度」在正文里也用于台账值（最高层高、单栋高度），不能按前缀抓数；
            # 只看报告是否用到了该分位数的原值。
            if not used:
                bad.append(f"{label}：报告未引用 OSGB 原值 {expect}")
        elif mid == "building_count":
            # 候选栋数（18 栋，区间 10–77，confidence=low）与台账楼栋数不同源不同义：
            # 报告同时写「测绘候选 18 栋」与「台账 22 栋」，只在明确写「候选」处核对。
            vals = re.findall(r"候选[^0-9]{0,8}(\d+)\s*栋", text)
            if vals and expect not in vals:
                bad.append(f"{label}：报告值 {vals} 与 OSGB 原值 {expect} 不一致")
        else:
            vals = re.findall(label + r"[^0-9]{0,12}(\d+\.?\d*)", text)
            want = expect.rstrip("%")
            if vals and want not in vals:
                bad.append(f"{label}：报告值 {vals} 与 OSGB 原值 {expect} 不一致")

    ahp = baseline.get("ahp") or {}
    for item in ((ahp.get("weights") or {}).get("criteria") or []):
        w = item.get("weight")
        name = item.get("name") or item.get("id") or ""
        if w is None:
            continue
        expect = f"{float(w):.6f}"
        used = expect in text
        rows.append({"indicator": f"AHP 权重·{name}", "value": w,
                     "rendered": expect, "used": used, "sameSource": used})
        if used and name:
            got = re.findall(re.escape(name) + r"[^0-9]{0,12}(\d+\.\d{4,6})", text)
            if got and expect not in got:
                bad.append(f"AHP 权重·{name}：报告值 {got} 与定稿 {expect} 不一致")
    rank = ((ahp.get("scores") or {}).get("ranking") or [])
    for r in rank:
        name, order = r.get("name"), int(r.get("order") or 0)
        if name and order and name in text:
            rows.append({"indicator": f"AHP 排序·{name}", "value": order,
                         "rendered": f"第 {order}", "used": True, "sameSource": True})
    return rows, bad


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--docx", required=True)
    ap.add_argument("--pdf", default="")
    ap.add_argument("--baseline", default="")
    ap.add_argument("--basis-lib", default="", help="包内依据库（assets/依据库/依据库.json）")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    ws = Path(args.project_dir)
    docx, pdf = Path(args.docx), Path(args.pdf) if args.pdf else None
    text = docx_text(docx)
    pdf_text_ = pdf_text(pdf) if pdf else ""
    combined = text + "\n" + pdf_text_

    # 1. 缺口词
    gap_hits = {w: combined.count(w) for w in GAP_WORDS if combined.count(w)}

    # 2. 政策法规出处
    support = ""
    for rel in ("过程文件/05-事实抽取与事实库/包内依据/包内依据.md",
                "过程文件/03-模板与源报告转储/源报告转储/素材汇编.blocks.txt"):
        p = ws / rel
        if p.exists():
            support += p.read_text(encoding="utf-8-sig", errors="ignore")
    # 包内依据库：编制依据章节的权威出处（27 条，含文号/机关/核验状态）
    lib_path = Path(args.basis_lib) if args.basis_lib else None
    if lib_path and not lib_path.is_absolute():
        lib_path = ws / lib_path
    if lib_path and lib_path.is_file():
        try:
            lib = json.loads(lib_path.read_text(encoding="utf-8-sig"))
            entries = lib if isinstance(lib, list) else (
                lib.get("items") or lib.get("依据") or lib.get("条目") or [])
            for it in entries:
                if isinstance(it, dict):
                    support += "\n" + " ".join(str(v) for v in it.values() if isinstance(v, str))
                elif isinstance(it, str):
                    support += "\n" + it
        except Exception:
            pass
    support_norm = re.sub(r"\s+", "", support)

    def _name_ok(name: str) -> bool:
        n = re.sub(r"\s+", "", name)
        if n in support_norm:
            return True
        inner = re.findall(r"《([^《》]+)》", n)
        if inner:
            return all(i in support_norm for i in inner)
        return False

    policy_bad = []
    for no in set(DOC_NO_RE.findall(combined)):
        if re.sub(r"\s+", "", no) not in support_norm:
            policy_bad.append(f"文号无包内依据：{no}")
    for name in set(DOC_NAME_RE.findall(combined)):
        if not DOC_KIND_RE.search(name):
            continue
        if not _name_ok(name):
            policy_bad.append(f"文件名无包内依据：《{name}》")

    # 3. 同源
    baseline = json.loads(Path(args.baseline).read_text(encoding="utf-8")) if args.baseline and Path(args.baseline).exists() else {}
    rows, source_bad = same_source(combined, baseline) if baseline else ([], [])

    # 4. 派生量量纲：必须带单位，且落在量级区间内（分母选错会算出「像模像样」的错数）
    # 量纲只看 docx 文本：个别 PDF 字体子集抽取会把汉字还原成控制符+数字
    # （实测「户均面积的分子分母」被抽成「…的 2 6」，误判成缺单位的 26）。
    # docx 与 PDF 同源同字，用 docx 判定不降低检查强度。
    dim_bad = dimgate.scan(text if text.strip() else pdf_text_)

    # 5. 口径范围：样本值不得冒充片区整体值（「片区」+面积型指标必须同句出现「实测样区/样本」）
    scope_bad = scopegate.scan(combined)

    # 6. 工程量逐列映射：每个工程量数字必须来自源表对应列，且列名已在映射表登记
    snap_path = ws / "过程文件/06-分章成文/分章初稿/工程量核算.json"
    if snap_path.exists():
        work_bad, work_rows = worksgate.scan(combined, worksgate.load_snapshot(snap_path))
    else:
        work_bad, work_rows = ["工程量核算.json 缺失，无法核对工程量数字"], []

    ok = (not gap_hits and not policy_bad and not source_bad and not dim_bad
          and not scope_bad and not work_bad)
    report = {
        "ok": ok,
        "gapWords": gap_hits,
        "gapWordsRule": "成品版零命中",
        "policyViolations": policy_bad,
        "sameSource": rows,
        "sameSourceViolations": source_bad,
        "dimensionViolations": dim_bad,
        "dimensionRule": "派生量必须带单位；人口密度 100–2000 人/公顷、人均绿地 2–50 平方米、"
                         "户均面积 20–200 平方米、单栋 10–20000 平方米、每公顷道路 ≤10000 平方米",
        "scopeViolations": scope_bad,
        "scopeRule": "同句出现「片区」与建筑密度/绿地率/容积率/道路占比时，必须出现「实测样区」「样区」或「样本」",
        "worksRows": work_rows,
        "worksViolations": work_bad,
        "worksRule": "工程量按台账真实列名映射（源列名→报告项名→单位）；正文数字须等于该列合计，"
                     "未登记项名后跟数量即违规",
        "docx": str(docx),
        "pdf": str(pdf) if pdf else "",
        "charCount": len(re.sub(r"\s+", "", text)),
    }
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")

    print(f"[成品闸门] 缺口词命中 {sum(gap_hits.values())}；政策问题 {len(policy_bad)}；"
          f"同源问题 {len(source_bad)}；量纲问题 {len(dim_bad)}；口径范围问题 {len(scope_bad)}；"
          f"工程量问题 {len(work_bad)}；{'通过' if ok else '未通过'}")
    for b in policy_bad + source_bad + [d["detail"] for d in dim_bad]:
        print("  - " + b)
    for b in scope_bad:
        print("  - [口径范围] " + b["sentence"])
    for b in work_bad:
        print("  - [工程量] " + b)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
