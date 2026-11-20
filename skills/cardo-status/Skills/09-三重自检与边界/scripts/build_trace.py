#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第 09 步：生成 cardo/trace/v1 成果回溯表（只存引用，不复制正文）。

对成稿正文逐条比对事实库：事实的引文（或数值+单位）在正文出现，才登记一条回溯项。
所有回溯项都带 factId → check_trace 覆盖率 100%。

**匹配口径**（2026-11-01 修）：原先用 `f"{val}{unit}"` 字符串硬匹配，实测命中极少——
事实库存的是 `744.0户`、`6.268万㎡`（浮点 + 符号单位），
正文写的是 `744户`、`6.268万平方米`（整数值 + 中文单位名），
字符串不等 → 73 条事实只命中 1 条，回溯表近乎为空。
故改为三路匹配：
  ① 引文前缀命中（最可靠，优先）
  ② **数值归一 + 单位别名**后命中（744.0 → 744；万㎡ ↔ 万平方米；m ↔ 米）
  ③ **派生值登记**：正文里标了「加总/测算」的数值单独登记为 derived 条目
     —— 派生值本就不在事实库里（事实库存的是分项原值），要求逐字命中是不合理的。
     派生条目**不参与** check_trace 的 factId 覆盖率计算（它们没有 factId），
     但要计入回溯表条数与「正文数字可溯源」的观感指标。
"""
import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _num_match import (  # noqa: E402
    canon_token, canon_unit, derived_tokens, norm_text, numkey, reconcile,
)


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def norm(s):
    return norm_text(s)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--artifact", required=True)
    ap.add_argument("--facts", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--baseline-version", default="1")
    args = ap.parse_args(argv)

    report = open(args.artifact, encoding="utf-8").read()
    nreport = norm(report)
    facts = json.load(open(args.facts, encoding="utf-8")) if os.path.exists(args.facts) else {"facts": []}
    rel_artifact = os.path.relpath(args.artifact, args.project_dir).replace("\\", "/")

    items, seen = [], set()
    for f in facts.get("facts", []):
        if f.get("status") not in ("confirmed", None):
            continue
        val = f.get("value")
        unit = str(f.get("unit") or "")
        quote = norm(f.get("source", {}).get("quote", ""))
        token = norm(f"{val}{unit}") if val is not None else ""
        ctoken = canon_token(val, unit)
        hit = False
        if quote and len(quote) >= 12 and quote[:40] in nreport:
            hit = True                      # ① 引文命中（最可靠）
        elif token and token in nreport:
            hit = True                      # ②a 原样 token
        elif ctoken and ctoken in nreport:
            hit = True                      # ②b 归一 + 单位别名
        if not hit:
            continue
        key = (f.get("id"), str(val), unit)
        if key in seen:
            continue
        seen.add(key)
        items.append({
            "id": f"TR-{len(items) + 1:04d}",
            "locator": f"{f.get('section','')} / {f.get('source',{}).get('locator','')}",
            "factId": f.get("id"),
            "confidence": f.get("confidence", "high"),
            "status": "confirmed",
            "kind": "fact",
        })
        if len(items) >= 80:
            break

    # ③ 派生值登记：正文里标了「加总/测算」的数值单独登记。
    # 它们是按材料算出的量（如 22 栋、1239 户、19820 ㎡），
    # 事实库存的是分项原值（如 11 栋 + 2 栋 + …），逐字命中本就不成立。
    # 这些条目**不带 factId** → 不参与 check_trace 的 factId 覆盖率计算，
    # 但让回溯表如实反映「正文里有多少数字能追到材料」。
    #
    # 状态不再一律 confirmed（2026-11 修）：此前派生值无条件记 confirmed，
    # 于是「分项 19820 却写合计 17420」这类硬伤在 09 自检里全部绿灯。
    # 现在按 reconcile() 的算术对账三态：
    #   conflict   —— 所在行有算术 issue（加总/占比/计提/跨指标冲突），check_trace 判 fail
    #   confirmed  —— 该「值+单位+行」被某条规则算通
    #   unverified —— 格式上无法判定（不冒充通过，计入统计但不阻塞）
    rec = reconcile(report)
    bad_lines = set()
    for issue in rec["issues"]:
        bad_lines.update(issue.get("lines") or [issue.get("line")])
    verified = rec["verified"]

    dseen = {}
    for num, unit, no in derived_tokens(report):
        cnum = numkey(num)
        cunit = canon_unit(unit)
        key = (cnum, cunit)
        if no in bad_lines:
            status = "conflict"
        elif (cnum, cunit, no) in verified:
            status = "confirmed"
        else:
            status = "unverified"
        # 同一个「值+单位」只登记一条，但状态取最严重的（conflict > unverified > confirmed）
        old = dseen.get(key)
        if old is not None:
            rank = {"confirmed": 0, "unverified": 1, "conflict": 2}
            if rank[status] > rank[old["status"]]:
                old["status"] = status
                old["locator"] = f"正文 行 {no}（派生值：加总/测算）"
            continue
        # 已经是 confirmed 事实的（值+单位一致）不重复登记
        if any(str(it.get("factId") or "") and it.get("value") == cnum
               and it.get("unit") == cunit for it in items):
            continue
        entry = {
            "id": f"TR-{len(items) + len(dseen) + 1:04d}",
            "locator": f"正文 行 {no}（派生值：加总/测算）",
            "factId": None,
            "confidence": "derived",
            "status": status,
            "kind": "derived",
            "value": cnum,
            "unit": cunit,
        }
        dseen[key] = entry
        if len(items) + len(dseen) >= 200:
            break
    items.extend(dseen.values())

    for issue in rec["issues"]:
        print(f"[09] CHK-NUM {issue['kind']}: {issue['message']}")
    nconflict = sum(1 for it in items if it.get("status") == "conflict")
    nunverified = sum(1 for it in items if it.get("status") == "unverified")
    print(f"[09] 算术对账：issue {len(rec['issues'])} / 通过 {len(verified)} / "
          f"未验证 {nunverified} / 冲突 {nconflict}")

    nfact = sum(1 for it in items if it.get("kind") != "derived")
    nderived = len(items) - nfact
    trace = {
        "schema": "cardo/trace/v1",
        "generatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "artifact": {"path": rel_artifact, "sha256": sha256_file(args.artifact), "kind": "report"},
        "generatedBy": {"skill": "09-三重自检与边界", "script": "build_trace.py", "version": "1.2.0"},
        "baselineVersion": str(args.baseline_version),
        "scope": {"level": "strong", "note": "交付物正文；断链即失败"},
        "stats": {"total": len(items), "fromFacts": nfact, "derived": nderived,
                  "numIssues": len(rec["issues"]), "numConflict": nconflict,
                  "numUnverified": nunverified},
        "items": items,
    }
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as w:
        json.dump(trace, w, ensure_ascii=False, indent=1)
    print(f"[09] 回溯表 items={len(items)}（事实 {nfact} / 派生 {nderived}） → {args.out}")
    return 0 if items else 1


if __name__ == "__main__":
    sys.exit(main())
