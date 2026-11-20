#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_ledger.py -- 检索台账校验（cardo/research-ledger/v1），写作前／成文后各跑一次。

规则实现只有一份：09-三重自检与边界/scripts/_ledger.py。本脚本只是命令行外壳，
用于 06 成文之前把「无效条目」挡在写作之外，避免作者引用一条压根不存在的来源。

用法：
    python check_ledger.py --project-dir <项目目录> [--ledger <检索台账.json>] [--out <台账检查.json>]
退出码：0 全过（含"本项目没有台账"）；1 有条目无效或任务未闭环；2 输入错误。
"""

from __future__ import annotations

import argparse
import json
import os
import sys

_SKILLS = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(_SKILLS, "09-三重自检与边界", "scripts"))
import _ledger  # noqa: E402

__version__ = "0.1.0"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="check_ledger.py", description="检索台账校验（只读）")
    ap.add_argument("--project-dir", required=True, help="项目目录（台账 file 字段按此解析）")
    ap.add_argument("--ledger", default="", help="台账路径（缺省按工作区布局自动找）")
    ap.add_argument("--tasks", default="", help="检索任务.json（给了就检查 P0 任务是否闭环）")
    ap.add_argument("--out", default="", help="把检查结果写成 JSON（可选）")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args(argv)

    pd = os.path.abspath(args.project_dir)
    if not os.path.isdir(pd):
        print(f"ERROR: --project-dir 不存在：{args.project_dir}", file=sys.stderr)
        return 2
    led = _ledger.load(pd, args.ledger)
    problems = list(led["problems"])

    # LEDGER-3：P0 任务必须闭环（有有效条目 或 有已查无果记录）
    tasks_path = args.tasks
    if not tasks_path and led["path"]:
        cand = os.path.join(os.path.dirname(led["path"]), "检索任务.json")
        tasks_path = cand if os.path.isfile(cand) else ""
    if tasks_path and os.path.isfile(tasks_path):
        try:
            with open(tasks_path, encoding="utf-8-sig") as fh:
                tasks = json.load(fh).get("tasks") or []
        except (OSError, ValueError) as exc:
            problems.append(f"检索任务解析失败：{type(exc).__name__}")
            tasks = []
        covered = {str(e.get("taskId") or "") for e in led["entries"] if isinstance(e, dict)}
        covered |= {str(e.get("metric") or "") for e in led["entries"] if isinstance(e, dict)}
        for t in tasks:
            if str(t.get("priority") or "").upper() != "P0":
                continue
            tid = str(t.get("taskId") or "")
            metric = str(t.get("metric") or "")
            if tid not in covered and metric not in covered:
                problems.append(f"任务 {tid}（{metric}）未闭环：既无有效条目，也无「已查无果」记录")

    if not led["path"]:
        print("[ledger] 本项目未建立检索台账（检索台账.json）——按无外部数据项目处理")
    else:
        print(f"[ledger] {os.path.relpath(led['path'], pd)}：条目 {len(led['entries'])} / "
              f"有效 {len(led['valid'])} / 无效 {len(led['invalid'])}")
        for eid in led["valid"]:
            print(f"[ledger] OK   {eid}")
        for eid, bad in led["invalid"].items():
            print(f"[ledger] BAD  {eid}：{'；'.join(bad)}")
    for p in problems:
        print(f"[ledger] FAIL {p}")

    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        payload = {"schema": "cardo/ledger-check/v1", "ledger": led["path"],
                   "entries": len(led["entries"]), "valid": sorted(led["valid"]),
                   "invalid": led["invalid"], "problems": problems,
                   "ok": not problems}
        with open(args.out, "w", encoding="utf-8", newline="\n") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
