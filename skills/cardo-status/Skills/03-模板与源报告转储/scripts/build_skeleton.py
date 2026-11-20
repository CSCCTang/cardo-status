#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第 03 步门禁：模板与源报告转储 → 章节骨架 + 模板指纹校验。

产物：过程文件/03-模板与源报告转储/章节骨架.json
门禁：模板 sha256 与章节契约逐字一致；模板一级/二级节点数 = 契约节点数。
"""
import argparse
import hashlib
import json
import os
import sys

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--template-outline", required=True)
    ap.add_argument("--source-outline", required=True)
    ap.add_argument("--contract", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)

    with open(args.contract, encoding="utf-8") as f:
        contract = yaml.safe_load(f) or {}
    expected = (contract.get("template") or {}).get("sha256", "")
    actual = sha256(args.template)
    tpl = json.load(open(args.template_outline, encoding="utf-8"))
    src = json.load(open(args.source_outline, encoding="utf-8"))
    tpl_nodes = [x for x in tpl if x.get("kind") == "heading" and x.get("level", 9) <= 2]
    src_nodes = [x for x in src if x.get("kind") == "heading" and x.get("level", 9) <= 2]
    c_nodes = contract.get("nodes") or []

    fails = []
    if not expected:
        fails.append("章节契约未登记模板 sha256")
    elif actual != expected:
        fails.append(f"模板 sha256 不一致：{actual} != {expected}")
    if len(tpl_nodes) != len(c_nodes) or len(c_nodes) != 23:
        fails.append(f"模板节点数 {len(tpl_nodes)} / 契约 {len(c_nodes)}，应为 23")

    skeleton = {
        "schema": "cardo/skeleton/v1",
        "templateFile": os.path.basename(args.template),
        "templateSha256": actual,
        "contractSha256": expected,
        "nodeCount": len(tpl_nodes),
        "nodes": [{"index": i + 1, "level": x["level"], "title": x["text"]}
                  for i, x in enumerate(tpl_nodes)],
        "sourceHeadings": len(src_nodes),
        "gate": {"ok": not fails, "fails": fails},
    }
    os.makedirs(args.out, exist_ok=True)
    out = os.path.join(args.out, "章节骨架.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(skeleton, f, ensure_ascii=False, indent=1)
    print(f"[03] 模板节点 {len(tpl_nodes)}（H1={sum(1 for x in tpl_nodes if x['level']==1)} "
          f"H2={sum(1 for x in tpl_nodes if x['level']==2)}）；源报告标题 {len(src_nodes)}")
    print(f"[03] 模板指纹 {'一致' if actual == expected else '不一致'} → {out}")
    if fails:
        for x in fails:
            print(f"[03][FAIL] {x}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
