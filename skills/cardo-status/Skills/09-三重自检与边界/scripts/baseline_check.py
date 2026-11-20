#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""验证基线：首次运行写基线，其后每次运行与基线 diff（只报告，不拦截）。

基线内容：规则/模板指纹、节点数、映射分布、事实数、契约检查计数、分章自检通过数、成稿指纹。
"""
import argparse
import json
import os
import sys

# 基线按包命名空间隔离：四包可能同机共存，文件名里带包名，内容再带 package，
# 避免 A 包的基线被 B 包当成自己的（实测：Status 基线被 Invest 读走，08 自检 8 项失败）。
PACKAGE = "CardoStatus"


def load(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return None


def load_own(path):
    """只认本包写的基线：package 不匹配（含旧版无字段）一律当不存在。"""
    data = load(path)
    if isinstance(data, dict) and data.get("package") == PACKAGE:
        return data
    if data is not None:
        print(f"[09] 忽略非本包基线（package={data.get('package')!r}）：{path}")
    return None


def flat(d, prefix=""):
    out = {}
    for k, v in (d or {}).items():
        key = f"{prefix}{k}"
        if isinstance(v, dict):
            out.update(flat(v, key + "."))
        else:
            out[key] = v
    return out


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--baseline", required=True)
    ap.add_argument("--metrics", required=True, help="本次运行指标 JSON")
    args = ap.parse_args(argv)
    cur = json.load(open(args.metrics, encoding="utf-8"))
    if isinstance(cur, dict):
        cur["package"] = PACKAGE
    base = load_own(args.baseline)
    if not base:
        os.makedirs(os.path.dirname(os.path.abspath(args.baseline)), exist_ok=True)
        with open(args.baseline, "w", encoding="utf-8") as w:
            json.dump(cur, w, ensure_ascii=False, indent=1)
        print(f"[09] 首次运行：已建立验证基线 → {args.baseline}")
        return 0
    a, b = flat(base), flat(cur)
    diffs = []
    # 易变字段不参与基线 diff，否则同一输入的每次运行都会报差异
    volatile = ("generatedAt", "durationSeconds")
    for k in sorted(set(a) | set(b)):
        if any(v in k for v in volatile):
            continue
        if a.get(k) != b.get(k):
            diffs.append(f"{k}: 基线={a.get(k)!r} 本次={b.get(k)!r}")
    if diffs:
        print(f"[09] 与验证基线 diff {len(diffs)} 处：")
        for d in diffs[:20]:
            print("   " + d)
    else:
        print("[09] 与验证基线一致（0 处差异）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
