#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第 07 步：把事实库中「用地分类面积」转成计算输入（表 3.1 底表）。

只接受同时满足以下条件的 cell：
* 事实所属章节为 t10（土地与空间利用现状）；
* 单位可换算到平方米；
* 标签/引文中出现「用地」且出现「面积」；
并按「用地类别」聚合；无可用事实 → cells=[]（计算引擎落 gap，不补数）。
"""
import argparse
import json
import os
import re
import sys

UNIT_TO_M2 = {"平方米": 1.0, "公顷": 10000.0, "万平方米": 10000.0, "平方公里": 1000000.0, "亩": 666.6667}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--facts", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--note", default="")
    args = ap.parse_args(argv)
    facts = json.load(open(args.facts, encoding="utf-8")) if os.path.exists(args.facts) else {"facts": []}

    cells, seen = [], set()
    for f in facts.get("facts", []):
        secs = f.get("sections") or [f.get("section")]
        if "t10" not in secs:
            continue
        text = f"{f.get('label','')}{f.get('source',{}).get('quote','')}"
        if "用地" not in text or "面积" not in text:
            continue
        unit = str(f.get("unit") or "")
        if unit not in UNIT_TO_M2:
            continue
        label = re.sub(r"^\s*[\d.:：]+", "", f.get("label", "")).strip()
        key = (label, unit, f.get("id"))
        if key in seen:
            continue
        seen.add(key)
        val = float(f["value"]) * UNIT_TO_M2[unit]
        cells.append({"key": "area_i", "label": label, "value": round(val, 2), "unit": "平方米",
                      "ref": {"factId": f.get("id"), "section": "t10",
                              "locator": f.get("source", {}).get("locator", "")}})

    out = {"schema": "cardo/calc-input/v1",
           "note": args.note or "表 3.1 底表：由事实库 t10 的用地面积事实转换；不足则留空，不补数。",
           "cells": cells}
    if not cells:
        # 原始报告未提供表 3.1：保留一个空值占位，让计算引擎按 missing_policy=gap 出缺项，
        # 而不是让管线 02 直接判“输入坏”。value=null 表示“未知”，不构成推断。
        out["cells"] = [{"key": "area_i", "label": "用地分类（原报告未提供表 3.1）",
                         "value": None, "unit": "平方米", "ref": {"factId": None}}]
        out["fallback"] = "no-area-facts"
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as w:
        json.dump(out, w, ensure_ascii=False, indent=1)
    print(f"[07] 计算输入 cells={len(cells)}（无则按缺项，由 calc_metrics 判 gap）→ {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
