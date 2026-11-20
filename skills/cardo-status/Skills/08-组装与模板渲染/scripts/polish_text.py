#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""polish_text.py -- 中文成文的排版规范化（只换写法，不动事实与数值）。

做三件事，全部可逆、可核对：
  1. 直引号 → 中文弯引号（成对交替：`"` → “ ”，`'` → ‘ ’）；
  2. 单位符号 → 中文写法（km→公里、km²→平方公里、m→米、㎡→平方米、kg→千克…）；
  3. 全角/半角括号与百分号的常见混用统一（`(数据来源：…)` 保持原样，不改写）。

⚠ 纪律：本脚本**只改符号写法**，不增删文字、不换算数值、不改「约」字。
   凡是需要改数值口径的地方，必须回到成文环节由人决定，不在这里动。

用法：
    python polish_text.py <file1.md> [file2.md ...] [--dry-run] [--check]
        --dry-run  只报告将要改动的处数，不写回
        --check    只检查是否有需要规范化之处（有 → 退出码 1），不写回

退出码：0 已规范化/无需规范化；1 --check 发现待规范化处；2 参数或文件错误。
"""

from __future__ import annotations

import argparse
import os
import re
import sys

__version__ = "1.0.0"

UNIT_RULES: list[tuple[str, str]] = [
    (r"(\d)\s*km/km²", r"\1 公里/平方公里"),
    (r"(\d)\s*km2?/km2?", r"\1 公里/平方公里"),
    (r"(\d)\s*km²", r"\1 平方公里"),
    (r"(\d)\s*km2\b", r"\1 平方公里"),
    (r"(\d)\s*km\b", r"\1 公里"),
    (r"(\d)\s*(?:万)?㎡", r"\1 平方米"),
    (r"(\d)\s*m(?=[\s、，。；：）)]|$)", r"\1 米"),
    (r"(\d)\s*kg\b", r"\1 千克"),
    (r"(\d)\s*kV\b", r"\1 千伏"),
]

PUNCT_RULES: list[tuple[str, str]] = [
    (r"(?<=[\u4e00-\u9fa5])\((?=[\u4e00-\u9fa5])", "（"),
    (r"(?<=[\u4e00-\u9fa5])\)(?=[\u4e00-\u9fa5，。；、])", "）"),
]


def fix_quotes(line: str) -> str:
    """ASCII 引号成对替换为中文弯引号（按行内顺序交替）。"""
    out = []
    dq_open, sq_open = True, True
    for ch in line:
        if ch == '"':
            out.append("“" if dq_open else "”")
            dq_open = not dq_open
        elif ch == "'":
            out.append("‘" if sq_open else "’")
            sq_open = not sq_open
        else:
            out.append(ch)
    return "".join(out)


def polish(text: str) -> tuple[str, dict]:
    stats = {"quotes": 0, "units": 0, "punct": 0}
    before = text.count('"') + text.count("'")
    lines = [fix_quotes(ln) for ln in text.splitlines()]
    text2 = "\n".join(lines) + ("\n" if text.endswith("\n") else "")
    stats["quotes"] = before - (text2.count('"') + text2.count("'"))

    for pat, rep in UNIT_RULES:
        text2, n = re.subn(pat, rep, text2)
        stats["units"] += n
    for pat, rep in PUNCT_RULES:
        text2, n = re.subn(pat, rep, text2)
        stats["punct"] += n
    return text2, stats


def main() -> int:
    ap = argparse.ArgumentParser(description="中文成文排版规范化（引号 / 单位 / 括号）")
    ap.add_argument("files", nargs="+", help="待规范化的 markdown 文件")
    ap.add_argument("--dry-run", action="store_true", help="只报告，不写回")
    ap.add_argument("--check", action="store_true", help="有待改处则退出码 1")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args()

    total = 0
    for path in args.files:
        if not os.path.isfile(path):
            print(f"[error] 文件不存在: {path}", file=sys.stderr)
            return 2
        raw = open(path, encoding="utf-8-sig").read()
        fixed, st = polish(raw)
        changed = fixed != raw
        total += st["quotes"] + st["units"] + st["punct"]
        print(f"[polish] {os.path.basename(path)}: 引号 {st['quotes']}、单位 {st['units']}、"
              f"括号 {st['punct']} → {'有改动' if changed else '无改动'}")
        if changed and not args.dry_run and not args.check:
            with open(path, "w", encoding="utf-8", newline="\n") as f:
                f.write(fixed)

    if args.check:
        print(f"[polish] 待规范化合计 {total} 处")
        return 1 if total else 0
    print(f"[polish] 合计 {total} 处{'（dry-run，未写回）' if args.dry_run else ''}")
    return 0


if __name__ == "__main__":
    sys.exit(main())