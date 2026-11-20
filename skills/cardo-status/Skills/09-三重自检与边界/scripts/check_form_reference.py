#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""形式参考件串项守卫：成品正文不得出现参考件的专名与数值。

用法:
    python check_form_reference.py --project-dir <工作区> --artifact <正文.md> \
        --register <形式参考.json> --out <输出目录>

退出码: 0 通过（或无参考件）；1 命中（串入）；2 参数/输入错误。
"""
from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path


def load_json(p: Path) -> dict:
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def normalize(text: str) -> tuple[str, list[int]]:
    """去掉空白（含全角空格）与换行，返回 (紧凑文本, 每个字符所在行号)。"""
    chars, lines = [], []
    ln = 1
    for ch in text:
        if ch == "\n":
            ln += 1
            continue
        if ch in " \t\r\u3000\ufeff":
            continue
        chars.append(ch)
        lines.append(ln)
    return "".join(chars), lines


def is_digit(ch: str) -> bool:
    return ch.isdigit() or ch == "."


def find_token(text: str, tok: str, is_num: bool) -> list[int]:
    """返回命中起点下标列表；数值要求两侧不是数字/小数点，避免 113栋 误命中 13栋。"""
    out, start = [], 0
    while True:
        i = text.find(tok, start)
        if i < 0:
            break
        ok = True
        if is_num:
            if i > 0 and is_digit(text[i - 1]):
                ok = False
            j = i + len(tok)
            if j < len(text) and is_digit(text[j]):
                ok = False
        if ok:
            out.append(i)
        start = i + 1
    return out


def context(raw_lines: list[str], line_no: int, tok: str, span: int = 30) -> str:
    s = raw_lines[line_no - 1] if 0 < line_no <= len(raw_lines) else ""
    k = s.find(tok)
    if k < 0:
        return s.strip()[:80]
    a, b = max(0, k - span), min(len(s), k + len(tok) + span)
    return ("…" if a > 0 else "") + s[a:b].strip() + ("…" if b < len(s) else "")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--artifact", required=True, help="成品正文（.md）")
    ap.add_argument("--register", required=True, help="形式参考登记（形式参考.json）")
    ap.add_argument("--out", required=True, help="检查报告输出目录")
    args = ap.parse_args()

    ws = Path(args.project_dir).resolve()
    artifact = Path(args.artifact)
    if not artifact.is_absolute():
        artifact = ws / artifact
    out_dir = Path(args.out)
    if not out_dir.is_absolute():
        out_dir = ws / out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    reg_path = Path(args.register)
    if not reg_path.is_absolute():
        reg_path = ws / reg_path
    reg = load_json(reg_path)
    refs = reg.get("reference") or []
    black = reg.get("blacklist") or {}
    proper = [str(x) for x in (black.get("proper_nouns") or []) if str(x).strip()]
    numbers = [str(x) for x in (black.get("numbers") or []) if str(x).strip()]

    if not artifact.is_file():
        print(f"错误：找不到成品正文 {artifact}", flush=True)
        return 2

    raw = artifact.read_text(encoding="utf-8-sig", errors="replace")
    raw_lines = raw.splitlines()
    text, line_of = normalize(raw)

    hits = []
    for tok in proper:
        for i in find_token(text, tok, False):
            hits.append({"type": "专名", "token": tok, "line": line_of[i],
                         "context": context(raw_lines, line_of[i], tok)})
    for tok in numbers:
        for i in find_token(text, tok, True):
            hits.append({"type": "数值", "token": tok, "line": line_of[i],
                         "context": context(raw_lines, line_of[i], tok)})

    ok = not hits
    now = datetime.now().astimezone().isoformat(timespec="seconds")
    data = {
        "schema": "cardo/form-reference-check/v1",
        "checked_at": now,
        "artifact": str(artifact),
        "register": str(reg_path),
        "references": len(refs),
        "reference_files": [r.get("rel") or r.get("path") for r in refs],
        "blacklist": {"proper_nouns": len(proper), "numbers": len(numbers)},
        "result": "PASS" if ok else "FAIL",
        "hits": hits,
    }
    (out_dir / "形式参考串入检查.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    md = ["# 形式参考件串入检查", "",
          f"- 检查时间：{now}",
          f"- 参考件：{len(refs)} 件" + (
              "（" + "、".join(str(x) for x in data["reference_files"]) + "）" if refs else ""),
          f"- 黑名单：专名 {len(proper)} 条 / 数值 {len(numbers)} 条",
          f"- 结论：{'PASS（零串入）' if ok else 'FAIL（命中 ' + str(len(hits)) + ' 处）'}",
          "",
          "> 口径：只学参考件的章节骨架、句式、表格样式与依据排列方式；"
          "专有名词（项目名/地名/街道/社区名）与数值一律不得进入正文。", ""]
    if hits:
        md += ["## 命中明细", "", "| 类型 | 命中词 | 行号 | 上下文 |", "| --- | --- | --- | --- |"]
        for h in hits[:200]:
            md.append(f"| {h['type']} | {h['token']} | {h['line']} | "
                      f"{h['context'].replace('|', '／')} |")
    (out_dir / "形式参考串入检查.md").write_text("\n".join(md) + "\n", encoding="utf-8")

    print(json.dumps({"result": data["result"], "hits": len(hits),
                      "proper_nouns": len(proper), "numbers": len(numbers)},
                     ensure_ascii=False))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
