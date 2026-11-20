#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_table31.py -- 把 T23 计算轨迹渲染成「表 3.1 用地与空间利用统计」人读 md（T15）。

定位：表 3.1（TBL-07）的**渲染薄层**，只读计算轨迹（cardo/calc-trace/v1），不取数、不算数。
取数归上游（T22 事实库 / T15 底表 / 显式 calc-input），口径与计算归 calc_metrics.py；
本脚本只负责把 CALC-01/02/03 的结果排版成表，并把缺项/待确认逐条列明。

用法：
    python build_table31.py --trace <指标计算-YYYYMMDD.json> --out <表3.1-YYYYMMDD.md>
        [--case-id 小区-1] [--title "表 3.1 用地与空间利用统计"]

约定：
  * 表列 = 用地分类 / 面积（㎡）/ 占比（%）/ 状态 / 来源（ref.source.file + locator + factId）；
  * 分类明细行来自 CALC-02 的逐项条目；合计行来自 CALC-01；校验说明来自 CALC-03；
  * 状态非 ok 的指标不产数值，按「缺项 / 待确认」显式列出；**不补数、不换算**；
  * 只写 --out；产物不写盘符/项目名（引用一律用 calc-input 里的匿名占位）。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

__version__ = "0.1.0"

STATUS_LABEL = {"ok": "ok", "gap": "缺项", "needs_confirm": "待确认", "fail": "校验失败"}


def read_trace(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except OSError as exc:
        raise SystemExit(f"错误：无法读取计算轨迹 {path.name}：{exc}")
    except ValueError as exc:
        raise SystemExit(f"错误：计算轨迹不是合法 JSON {path.name}：{exc}")
    if not isinstance(data, dict) or data.get("schema") != "cardo/calc-trace/v1":
        raise SystemExit(f"错误：计算轨迹 schema 必须是 cardo/calc-trace/v1：{path.name}")
    return data


def fmt(value) -> str:
    if value is None:
        return "—"
    if isinstance(value, float):
        text = f"{value:.6f}".rstrip("0").rstrip(".")
        return text or "0"
    return str(value)


def esc(text) -> str:
    return str(text).replace("|", "／").replace("\n", " ")


def ref_text(cell: dict) -> str:
    ref = cell.get("ref") if isinstance(cell.get("ref"), dict) else {}
    source = ref.get("source") if isinstance(ref.get("source"), dict) else {}
    parts = [str(source.get("file") or ""), str(source.get("locator") or "")]
    if ref.get("factId"):
        parts.insert(0, str(ref["factId"]))
    return " · ".join(p for p in parts if p)


def items_of(trace: dict, metric: str) -> list[dict]:
    return [it for it in trace.get("items") or [] if it.get("metric") == metric]


def render(trace: dict, case_id: str, title: str) -> str:
    rule_version = trace.get("ruleVersion")
    engine = trace.get("engineVersion")
    summary = trace.get("summary") or {}
    rows = items_of(trace, "CALC-02")
    total = items_of(trace, "CALC-01")
    check = items_of(trace, "CALC-03")
    lines = [f"# {title}", ""]
    if case_id:
        lines.append(f"- 案例：{case_id}（匿名 ID）")
    lines += [f"- 口径：{trace.get('rules')}（version {rule_version}）",
              f"- 输入：{trace.get('input')}",
              f"- 引擎：calc_metrics.py {engine}" + (f"；案例 {case_id}" if case_id else ""),
              f"- 结论：ok {summary.get('ok', 0)} / 缺项 {summary.get('gap', 0)} / "
              f"待确认 {summary.get('needs_confirm', 0)} / 校验失败 {summary.get('fail', 0)}"
              f"（exit {trace.get('exit')}）", ""]
    lines += ["| 用地分类 | 面积（㎡） | 占比（%） | 状态 | 来源 |",
              "| --- | --- | --- | --- | --- |"]
    def check_of(name: str) -> str:
        """取 CALC-03 里指定带参校验的状态（items[].checks）；无该字段的旧轨迹从 note 兜底推断。"""
        if check:
            for entry in check[0].get("checks") or []:
                if str(entry.get("name")) == name:
                    return str(entry.get("status") or "not_run")
            note = str(check[0].get("note") or "")
            if name == "sum_row":
                if "合计行校验通过" in note:
                    return "ok"
                if "合计行校验失败" in note:
                    return "fail"
        return "not_run"

    if rows and all(it.get("status") == "ok" for it in rows):
        for item in rows:
            cells = item.get("inputs") or []
            area = cells[0].get("value") if cells else None
            source = ref_text(cells[0]) if cells else ""
            label = str(item.get("label") or "—")
            if "未分类" in label:
                label = label.replace("，未分类", "") + "（未分类，非 TBL-07 分类行）"
            lines.append(f"| {esc(label)} | {fmt(area)} | {fmt(item.get('value'))} | ok | {esc(source)} |")
        total_value = total[0].get("value") if total else None
        # 合计状态 = CALC-01 ok 且 CALC-03 的 sum_row 通过；sum_row 未执行 → 显式「未校验」，fail → fail。
        # 合计占比只取 CALC-03 的值（不在渲染层累加 CALC-02；值缺失 → 「—」）。
        sum_row = check_of("sum_row")
        total_src_ok = bool(total) and total[0].get("status") == "ok"
        total_status = ("ok" if sum_row == "ok" and total_src_ok else
                        "fail" if sum_row == "fail" else "未校验（CALC-03 gap）")
        ratio_total = check[0].get("value") if check and check[0].get("value") is not None else None
        lines.append(f"| **合计** | {fmt(total_value)} | {fmt(ratio_total)} | {total_status} | "
                     f"逐行之和（CALC-01）；合计行校验（CALC-03 sum_row: {sum_row}）；占比合计（CALC-03 值） |")
    else:
        missing = "；".join(STATUS_LABEL.get(it.get("status"), str(it.get("status")))
                           for it in (rows or total or check)) or "缺项"
        lines.append(f"| 用地分类明细（缺项，非 TBL-07 分类行） | — | — | {missing} | 源材料未提供可复核的完整用地分类面积（不补数） |")
        lines.append(f"| **合计** | — | — | {missing} | 无完整分类底表，合计不可得 |")
    lines += ["", "## 口径与校验说明", ""]
    for item in check or []:
        lines.append(f"- CALC-03 合计行一致性：{STATUS_LABEL.get(item.get('status'), item.get('status'))}"
                     f"（{item.get('value') if item.get('value') is not None else '—'}）；{esc(item.get('note') or '')}")
    if not check:
        lines.append("- CALC-03 未产出（源材料缺分类明细）")
    lines += ["", "## 缺项 / 待确认（显式，不入数值）", ""]
    flagged = [it for it in trace.get("items") or [] if it.get("status") != "ok"]
    unclassified = any("未分类" in str(it.get("label") or "") for it in rows)
    if unclassified:
        lines.append("- 用地分类明细：缺项（源材料只提供用地总量、未提供分类面积；按口径不归并、不补数，"
                     "分类口径待上游底表/事实库补齐）")
    if flagged:
        for item in flagged:
            lines.append(f"- {item.get('metric')} {STATUS_LABEL.get(item.get('status'), item.get('status'))}："
                         f"{esc(item.get('note') or '')}")
    else:
        lines.append("- 无（全部指标 status=ok）")
    lines += ["", "## 边界", "",
              "- 本表只渲染计算轨迹；取数由上游底表提供，未读正文、未补数、未换算单位。",
              "- 口径争议（分类口径/单位/占比基数）走 G-table 闸门；本表不自动拍板。",
              "- 表 3.1（TBL-07）当前在表格契约里为 `planned`；是否升 active 属改标准，须升版本 + 回归。", ""]
    return "\n".join(lines)


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    parser = argparse.ArgumentParser(description="把计算轨迹渲染成表 3.1 人读 md（T15）")
    parser.add_argument("--trace", required=True, help="指标计算 JSON（cardo/calc-trace/v1）")
    parser.add_argument("--out", required=True, help="输出 md 路径")
    parser.add_argument("--case-id", default="", help="匿名案例 ID（只用于标题下方展示）")
    parser.add_argument("--title", default="表 3.1 用地与空间利用统计", help="表标题")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    trace_path = Path(args.trace)
    if not trace_path.is_file():
        print(f"[table31] 错误：计算轨迹不存在：{trace_path.name}", file=sys.stderr)
        return 2
    out_path = Path(args.out)
    try:
        out_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        print(f"[table31] 错误：无法创建输出目录：{exc}", file=sys.stderr)
        return 1
    trace = read_trace(trace_path)
    try:
        text = render(trace, args.case_id, args.title)
        out_path.write_text(text, encoding="utf-8", newline="\n")
    except OSError as exc:
        print(f"[table31] 错误：无法写入表格：{exc}", file=sys.stderr)
        return 1
    print(f"[table31] 已写 {out_path.name}（trace exit {trace.get('exit')}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())