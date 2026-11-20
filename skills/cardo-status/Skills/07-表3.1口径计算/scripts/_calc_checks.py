#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_calc_checks.py -- 指标计算与校验层（T23 v0，从 _calc_lib 拆出以控制单文件行数）。

分工：
  * _calc_lib.py：受限公式解析/求值、序列与单元格读取、状态常量；
  * 本模块：逐指标计算（compute_metric）、参数化校验（sum_row / ratio_sum_100）、退出码聚合。

关键约定（L3 P1-1 后）：「公式必需输入」缺失 → 不算值；「校验必需输入」缺失只让对应校验
not_run，不短路整条指标——因此缺 ratio_i 时仍执行 sum_row（合计=逐行之和），结果写进
items[].checks（机读）与 note（人读）。校验/计算规则语义仍以 规则/计算口径.yaml 为准。
"""
from __future__ import annotations

import decimal
import sys
from pathlib import Path

# 运行时是 python313._pth 模式，脚本目录不进 sys.path：兄弟模块必须显式加进来
sys.path.insert(0, str(Path(__file__).resolve().parent))

from _calc_lib import (  # noqa: E402
    ENGINE_VERSION, FAIL, GAP, NEEDS_CONFIRM, OK, Ctx, DivideZero, Series,
    cell_input, cell_label, collect_refs, deci, eval_node, norm_unit,
)


def missing_status(rule: dict) -> str:
    return NEEDS_CONFIRM if str(rule.get("missing_policy") or "gap") == "confirm" else GAP


def rounded(rule: dict, value: decimal.Decimal) -> float:
    decimals = int((rule.get("rounding") or {}).get("decimals", 2))
    return float(value.quantize(decimal.Decimal(1).scaleb(-decimals),
                                rounding=decimal.ROUND_HALF_UP))


def make_item(rule: dict, *, status: str, value=None, label: str = "",
              note: str = "", inputs: list[dict] | None = None,
              checks: list[dict] | None = None) -> dict:
    item = {"metric": rule["id"], "name": rule.get("name") or ""}
    if label:
        item["label"] = label
    item.update({"value": rounded(rule, value) if value is not None else None,
                 "unit": rule.get("unit") or "", "formula": rule.get("formula") or "",
                 "inputs": inputs or [], "ruleVersion": rule["_rulesVersion"],
                 "engineVersion": ENGINE_VERSION, "status": status, "note": note})
    if checks:
        item["checks"] = checks
    return item


def _valid(series: list[dict]) -> bool:
    """序列可用于计算/校验：非空，且所有值都是有限数（排除 null/布尔/非数）。"""
    return bool(series) and all(
        c.get("value") is not None and not isinstance(c.get("value"), bool)
        and isinstance(c.get("value"), (int, float)) for c in series)


def apply_checks(rule: dict, cells: dict, blocked: set[str] | None = None) -> tuple[str, list[str], list[dict]]:
    """参数化检查：sum_row（合计=逐行之和）、ratio_sum_100（占比合计≈100）。

    返回 (status, notes, checks)：checks 是机读校验结果（name/status/note），供轨迹与渲染层消费；
    校验可用输入与「公式必需输入」分离——缺 ratio_i 不阻止 sum_row；blocked 里的 key 视为不可用。
    """
    status, notes, checks = OK, [], []
    blocked = blocked or set()
    tolerances = rule.get("_tolerances") or {}
    for chk in rule.get("checks") or []:
        if not isinstance(chk, dict):
            continue
        name = str(chk.get("name") or "")
        series_key = str(chk.get("series") or "")
        series = cells.get(series_key) or []
        if name == "sum_row":
            total_key = str(chk.get("total") or "")
            total_cells = cells.get(total_key) or []
            usable = (_valid(series) and series_key not in blocked and total_key not in blocked
                      and bool(total_cells) and total_cells[0].get("value") is not None)
            if not usable:
                text = "缺 area_i/area_total，未校验合计=逐行之和"
                checks.append({"name": name, "status": "not_run", "note": text})
                notes.append(text)
                continue
            total = deci(total_cells[0]["value"])
            row_sum = sum((deci(c["value"]) for c in series), decimal.Decimal(0))
            tolerance = deci(tolerances.get("sum_row", 0.01))
            deviation = abs(row_sum - total) / max(abs(total), decimal.Decimal(1))
            bad = deviation > tolerance
            text = (f"合计行校验失败：合计 {total} ≠ 逐行之和 {row_sum}"
                    f"（相对偏差 {deviation:.4f} > {tolerance}，不自动改数）" if bad else
                    f"合计行校验通过：合计 {total} = 逐行之和 {row_sum}")
            checks.append({"name": name, "status": FAIL if bad else OK, "note": text})
            notes.append(text)
            if bad:
                status = FAIL
        elif name == "ratio_sum_100":
            usable = _valid(series) and series_key not in blocked
            if not usable:
                text = "缺占比列（或不可用），未校验占比合计≈100%"
                checks.append({"name": name, "status": "not_run", "note": text})
                notes.append(text)
                continue
            ratio_sum = sum((deci(c["value"]) for c in series), decimal.Decimal(0))
            target = deci(chk.get("target", 100))
            tolerance = deci(tolerances.get("ratio_sum", 0.5))
            bad = abs(ratio_sum - target) > tolerance
            text = (f"占比合计 {ratio_sum}% 偏离 {target}%（容差 {tolerance} 个百分点，走 G-table 确认口径）"
                    if bad else f"占比合计 {ratio_sum}% ≈ {target}%，校验通过")
            checks.append({"name": name, "status": NEEDS_CONFIRM if bad else OK, "note": text})
            notes.append(text)
            if bad and status != FAIL:
                status = NEEDS_CONFIRM
    return status, notes, checks


def _primary_check(rule: dict, checks: list[dict]) -> dict | None:
    """规则里第一项带参检查 = 主校验；主校验未执行时该指标不能算 ok（防止「缺项当通过」）。"""
    names = [str(c.get("name") or "") for c in (rule.get("checks") or []) if isinstance(c, dict)]
    return next((c for c in checks if c["name"] == names[0]), None) if names else None


def compute_metric(rule: dict, cells: dict, results: dict) -> tuple[list[dict], decimal.Decimal | None]:
    """算一个指标：返回 (轨迹条目, 标量结果或 None)。

    「公式必需输入」缺失 → 不算值（value=None）；「校验必需输入」缺失只让对应校验 not_run，不短路整条指标。
    """
    keys = {str(spec["key"]).strip() for spec in rule["inputs"]}
    formula_keys = collect_refs(rule["_ast"]) & keys
    unmet: list[str] = []
    check_missing: list[str] = []
    blocked: set[str] = set()
    for spec in rule["inputs"]:
        key = str(spec["key"]).strip()
        series = cells.get(key) or []
        if not series:
            if spec.get("required"):
                (unmet if key in formula_keys else check_missing).append(key)
                if key not in formula_keys:
                    blocked.add(key)
            continue
        for cell in series:
            label, value = cell_label(cell), cell.get("value")
            unit = norm_unit(cell.get("unit"))
            want = norm_unit(spec.get("unit"))
            if value is None:
                note = f"缺数 {key}（{label}）"
            elif isinstance(value, bool) or not isinstance(value, (int, float)):
                note = f"非数值 {key}（{label}）={value!r}，无法计算"
            elif not unit:
                note = f"缺单位 {key}（{label}），口径要求 {spec.get('unit')}"
            elif want and unit != want:
                note = (f"单位不一致 {key}（{label}）：{cell.get('unit')} ≠ 口径 "
                        f"{spec.get('unit')}（不自动换算，走 G-table）")
            else:
                continue
            if key in formula_keys:
                status = missing_status(rule) if value is None else (
                    FAIL if isinstance(value, bool) or not isinstance(value, (int, float)) else NEEDS_CONFIRM)
                return [make_item(rule, status=status, note=note)], None
            check_missing.append(f"{key}（{label}）")
            blocked.add(key)
            break
    value = None
    if not unmet:
        try:
            value = eval_node(rule["_ast"], Ctx(cells, results))
        except DivideZero as exc:
            return [make_item(rule, status=NEEDS_CONFIRM,
                              note=f"除零：{exc}（不自动产出 0，走 G-table/G6 闸门）")], None
    status, notes, checks = apply_checks(rule, cells, blocked)
    primary = _primary_check(rule, checks)
    if primary is not None and primary["status"] == "not_run" and status == OK:
        status = GAP
    if unmet:
        notes.insert(0, f"缺计算必需输入 {'、'.join(unmet)}，值未算")
    if check_missing:
        notes.insert(0, f"校验输入缺 {'、'.join(check_missing)}（对应校验未执行）")
    if value is None and status == OK and primary is None:
        status = missing_status(rule)  # 没有参数化校验可兜底 → 值算不出就是缺项
    note = "；".join(notes)
    if status in (FAIL, NEEDS_CONFIRM, GAP):
        return [make_item(rule, status=status, note=note, checks=checks)], None
    if isinstance(value, Series):
        return [make_item(rule, status=OK, value=number, label=label,
                          note=note or "分母=逐项面积合计", inputs=[cell_input(cell)], checks=checks)
                for (label, number), cell in zip(value, value.cells)], None
    used = [cell_input(cell) for spec in rule["inputs"]
            for cell in (cells.get(str(spec["key"]).strip()) or [])]
    return [make_item(rule, status=OK, value=value, note=note, inputs=used, checks=checks)], value


def aggregate_exit(items: list[dict]) -> int:
    """2 由调用方在前置阶段抛 BadInput；这里聚合 1/3/0（优先级 fail > needs_confirm > gap）。"""
    statuses = {item["status"] for item in items}
    if FAIL in statuses:
        return 1
    if NEEDS_CONFIRM in statuses:
        return 3
    if GAP in statuses:
        return 1
    return 0