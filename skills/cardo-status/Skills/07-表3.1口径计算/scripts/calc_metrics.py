#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""calc_metrics.py -- 通用指标计算薄引擎（T23 v0；口径来自 规则/计算口径.yaml）。

用法：
    python calc_metrics.py --project-dir <项目目录> [--rules <口径.yaml>] [--input <计算输入.json>]
        [--metrics CALC-01,CALC-02] [--out <输出目录>] [--format text|json] [--no-timestamp]
    # 校准/回归（标准项目库布局）：--root <项目库根> --project <项目名> [--direction <方向>]

做：按口径算指标，产 05_过程/指标计算/指标计算-<日期>.json + .md（结果 + 计算轨迹）。
不做：不读正文找数（取数归 T22 事实库 / T15 统计底表 / 显式输入）、不自动补数、不换算单位、
      不拍板口径争议（单位不一致/除零/占比合计超差 → needs_confirm，走 G-table/G2/G6 闸门）。

输入 schema cardo/calc-input/v1：
    {"schema": "cardo/calc-input/v1",
     "cells": [{"key": "area_i", "label": "居住用地", "value": 1234.5, "unit": "平方米",
                "ref": {"factId": "F-0007", "source": {"file": "…", "locator": "block 12",
                         "quote": "…", "docVersion": "…"}}}]}
输出 schema cardo/calc-trace/v1（结果 + 轨迹；与 T24 回溯表兼容的子集）：
    {"schema": "cardo/calc-trace/v1", "ruleVersion": 1, "engineVersion": "0.1.0",
     "generatedAt": "…", "rules": "…", "input": "…", "items": [
       {"metric": "CALC-01", "name": "…", "value": 123.45, "unit": "平方米",
        "formula": "sum(area_i)", "inputs": [{…}], "ruleVersion": 1,
        "engineVersion": "0.1.0", "status": "ok|gap|needs_confirm|fail", "note": ""}]}

退出码（优先级从上到下）：
    2  输入/口径缺失或坏（参数、项目目录、输入 JSON、口径 YAML、未知指标 id）
    1  计算/校验失败（status=fail，如非数值/合计行不符），或存在缺数（status=gap，未全部算出）
    3  口径待确认（status=needs_confirm，如单位不一致/除零/占比合计超容差）
    0  全部算出（全部 status=ok）

约定：只读输入、只写 --out；同输入同参数输出逐字节一致（--no-timestamp 去掉 generatedAt）；
资源一律用 Path(__file__) 定位；产物路径只写相对/占位（<技能目录>/…、<项目目录>/…），不写盘符。
"""
from __future__ import annotations

import argparse
import datetime
import json
import math
import os
import re
import sys
from pathlib import Path

# 同目录共享底座（T19 的 YAML 子集解析 + T23 计算底座）；只加脚本目录到 sys.path，不写死路径。
sys.path.insert(0, str(Path(__file__).resolve().parent))
from _contract_lib import parse_yaml, read_text  # noqa: E402
from _calc_lib import (  # noqa: E402
    CHECK_NAMES, DEFAULT_RULES, ENGINE_VERSION, FAIL, GAP, NEEDS_CONFIRM, OK, SCHEMA_TRACE,
    STATUS_ORDER, BadInput, CalcRunError, Parser, collect_refs, disp, esc, fmt, tokenize,
)
from _calc_checks import aggregate_exit, compute_metric  # noqa: E402

SCHEMA_RULES = "cardo/calc-rules/v1"
SCHEMA_INPUT = "cardo/calc-input/v1"
PROJECT_MARKERS = ("00_项目索引.md", ".cardo")
__version__ = ENGINE_VERSION


# ---------------------------------------------------------------- 口径文件
def load_rules(path: Path) -> tuple[int, list[dict]]:
    if not path.is_file():
        raise BadInput(f"口径文件不存在：{disp(path)}（用 --rules 指定，或确认 skill 自带口径）")
    try:
        data = parse_yaml(read_text(path))
    except ValueError as exc:
        raise BadInput(f"口径文件解析失败：{disp(path)}｜{exc}")
    if not isinstance(data, dict) or data.get("schema") != SCHEMA_RULES:
        raise BadInput(f"口径 schema 必须是 {SCHEMA_RULES}：{disp(path)}")
    version = data.get("version")
    if not isinstance(version, int) or isinstance(version, bool) or version < 1:
        raise BadInput("口径 version 必填且为正整数（改口径 = 改标准 → version +1）")
    metrics = data.get("metrics")
    if not isinstance(metrics, list) or not metrics:
        raise BadInput("口径 metrics 不能为空")
    tolerances = data.get("tolerances") if isinstance(data.get("tolerances"), dict) else {}
    seen: list[str] = []
    for i, rule in enumerate(metrics):
        if not isinstance(rule, dict):
            raise BadInput(f"metrics[{i}] 必须是映射")
        mid = rule.get("id")
        if not isinstance(mid, str) or not re.fullmatch(r"CALC-\d{2}", mid):
            raise BadInput(f"metrics[{i}].id 必须形如 CALC-nn：{mid!r}")
        if mid in seen:
            raise BadInput(f"指标 id 重复：{mid}")
        for field in ("name", "unit", "formula"):
            if not isinstance(rule.get(field), str) or not rule[field].strip():
                raise BadInput(f"{mid}.{field} 必填")
        inputs = rule.get("inputs")
        if not isinstance(inputs, list) or not inputs:
            raise BadInput(f"{mid}.inputs 不能为空")
        for spec in inputs:
            if not isinstance(spec, dict) or not str(spec.get("key") or "").strip():
                raise BadInput(f"{mid}.inputs[].key 必填")
        rounding = rule.get("rounding") if isinstance(rule.get("rounding"), dict) else {}
        decimals = rounding.get("decimals", 2)
        if not isinstance(decimals, int) or isinstance(decimals, bool) or not 0 <= decimals <= 6:
            raise BadInput(f"{mid}.rounding.decimals 必须是 0–6 的整数")
        if str(rounding.get("mode") or "half_up") != "half_up":
            raise BadInput(f"{mid}.rounding.mode 只支持 half_up")
        if str(rule.get("missing_policy") or "gap") not in ("gap", "confirm"):
            raise BadInput(f"{mid}.missing_policy 只支持 gap/confirm")
        for chk in rule.get("checks") or []:
            name = (chk if isinstance(chk, str) else str(chk.get("name") or "")) if isinstance(chk, (str, dict)) else ""
            if name not in CHECK_NAMES:
                raise BadInput(f"{mid}.checks 含未知检查项：{name!r}（可用 {'、'.join(sorted(CHECK_NAMES))}）")
            if isinstance(chk, dict):
                if name == "sum_row" and not (chk.get("series") and chk.get("total")):
                    raise BadInput(f"{mid}.checks.sum_row 需要 series/total 参数")
                if name == "ratio_sum_100" and not chk.get("series"):
                    raise BadInput(f"{mid}.checks.ratio_sum_100 需要 series 参数")
        ast = Parser(tokenize(rule["formula"]), rule["formula"]).parse()
        keys = {str(s.get("key")).strip() for s in inputs}
        for ref in collect_refs(ast):
            if ref not in keys and ref not in seen:
                raise BadInput(f"{mid}.formula 引用了未知键/未先算指标：{ref}")
        rule["_ast"], rule["_rulesVersion"], rule["_tolerances"] = ast, version, tolerances
        seen.append(mid)
    return version, metrics


def select_metrics(rules: list[dict], spec: str) -> list[dict]:
    raw = (spec or "").strip()
    if not raw:
        return list(rules)
    want: list[str] = []
    for item in raw.split(","):
        item = item.strip()
        if item and item not in want:
            want.append(item)
    by_id = {r["id"]: r for r in rules}
    unknown = [w for w in want if w not in by_id]
    if unknown:
        raise BadInput(f"未知指标 id：{'、'.join(unknown)}（可用：{'、'.join(by_id)}）")
    return [by_id[w] for w in want]


# ---------------------------------------------------------------- 计算输入
def load_input(path: Path) -> dict[str, list[dict]]:
    if not path.is_file():
        raise BadInput(f"计算输入不存在：{disp(path)}（用 --input 指定 cardo/calc-input/v1 JSON）")
    try:
        data = json.loads(read_text(path))
    except ValueError as exc:
        raise BadInput(f"计算输入不是合法 JSON：{disp(path)}｜{exc}")
    if not isinstance(data, dict) or data.get("schema") != SCHEMA_INPUT:
        raise BadInput(f"计算输入 schema 必须是 {SCHEMA_INPUT}：{disp(path)}")
    cells = data.get("cells")
    if not isinstance(cells, list) or not cells:
        raise BadInput("计算输入 cells 必须是非空数组")
    index: dict[str, list[dict]] = {}
    for i, cell in enumerate(cells):
        if not isinstance(cell, dict):
            raise BadInput(f"cells[{i}] 必须是映射")
        key = str(cell.get("key") or "").strip()
        if not key:
            raise BadInput(f"cells[{i}].key 必填")
        value = cell.get("value")
        if value is not None and (isinstance(value, bool) or not isinstance(value, (int, float))
                                  or not math.isfinite(value)):
            raise BadInput(f"cells[{i}].value 必须是有限数值或 null（key={key}）：{value!r}")
        ref = cell.get("ref")
        if ref is not None and not isinstance(ref, dict):
            raise BadInput(f"cells[{i}].ref 必须是对象（key={key}）：{type(ref).__name__}")
        index.setdefault(key, []).append(cell)
    for key, group in index.items():
        if not key.endswith("_i") and len(group) > 1:
            raise BadInput(f"标量键 {key} 只允许一条 cell（当前 {len(group)} 条），"
                           "请先确认取数口径再重跑（不静默取第一条）")
    return index


# ---------------------------------------------------------------- 轨迹与落盘
def build_trace(args, base: Path, rules_path: Path, version: int, input_path: Path,
                items: list[dict], code: int) -> dict:
    trace = {"schema": SCHEMA_TRACE, "ruleVersion": version, "engineVersion": ENGINE_VERSION}
    if not args.no_timestamp:
        trace["generatedAt"] = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    trace["rules"], trace["input"] = disp(rules_path, base), disp(input_path, base)
    trace["items"] = items
    summary = {name: 0 for name in STATUS_ORDER}
    for item in items:
        summary[item["status"]] += 1
    trace["summary"], trace["exit"] = summary, code
    return trace


def render_md(trace: dict) -> str:
    lines = [f"# 指标计算轨迹（{trace['schema']}）", "",
             f"- 口径：{trace['rules']}（version {trace['ruleVersion']}）",
             f"- 输入：{trace['input']}",
             f"- 引擎：calc_metrics.py {trace['engineVersion']}"]
    if trace.get("generatedAt"):
        lines.append(f"- 生成时间：{trace['generatedAt']}")
    summary = trace["summary"]
    lines.append(f"- 结论：ok {summary.get(OK, 0)} / 缺数 {summary.get(GAP, 0)} / "
                 f"待确认 {summary.get(NEEDS_CONFIRM, 0)} / 失败 {summary.get(FAIL, 0)}"
                 f"（exit {trace['exit']}）")
    lines += ["", "| 指标 | 名称 | 标签 | 值 | 单位 | 状态 | 说明 |",
              "| --- | --- | --- | --- | --- | --- | --- |"]
    for item in trace["items"]:
        lines.append(f"| {item['metric']} | {esc(item['name'])} | {esc(item.get('label', ''))} | "
                     f"{fmt(item['value'])} | {esc(item['unit'])} | {item['status']} | {esc(item['note'])} |")
    lines += ["", "## 输入明细"]
    for item in trace["items"]:
        if not item["inputs"]:
            continue
        title = item["metric"] + (f"（{item['label']}）" if item.get("label") else "")
        lines += ["", f"### {title}", "", "| 键 | 标签 | 值 | 单位 | 引用 |",
                  "| --- | --- | --- | --- | --- |"]
        for cell in item["inputs"]:
            ref = cell.get("ref") or {}
            source = ref.get("source") if isinstance(ref.get("source"), dict) else {}
            ref_text = ref.get("factId") or source.get("file") or ""
            lines.append(f"| {esc(cell.get('key', ''))} | {esc(cell.get('label', ''))} | "
                         f"{fmt(cell.get('value'))} | {esc(cell.get('unit', ''))} | {esc(ref_text)} |")
    return "\n".join(lines) + "\n"


def write_outputs(out_dir: Path, base: Path, json_text: str, md_text: str) -> tuple[list[str], bool]:
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise CalcRunError(f"无法写入输出目录（可能是只读目录/无权限）：{disp(out_dir, base)}｜{exc.strerror or exc}")
    day = datetime.date.today().strftime("%Y%m%d")
    for number in range(1, 10):
        suffix = "" if number == 1 else f"-v{number}"
        json_path = out_dir / f"指标计算-{day}{suffix}.json"
        md_path = out_dir / f"指标计算-{day}{suffix}.md"
        if json_path.exists() or md_path.exists():
            try:
                same = (json_path.is_file() and md_path.is_file() and
                        read_text(json_path) == json_text and read_text(md_path) == md_text)
            except OSError:
                same = False
            if same:
                return [disp(json_path, base), disp(md_path, base)], True
            continue
        try:
            json_path.write_text(json_text, encoding="utf-8", newline="\n")
            md_path.write_text(md_text, encoding="utf-8", newline="\n")
        except OSError as exc:
            raise CalcRunError(f"无法写入计算轨迹（可能是只读目录/无权限）：{disp(out_dir, base)}｜{exc.strerror or exc}")
        return [disp(json_path, base), disp(md_path, base)], False
    raise CalcRunError("同名轨迹已存在 ≥9 个版本，请先清理/归档输出目录再重跑")


# ---------------------------------------------------------------- 项目目录（双口径）
def resolve_project_dir(args) -> Path:
    raw = str(args.project_dir or "").strip()
    if raw:
        path = Path(raw).expanduser()
        if not path.is_dir():
            raise BadInput(f"--project-dir 不存在或不是目录：{disp(path)}（请指向用户自选的项目目录）")
        return path.resolve()
    root_arg = str(args.root or "").strip() or os.environ.get("CARDO_LIB_ROOT", "").strip()
    project = str(args.project or "").strip()
    if root_arg or project:
        if not (root_arg and project):
            raise BadInput("--root 与 --project 必须同时给（或用 --project-dir 指向任意项目目录）")
        root = Path(root_arg).expanduser()
        if not root.is_dir():
            raise BadInput(f"--root 不存在或不是目录：{disp(root)}")
        if any(sep in project for sep in ("/", "\\")):
            raise BadInput(f"--project 只能是项目文件夹名，不能含路径分隔符：{project!r}")
        candidates: list[Path] = []
        if str(args.direction or "").strip():
            candidates.append(root / str(args.direction).strip() / project)
        else:
            candidates += sorted(d / project for d in root.iterdir() if d.is_dir() and (d / project).is_dir())
        candidates.append(root / project)
        hits = [c for c in candidates if c.is_dir()]
        if not hits:
            raise BadInput(f"项目目录不存在：<项目库根>/{project}（或用 --direction 指定方向）")
        if len({c.resolve() for c in hits}) > 1:
            raise BadInput("同名项目在多个方向下都存在，请用 --direction 指定方向")
        return hits[0].resolve()
    cwd = Path.cwd()
    if (cwd / PROJECT_MARKERS[0]).is_file() or (cwd / PROJECT_MARKERS[1]).is_dir():
        return cwd.resolve()
    raise BadInput("未指定项目目录：请用 --project-dir <项目目录>（--root/--project 只用于标准项目库校准/回归）")


# ---------------------------------------------------------------- CLI
def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="通用指标计算薄引擎（口径来自 规则/计算口径.yaml；T23）",
        epilog="退出码：0 全部算出；1 计算/校验失败或缺数；2 输入/口径缺失或坏；3 口径待确认")
    parser.add_argument("--project-dir", default="", help="项目目录（任意路径；最高优先级，T20 口径）")
    parser.add_argument("--root", default="", help="[校准/回归] 项目库根（与 --project 搭配）")
    parser.add_argument("--project", default="", help="[校准/回归] 项目名（与 --root 搭配）")
    parser.add_argument("--direction", default="", help="[校准/回归] 方向（缺省自动在库根下找同名项目）")
    parser.add_argument("--rules", default="", help="口径 YAML（缺省本 skill 的 规则/计算口径.yaml）")
    parser.add_argument("--input", default="", help="计算输入 JSON（缺省 <项目>/05_过程/指标计算/计算输入.json）")
    parser.add_argument("--metrics", default="", help="只算这些指标 id，逗号分隔（缺省全部）")
    parser.add_argument("--out", default="", help="输出目录（缺省 <项目>/05_过程/指标计算/）")
    parser.add_argument("--format", default="text", choices=("text", "json"), help="stdout 格式")
    parser.add_argument("--no-timestamp", action="store_true", help="不写 generatedAt（幂等测试用）")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser.parse_args(argv)


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    args = parse_args(argv)
    try:
        project_dir = resolve_project_dir(args)
        rules_path = Path(args.rules).expanduser() if str(args.rules).strip() else DEFAULT_RULES
        version, rules = load_rules(rules_path)
        metrics = select_metrics(rules, args.metrics)
        input_path = (Path(args.input).expanduser() if str(args.input).strip()
                      else project_dir / "05_过程" / "指标计算" / "计算输入.json")
        cells = load_input(input_path)
        out_dir = Path(args.out).expanduser() if str(args.out).strip() else project_dir / "05_过程" / "指标计算"
        items, results = [], {}
        for rule in metrics:
            produced, result = compute_metric(rule, cells, results)
            items.extend(produced)
            if result is not None:
                results[rule["id"]] = result
        code = aggregate_exit(items)
        trace = build_trace(args, project_dir, rules_path, version, input_path, items, code)
        json_text = json.dumps(trace, ensure_ascii=False, indent=2) + "\n"
        written, idempotent = write_outputs(out_dir, project_dir, json_text, render_md(trace))
        if args.format == "json":
            sys.stdout.write(json_text)
        else:
            print(f"[calc] 项目：{disp(project_dir)}｜口径 version {version}｜引擎 {ENGINE_VERSION}")
            for item in items:
                print(f"[calc] {item['metric']} {item['status']:<13} {fmt(item['value'])} "
                      f"{item['unit']}  {item['note']}")
            print(f"[calc] 产物：{'、'.join(written)}"
                  f"{'（幂等：内容一致，未重写）' if idempotent else ''}")
            print(f"[calc] exit={code}")
        return code
    except BadInput as exc:
        print(f"[calc] 错误：{exc}", file=sys.stderr)
        return 2
    except (CalcRunError, OSError) as exc:
        print(f"[calc] 计算失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())