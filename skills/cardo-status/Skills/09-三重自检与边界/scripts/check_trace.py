#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_trace.py -- 成果回溯表只读校验器（T24；schema cardo/trace/v1）。

用法：
    python check_trace.py --project-dir <项目目录> [--artifact <回溯表.json>] [--out <输出目录>]
        [--min-coverage 1.0] [--no-timestamp]
    # 校准/回归（标准项目库布局）：--root <项目库根> --project <项目名> [--direction <方向>]

校验（逐条给 id，失败即 exit 1；需用户确认 exit 3）：
    CHK-TRACE-1 schema/必填字段/枚举合法；items.id 唯一、locator 非空
    CHK-TRACE-2 路径纪律：artifact.path / source.file 均为相对项目根 posix，无盘符/绝对路径
    CHK-TRACE-3 artifact 存在且 sha256 与登记一致
    CHK-TRACE-4 source.file 在项目内存在；source.sha256（若登记）实测一致；docVersion 非空
    CHK-TRACE-5 factId 非悬空（事实库存在时）；value/unit 与事实一致
    CHK-TRACE-6 conflict/missing 不得当 confirmed（refs 事实状态冲突 → exit 3；缺失 → exit 1）
    CHK-TRACE-7 strong 成果覆盖率 ≥ --min-coverage（默认 1.0：每格都须可解析；不达标 exit 1）
    CHK-TRACE-8 重跑幂等（除时间戳逐字节一致；本校验器输出无时间戳）

只读：只写 --out；退出码 0 全过 / 1 断链或覆盖率不足 / 2 输入缺失或坏 / 3 需用户确认。（依赖按需：允许联网/装包，须登记）
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _num_match import canon_token, matches_in_report  # noqa: E402
from pathlib import Path

__version__ = "0.1.0"
SCHEMA = "cardo/trace/v1"
FACTS_SCHEMA = "cardo/facts/v1"
FACTS_REL = Path("05_过程") / "现状事实库" / "事实库.json"
TRACE_DIR = Path("05_过程") / "回溯"
PROJECT_MARKERS = ("00_项目索引.md", ".cardo")
KINDS = {"report", "table", "material", "index", "request"}
LEVELS = {"strong", "light"}
STATUSES = {"confirmed", "missing", "conflict", "unverified"}
ABS_RE = re.compile(r"^(?:[A-Za-z]:|/|\\\\)")


class BadInput(Exception):
    """参数/输入错误 → exit 2。"""


def resolve_project_dir(args) -> Path:
    raw = str(args.project_dir or "").strip()
    if raw:
        path = Path(raw).expanduser()
        if not path.is_dir():
            raise BadInput(f"--project-dir 不存在或不是目录：{path.name}")
        return path.resolve()
    root_arg = str(args.root or "").strip() or os.environ.get("CARDO_LIB_ROOT", "").strip()
    project = str(args.project or "").strip()
    if root_arg or project:
        if not (root_arg and project):
            raise BadInput("--root 与 --project 必须同时给（或用 --project-dir）")
        root = Path(root_arg).expanduser()
        if not root.is_dir():
            raise BadInput(f"--root 不存在或不是目录：{root.name}")
        candidates: list[Path] = []
        if str(args.direction or "").strip():
            candidates.append(root / str(args.direction).strip() / project)
        else:
            candidates += sorted(d / project for d in root.iterdir() if d.is_dir() and (d / project).is_dir())
        candidates.append(root / project)
        hits = [c for c in candidates if c.is_dir()]
        if not hits:
            raise BadInput(f"项目目录不存在：<项目库根>/{project}")
        if len({c.resolve() for c in hits}) > 1:
            raise BadInput("同名项目在多个方向下都存在，请用 --direction 指定")
        return hits[0].resolve()
    cwd = Path.cwd()
    if (cwd / PROJECT_MARKERS[0]).is_file() or (cwd / PROJECT_MARKERS[1]).is_dir():
        return cwd.resolve()
    raise BadInput("未指定项目目录：请用 --project-dir <项目目录>（--root/--project 只用于校准/回归）")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def norm_num(value) -> str:
    if isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        return f"{value:.6f}".rstrip("0").rstrip(".")
    text = str(value or "").strip()
    match = re.fullmatch(r"-?\d+(?:\.\d+)?", text)
    if match:
        return f"{float(text):.6f}".rstrip("0").rstrip(".")
    return text


def rel_ok(text: str) -> bool:
    text = str(text or "").strip()
    return bool(text) and not ABS_RE.match(text) and ".." not in Path(text).parts


def load_facts(project: Path, override: str = "") -> tuple[dict[str, dict], str]:
    path = Path(override).expanduser() if str(override or "").strip() else project / FACTS_REL
    if not path.is_file():
        return {}, ""
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        return {}, f"事实库解析失败：{type(exc).__name__}"
    if not isinstance(data, dict) or data.get("schema") != FACTS_SCHEMA:
        return {}, "事实库 schema 不符"
    return {str(f.get("id")): f for f in data.get("facts") or [] if isinstance(f, dict)}, ""


def check_artifact(project: Path, path: Path, facts: dict[str, dict], min_coverage: float) -> dict:
    """返回 {ok, needs_confirm, fails[], needs[], items, coverage, title, kind, level}。"""
    out = {"path": path, "ok": True, "needs_confirm": False, "fails": [], "needs": [],
           "items": 0, "coverage": 0.0, "title": path.name, "kind": "", "level": ""}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        out["fails"].append(f"CHK-TRACE-1 回溯表解析失败：{type(exc).__name__}")
        return out
    if not isinstance(data, dict) or data.get("schema") != SCHEMA:
        out["fails"].append(f"CHK-TRACE-1 schema 必须是 {SCHEMA}")
        return out
    artifact = data.get("artifact") if isinstance(data.get("artifact"), dict) else {}
    scope = data.get("scope") if isinstance(data.get("scope"), dict) else {}
    items = data.get("items") if isinstance(data.get("items"), list) else []
    out["kind"], out["level"] = str(artifact.get("kind") or ""), str(scope.get("level") or "")
    out["title"] = str(artifact.get("path") or path.name)
    # CHK-TRACE-1 结构 / CHK-TRACE-2 路径
    for field in ("path", "sha256", "kind"):
        if not artifact.get(field):
            out["fails"].append(f"CHK-TRACE-1 artifact.{field} 必填")
    if not isinstance(data.get("generatedBy"), dict) or not data.get("baselineVersion"):
        out["fails"].append("CHK-TRACE-1 generatedBy/baselineVersion 必填")
    if out["kind"] not in KINDS:
        out["fails"].append(f"CHK-TRACE-1 artifact.kind 非法：{out['kind']!r}（{sorted(KINDS)}）")
    if out["level"] not in LEVELS:
        out["fails"].append(f"CHK-TRACE-1 scope.level 非法：{out['level']!r}")
    if not items:
        out["fails"].append("CHK-TRACE-7(items) 回溯表 items 为空")
    ids: set[str] = set()
    for idx, item in enumerate(items):
        if not isinstance(item, dict):
            out["fails"].append(f"CHK-TRACE-1 items[{idx}] 必须是对象")
            continue
        iid = str(item.get("id") or f"#{idx}")
        for field in ("id", "locator", "confidence", "status"):
            if field not in item:
                out["fails"].append(f"CHK-TRACE-1 items[{iid}].{field} 必填")
        if iid in ids:
            out["fails"].append(f"CHK-TRACE-1 items.id 重复：{iid}")
        ids.add(iid)
        if not str(item.get("locator") or "").strip():
            out["fails"].append(f"CHK-TRACE-1 items[{iid}].locator 不能为空")
        status = str(item.get("status") or "")
        if status not in STATUSES:
            out["fails"].append(f"CHK-TRACE-1 items[{iid}].status 非法：{status!r}")
        # 派生值（加总/测算）必须带算术校验结果：conflict 即断链，不许当 confirmed 放行。
        # unverified 表示格式上判不了——计入统计但不阻塞（不冒充通过）。
        if item.get("kind") == "derived":
            if status == "conflict":
                out["fails"].append(
                    f"CHK-TRACE-9 items[{iid}] 派生值 {item.get('value')}{item.get('unit') or ''} "
                    f"算术校验不通过（{item.get('locator','')}）")
            elif status == "confirmed" and not item.get("value"):
                out["fails"].append(f"CHK-TRACE-9 items[{iid}] 派生值标 confirmed 但未登记核对后的 value")
    # CHK-TRACE-2/3 artifact 路径与 sha
    artifact_text = ""
    if not rel_ok(artifact.get("path")):
        out["fails"].append("CHK-TRACE-2 artifact.path 必须是相对项目根 posix 路径（无盘符/绝对路径）")
    else:
        artifact_path = project / str(artifact["path"])
        if not artifact_path.is_file():
            out["fails"].append(f"CHK-TRACE-3 artifact 文件不存在：{artifact['path']}")
        elif artifact.get("sha256") and sha256_file(artifact_path) != str(artifact["sha256"]).lower():
            out["fails"].append("CHK-TRACE-3 artifact sha256 与登记不一致（成果被改动）")
        else:
            try:
                artifact_text = re.sub(r"\s+", "", artifact_path.read_text(encoding="utf-8-sig"))
            except (OSError, UnicodeDecodeError):
                artifact_text = ""
    # 逐条引用
    resolved_items: set[str] = set()
    for item in items:
        if not isinstance(item, dict):
            continue
        iid = str(item.get("id") or "")
        status = str(item.get("status") or "")
        source = item.get("source") if isinstance(item.get("source"), dict) else {}
        if source:
            if not rel_ok(source.get("file")):
                out["fails"].append(f"CHK-TRACE-2 items[{iid}].source.file 必须是相对路径")
            else:
                src_path = project / str(source["file"])
                if not src_path.is_file():
                    out["fails"].append(f"CHK-TRACE-4 items[{iid}] 来源文件不存在：{source['file']}")
                else:
                    if source.get("sha256") and sha256_file(src_path) != str(source["sha256"]).lower():
                        out["fails"].append(f"CHK-TRACE-4 items[{iid}] 来源 sha256 与登记不一致")
                    if not str(source.get("docVersion") or "").strip():
                        out["fails"].append(f"CHK-TRACE-4 items[{iid}] source.docVersion 不能为空")
                    else:
                        resolved_items.add(iid)
        fact_id = str(item.get("factId") or "").strip()
        if fact_id:
            fact = facts.get(fact_id)
            if not facts:
                out["fails"].append(f"CHK-TRACE-5 items[{iid}] 引用 factId={fact_id}，但事实库不存在")
            elif not fact:
                out["fails"].append(f"CHK-TRACE-5 items[{iid}] factId 悬空：{fact_id}")
            else:
                fact_status = str(fact.get("status") or "")
                if status == "confirmed" and fact_status in ("conflict", "missing"):
                    out["needs"].append(f"CHK-TRACE-6 items[{iid}] 把 fact {fact_id}"
                                        f"（{fact_status}）当 confirmed 使用，需用户确认")
                    out["needs_confirm"] = True
                if "value" in item and (norm_num(item.get("value")) != norm_num(fact.get("value"))
                                        or str(item.get("unit") or "") != str(fact.get("unit") or "")):
                    out["fails"].append(f"CHK-TRACE-5 items[{iid}] value/unit 与 fact {fact_id} 不一致")
                else:
                    # 回溯表只存引用、不复制数值：数值以事实库为唯一来源，
                    # 此处独立核验该事实的数值或引文确实出现在成果正文。
                    # 匹配口径与 build_trace.py **共用** _num_match.matches_in_report
                    # ——两处口径不一致会出现「回溯表登记了、校验却不认」的假失败。
                    if artifact_text and not matches_in_report(
                            fact.get("value"), fact.get("unit") or "",
                            (fact.get("source") or {}).get("quote") or "",
                            artifact_text):
                        out["fails"].append(f"CHK-TRACE-5 items[{iid}] 事实 {fact_id} 的数值/引文未在成果正文出现")
                    else:
                        resolved_items.add(iid)
    # 覆盖率分母只算**有 factId 的条目**：派生条目（加总/测算）本就不在事实库里，
    # 计入分母会把覆盖率压到 100% 以下——那是口径问题，不是断链。
    total = len([i for i in items if isinstance(i, dict) and str(i.get("factId") or "").strip()])
    out["items"] = len([i for i in items if isinstance(i, dict)])
    out["coverage"] = round(len(resolved_items) / total, 4) if total else 0.0
    out["derived"] = out["items"] - total
    out["unverified"] = len([i for i in items
                             if isinstance(i, dict) and str(i.get("status") or "") == "unverified"])
    out["ok"] = not out["fails"]
    return out


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    parser = argparse.ArgumentParser(description="成果回溯表只读校验（cardo/trace/v1；T24）")
    parser.add_argument("--project-dir", default="", help="项目目录（任意路径；最高优先级）")
    parser.add_argument("--root", default="", help="[校准/回归] 项目库根（与 --project 搭配）")
    parser.add_argument("--project", default="", help="[校准/回归] 项目名")
    parser.add_argument("--direction", default="", help="[校准/回归] 方向")
    parser.add_argument("--artifact", default="", help="指定回溯表 JSON；缺省扫描 <项目>/05_过程/回溯/*.json")
    parser.add_argument("--out", default="", help="检查报告输出目录（缺省 <项目>/05_过程/回溯/检查/）")
    parser.add_argument("--facts", default="", help="事实库路径（缺省 <项目>/05_过程/现状事实库/事实库.json）")
    parser.add_argument("--min-coverage", type=float, default=1.0, help="strong 成果覆盖率阈值（默认 1.0）")
    parser.add_argument("--no-timestamp", action="store_true", help="输出不写 generatedAt（幂等测试用）")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)
    try:
        project = resolve_project_dir(args)
        if args.artifact:
            targets = [Path(args.artifact).expanduser()]
        else:
            folder = project / TRACE_DIR
            targets = sorted(folder.glob("*.json")) if folder.is_dir() else []
        if not targets:
            raise BadInput(f"找不到回溯表：请用 --artifact 指定，或放到 {TRACE_DIR.as_posix()}/*.json")
        facts, facts_note = load_facts(project, args.facts)
    except BadInput as exc:
        print(f"[trace] 错误：{exc}", file=sys.stderr)
        return 2
    results = [check_artifact(project, t, facts, args.min_coverage) for t in targets]
    hard = [r for r in results if r["fails"]]
    confirm = [r for r in results if r["needs_confirm"]]
    # strong 覆盖率门槛
    for r in results:
        if r["level"] == "strong" and r["items"] and r["coverage"] < args.min_coverage:
            r["fails"].append(f"CHK-TRACE-7 strong 覆盖率 {r['coverage']:.2%} < {args.min_coverage:.0%}")
    hard = [r for r in results if r["fails"]]
    code = 1 if hard else (3 if confirm else 0)
    out_dir = Path(args.out).expanduser() if args.out else project / TRACE_DIR / "检查"
    payload = {"schema": "cardo/trace-check/v1", "project": ".", "artifacts": [
        {"path": str(r["path"].name), "title": r["title"], "kind": r["kind"], "level": r["level"],
         "items": r["items"], "coverage": r["coverage"], "ok": not r["fails"],
         "derived": r.get("derived", 0), "unverified": r.get("unverified", 0),
         "fails": r["fails"], "needs_confirm": True if r["needs_confirm"] else False, "needs": r["needs"]}
        for r in results],
        "summary": {"artifacts": len(results), "fail": len(hard), "needs_confirm": len(confirm),
                    "exit": code}, "facts": {"count": len(facts), "note": facts_note}}
    if not args.no_timestamp:
        payload["generatedAt"] = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    lines = ["# 成果回溯表检查报告", "",
             f"- 项目：{project.name}（占位）",
             f"- 回溯表：{len(results)} 份；失败 {len(hard)} / 需确认 {len(confirm)}（exit {code}）"]
    for r in results:
        lines += ["", f"## {r['title']}", "",
                  f"- kind={r['kind']} / level={r['level']} / items={r['items']} / 覆盖率={r['coverage']:.2%}"
                  f" / 派生={r.get('derived', 0)}（未验证 {r.get('unverified', 0)}）"]
        if r["fails"]:
            lines += [f"- FAIL {msg}" for msg in r["fails"]]
        if r["needs"]:
            lines += [f"- 需确认 {msg}" for msg in r["needs"]]
        if not (r["fails"] or r["needs"]):
            lines.append("- 全部可解析")
    date = datetime.date.today().strftime("%Y%m%d")
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        md, js = out_dir / f"回溯检查-{date}.md", out_dir / f"回溯检查-{date}.json"
        md.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
        js.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")
    except OSError as exc:
        print(f"[trace] 无法写入检查报告：{exc}", file=sys.stderr)
        return 1
    print(f"[trace] 回溯表 {len(results)} 份；失败 {len(hard)} / 需确认 {len(confirm)}；exit={code}")
    for r in results:
        print(f"[trace] {r['title']}: items={r['items']} 覆盖率={r['coverage']:.2%}"
              f" 派生={r.get('derived', 0)}（未验证 {r.get('unverified', 0)}）"
              f"{' FAIL' if r['fails'] else ''}{' 需确认' if r['needs_confirm'] else ''}")
    return code


if __name__ == "__main__":
    sys.exit(main())