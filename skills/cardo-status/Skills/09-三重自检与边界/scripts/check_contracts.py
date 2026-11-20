#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_contracts.py -- 报告 × 章节/表格契约的只读校验器（差异清单落 --out，不改任何产物）。

检查项（契约 assertions 登记）：CHK-CONTRACT-1 / CHK-TPL-1 / CHK-TPL-2 / CHK-SRC-1 /
CHK-BAN-1 / CHK-NUM-1 / CHK-GAP-1 / CHK-REQ-1 / CHK-TBL-* / CHK-DOCX-1 / CHK-DIR-1。
明细见 _contract_engine.py 与插件文档 plugins/cardo-status/docs/专题/章节与表格契约设计.md §4.3。

用法（T20 方向：--project-dir 优先，任意目录；--root/--project 只用于校准/回归）：
    python check_contracts.py --project-dir <项目目录> [--direction <方向>] \
        [--template <模板.docx>] [--report <正文.md>] [--docx <报告.docx>] \
        [--contracts <规则目录>] [--source-dump <转储文件或目录>] [--out <输出目录>] [--json]
    # 兼容校准集布局：--root <项目库根> --project <项目名> [--direction <方向>]
    # 模板解析优先级：--template > <项目>/05_过程/模板/ > skill 自带 模板/报告模板-现状分析.docx

只读：报告、模板、转储、契约都不改；只写 --out 下的 契约检查-YYYY-MM-DD.{json,md}。
退出码沿用 skill 约定：0 全过；1 硬失败；2 输入/参数缺失；3 需用户确认（方向）；4 docx 回读失败。
"""

from __future__ import annotations

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _contract_lib as lib  # noqa: E402
import _contract_engine as engine  # noqa: E402
import _ledger  # noqa: E402
import _skill_paths  # noqa: E402  （资源定位底座：自带副本优先，兄弟 skill 仅回退）

__version__ = "1.2.0"
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_CONTRACTS = os.path.join(SKILL_DIR, "规则")
DEFAULT_TEMPLATE_NAME = "报告模板-现状分析.docx"
PROJECT_MARKERS = ("00_项目索引.md", ".cardo")
REPORT_HINT = "现状分析报告"


class InputError(Exception):
    """输入/参数缺失（退出码 2）。"""

def load_contracts(contracts_dir: str) -> tuple[dict, dict]:
    ch_path = os.path.join(contracts_dir, "章节契约.yaml")
    tb_path = os.path.join(contracts_dir, "表格契约.yaml")
    for p in (ch_path, tb_path):
        if not os.path.isfile(p):
            raise InputError(
                f"契约文件缺失: {lib.rel_posix(p, os.getcwd())}（用 --contracts 指定规则目录）")
    try:
        chapter = lib.parse_yaml(lib.read_text(ch_path))
        tables = lib.parse_yaml(lib.read_text(tb_path))
    except ValueError as exc:
        raise InputError(f"契约 YAML 解析失败: {exc}") from exc
    if not isinstance(chapter, dict) or not isinstance(chapter.get("nodes"), list):
        raise InputError("章节契约结构异常（缺 nodes）")
    if not isinstance(tables, dict) or not isinstance(tables.get("tables"), list):
        raise InputError("表格契约结构异常（缺 tables）")
    return chapter, tables

def read_docx_structure(path: str) -> dict:
    """回读 docx：标题序列 / Title 数 / 表格数；打不开抛 RuntimeError（CHK-DOCX-1 按 exit 4）。"""
    try:
        from docx import Document
    except ImportError as exc:  # pragma: no cover - 运行时应自带
        raise RuntimeError("运行时缺少 python-docx（应与 build_report_docx.py 一致）") from exc
    try:
        doc = Document(path)
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"打开 docx 失败: {type(exc).__name__}: {exc}") from exc
    headings, titles = [], 0
    for p in doc.paragraphs:
        style = (p.style.name or "").strip()
        if style == "Title":
            titles += 1
        m = re.match(r"^(?:Heading|标题)\s*([1-9])$", style, re.IGNORECASE)
        if m:
            headings.append((int(m.group(1)), p.text.strip()))
    return {"headings": headings, "titles": titles, "tables": len(doc.tables)}


def sibling_rule(name: str) -> str:
    """定位齐套规则文件（必备材料清单 / 项目库配置）。

    旧写法把兄弟路径写死为 `<skill上级>/cardo-project-material/规则/<name>`，在独立调优
    布局下不可达 → 必备材料清单校验子项被静默跳过（materials_path 指向不存在的文件）。
    现在走 _skill_paths：本 skill 自带 `规则齐套/` 优先，兄弟目录仅作回退。
    """
    return _skill_paths.resolve_rule(name)


def get_directions(chapter: dict) -> list[str]:
    cfg = sibling_rule("项目库配置.yaml")
    if os.path.isfile(cfg):
        try:
            dirs = lib.parse_yaml(lib.read_text(cfg)).get("directions")
            if isinstance(dirs, list) and dirs:
                return [str(d) for d in dirs]
        except ValueError:
            pass
    keys: list[str] = []
    for node in chapter.get("nodes") or []:
        for d in ((node.get("requirements") or {}).get("by_direction") or {}):
            if d not in keys:
                keys.append(d)
    return keys


def resolve_project_dir(args) -> str:
    if args.project_dir:
        pd = os.path.abspath(args.project_dir)
        if not os.path.isdir(pd):
            raise InputError(f"--project-dir 不存在或不是目录: {args.project_dir}")
        return pd
    if args.root or args.project:
        if not (args.root and args.project):
            raise InputError("--root 与 --project 必须同时给（或改用 --project-dir）")
        root = os.path.abspath(args.root)
        cands = []
        if args.direction:
            cands.append(os.path.join(root, args.direction, args.project))
        cands.append(os.path.join(root, args.project))
        for pd in cands:
            if os.path.isdir(pd):
                return pd
        raise InputError(f"项目目录不存在: {' / '.join(cands)}")
    cwd = os.getcwd()
    if any(os.path.exists(os.path.join(cwd, marker)) for marker in PROJECT_MARKERS):
        return cwd
    raise InputError("未指定项目目录：请用 --project-dir <项目目录>（--root/--project 只用于标准项目库回归）")


def mapping_data(pd: str) -> dict:
    path = os.path.join(pd, "03_拆解成果", "章节映射.yaml")
    if not os.path.isfile(path):
        return {}
    try:
        data = lib.parse_yaml(lib.read_text(path))
        return data if isinstance(data, dict) else {}
    except ValueError:
        return {}


def resolve_report(args, pd: str) -> str:
    if args.report:
        if not os.path.isfile(args.report):
            raise InputError(f"--report 不存在: {args.report}")
        return os.path.abspath(args.report)
    folder = os.path.join(pd, "04_成果报告")
    cands = [os.path.join(folder, n) for n in sorted(os.listdir(folder))
             if n.endswith(".md") and REPORT_HINT in n] if os.path.isdir(folder) else []
    if not cands:
        raise InputError(f"找不到正文 md（{lib.rel_posix(folder, os.getcwd())} 下无 *{REPORT_HINT}*.md）；"
                     "请用 --report 指定")
    return max(cands, key=os.path.getmtime)


def resolve_docx(args, pd: str, report: str) -> str | None:
    if args.docx:
        if not os.path.isfile(args.docx):
            raise InputError(f"--docx 不存在: {args.docx}")
        return os.path.abspath(args.docx)
    stem = os.path.splitext(report)[0] + ".docx"
    if os.path.isfile(stem):
        return stem
    folder = os.path.dirname(report)
    cands = [os.path.join(folder, n) for n in os.listdir(folder)
             if n.endswith(".docx") and REPORT_HINT in n] if os.path.isdir(folder) else []
    return max(cands, key=os.path.getmtime) if cands else None


def resolve_template(args, pd: str, chapter: dict) -> str:
    """模板解析优先级（T20）：--template > <项目>/05_过程/模板/ > skill 自带 模板/报告模板-现状分析.docx。"""
    if args.template:
        if not os.path.isfile(args.template):
            raise InputError(f"--template 不存在: {args.template}")
        return os.path.abspath(args.template)
    names: list[str] = []
    base = os.path.basename(str((chapter.get("template") or {}).get("file") or ""))
    if base:
        names.append(base)
    if DEFAULT_TEMPLATE_NAME not in names:
        names.append(DEFAULT_TEMPLATE_NAME)
    for folder in (os.path.join(pd, "05_过程", "模板"), os.path.join(SKILL_DIR, "模板")):
        for name in names:
            cand = os.path.join(folder, name)
            if os.path.isfile(cand):
                return cand
    raise InputError("找不到模板：请用 --template 指定，或放入 <项目>/05_过程/模板/，"
                     "或使用 skill 自带 模板/" + DEFAULT_TEMPLATE_NAME)


def resolve_source(args, pd: str, mapping: dict, template: str) -> str:
    """来源转储：优先 --source-dump，其次映射 source 对应的 blocks/tables，最后全量转储。"""
    dump_dir = os.path.join(pd, "05_过程", "原文转储")
    if args.source_dump:
        if not os.path.exists(args.source_dump):
            raise InputError(f"--source-dump 不存在: {args.source_dump}")
        if os.path.isdir(args.source_dump):
            return "\n".join(lib.read_text(os.path.join(args.source_dump, n))
                             for n in sorted(os.listdir(args.source_dump))
                             if n.endswith((".blocks.txt", ".tables.txt")))
        return lib.read_text(args.source_dump)
    src = str(mapping.get("source") or "")
    stem = os.path.splitext(os.path.basename(src))[0] if src else ""
    if stem:
        parts = [os.path.join(dump_dir, stem + ext) for ext in (".blocks.txt", ".tables.txt")]
        if all(os.path.isfile(p) for p in parts):
            return "\n".join(lib.read_text(p) for p in parts)
    if os.path.isdir(dump_dir):
        tpl_stem = os.path.splitext(os.path.basename(template))[0]
        texts = [lib.read_text(os.path.join(dump_dir, n)) for n in sorted(os.listdir(dump_dir))
                 if n.endswith((".blocks.txt", ".tables.txt")) and
                 os.path.splitext(n)[0] != tpl_stem]
        if texts:
            return "\n".join(texts)
    return ""


def read_card_direction(pd: str) -> str:
    """只读 .cardo/project.json 的 direction（T20：不写；坏文件/非字符串一律当没有）。"""
    path = os.path.join(pd, ".cardo", "project.json")
    if not os.path.isfile(path):
        return ""
    try:
        import json
        data = json.loads(lib.read_text(path))
    except (OSError, ValueError):
        return ""
    value = data.get("direction") if isinstance(data, dict) else None
    return value.strip() if isinstance(value, str) else ""


def resolve_direction(args, pd: str, directions: list[str]) -> str | None:
    raw_arg = str(args.direction or "").strip()
    raw_map = str(mapping_data(pd).get("direction") or "").strip()
    raw_card = read_card_direction(pd)

    def canon(raw: str) -> str | None:
        raw = str(raw or "").strip()
        return raw if raw in directions else (raw + "项目" if raw + "项目" in directions else None)

    arg, mp, card = canon(raw_arg), canon(raw_map), canon(raw_card)
    # 命令行与项目记录（映射/.cardo）冲突 = 口径冲突，走闸门（exit 3），不许静默用其中一个
    for other in (mp, card):
        if arg and other and arg != other:
            return None
    return arg or mp or card


def match_tables(md_tables: list[dict], contracts: list[dict]) -> dict[str, list[dict]]:
    """表格按表头列识别（不依赖位置）：返回 契约表 id -> 正文表格列表。"""
    out: dict[str, list[dict]] = {}
    for tb in md_tables:
        hdr = [lib.norm_cell(c) for c in tb["header"]]
        for c in contracts:
            cols = [lib.norm_cell(x) for x in (c.get("columns") or [])]
            mode = str(c.get("columns_mode") or "exact")
            if (hdr == cols) if mode != "subset" else all(x in hdr for x in cols):
                out.setdefault(c["id"], []).append(tb)
                break
    return out


# --------------------------------------------------------------------------- #
# 差异清单（确定性输出：无时间戳、无绝对路径；同输入重复跑逐字节相同）
# --------------------------------------------------------------------------- #
def build_payload(ctx: dict, checks: list, manual: list, code: int) -> dict:
    summary = {s: sum(1 for c in checks if c.status == s) for s in ("fail", "hint", "skip")}
    summary.update({"checks": len(checks), "manual": len(manual), "exit": code})
    tpl_sha = lib.sha256_file(ctx["template"])
    return {
        "schema": "cardo/contract-check/v1",
        "project_dir": ".",
        "direction": ctx.get("direction") or "",
        "template": {
            "path": lib.rel_posix(ctx["template"], ctx["project_dir"]),
            "sha256": tpl_sha,
            "contract_sha256": str((ctx["chapter"].get("template") or {}).get("sha256") or ""),
            "match": tpl_sha == str((ctx["chapter"].get("template") or {}).get("sha256") or "").lower(),
        },
        "contracts": {
            "chapter": lib.rel_posix(ctx["chapter_path"], ctx["project_dir"]),
            "table": lib.rel_posix(ctx["table_path"], ctx["project_dir"]),
        },
        "summary": summary,
        "checks": [{"id": c.id, "status": c.status, "items": c.items} for c in checks],
        "facts": ctx.get("facts_gaps") or {},
        "manual": manual,
    }


def render_md(payload: dict) -> str:
    s = payload["summary"]
    lines = [
        "# 契约检查报告", "",
        f"- 项目目录：{payload['project_dir']}",
        f"- 方向：{payload['direction'] or '（未判定）'}",
        f"- 模板：{payload['template']['path']}（sha256 "
        f"{'一致' if payload['template']['match'] else '不一致'}）",
        f"- 契约：{payload['contracts']['chapter']} / {payload['contracts']['table']}",
        f"- 结论：失败 {s['fail']} / 提示 {s['hint']} / 跳过 {s['skip']} / "
        f"人工确认 {s['manual']}（exit {s['exit']}）", "",
    ]
    for label, status in (("失败项（必须处理）", "fail"), ("提示 / 可复核差异", "hint"),
                          ("跳过", "skip")):
        lines.append(f"## {label}")
        items = [c for c in payload["checks"] if c["status"] == status]
        if not items:
            lines.append("（无）")
        for c in items:
            lines.append(f"- [{c['id']}]")
            lines.extend(f"  - {it}" for it in c["items"])
        lines.append("")
    lines.append("## 人工确认清单（check: manual，不判 fail）")
    if not payload["manual"]:
        lines.append("（无）")
    for m in payload["manual"]:
        lines.append(f"- [{m['id']}] {m['node']}：{m['note'] or '（人工判断）'}")
    lines.append("")
    facts = payload.get("facts") or {}
    if facts:
        lines.append("## 事实库缺口（T22：missing/conflict 汇入数据缺口清单，不影响 exit）")
        lines.append(f"- 事实库：{facts.get('facts', 0)} 条事实" +
                     (f"；{facts['note']}" if facts.get("note") else ""))
        for miss in facts.get("missing") or []:
            lines.append(f"- missing：{miss.get('section') or '（未归属）'} {miss.get('heading') or ''}"
                         f"（{miss.get('note') or '该节无事实'}）")
        for cf in facts.get("conflict") or []:
            lines.append(f"- conflict：{cf.get('labelKey')}（{cf.get('unit')}）值 {cf.get('values')}")
        if not (facts.get("missing") or facts.get("conflict")):
            lines.append("- （无）")
        lines.append("")
    return "\n".join(lines)


def write_outputs(out_dir: str, payload: dict, day: str | None = None) -> tuple[str, str]:
    import json
    from datetime import date
    os.makedirs(out_dir, exist_ok=True)
    name = "契约检查-" + (day or date.today().isoformat())
    md_path, js_path = os.path.join(out_dir, name + ".md"), os.path.join(out_dir, name + ".json")
    with open(md_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(render_md(payload))
    with open(js_path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return md_path, js_path


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="check_contracts.py",
        description="报告 × 章节/表格契约的只读校验（差异清单落 --out，不改任何产物）。")
    ap.add_argument("--project-dir", default="", help="项目目录（任意路径，T20 口径；最高优先级）")
    ap.add_argument("--root", default="", help="[校准/回归] 项目库根（与 --project 搭配）")
    ap.add_argument("--project", default="", help="[校准/回归] 项目名（与 --root 搭配）")
    ap.add_argument("--direction", default="",
                    help="方向（可选；缺省读 章节映射.yaml 或 .cardo/project.json 的 direction）")
    ap.add_argument("--template", default="",
                    help="模板 .docx（缺省按 项目/05_过程/模板/ → skill 自带 模板/ 顺序找）")
    ap.add_argument("--report", default="", help="正文 md（缺省找 04_成果报告/*现状分析报告*.md）")
    ap.add_argument("--docx", default="", help="报告 docx（缺省取正文同名 .docx；找不到则跳过回读）")
    ap.add_argument("--contracts", default=DEFAULT_CONTRACTS, help="契约规则目录（默认本 skill 的 规则/）")
    ap.add_argument("--source-dump", default="", help="原始报告转储文件或目录（缺省用 05_过程/原文转储）")
    ap.add_argument("--ledger", default="",
                    help="检索台账 json（缺省按工作区布局找 过程文件/00-总控/网络补料/检索台账.json）")
    ap.add_argument("--out", default="", help="差异清单输出目录（默认 <项目>/05_过程/契约检查/）")
    ap.add_argument("--json", action="store_true", help="同时把 JSON 差异清单打印到 stdout")
    ap.add_argument("--gap-mode", action="store_true",
                    help="缺口版：篇幅类要求（min_chars）降级为提示——"
                         "缺料节写不满字数是缺料所致，不是写作缺陷")
    ap.add_argument("--edition", choices=["clean", "gap"], default="",
                    help="版本口径：clean（默认干净版，不要求缺项闭环、正文禁缺口口语）"
                         "／gap（缺项留痕版）；留空按 clean 处理")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return ap


def safe_error_text(text: str, pd: str) -> str:
    """错误文本去绝对路径：项目内路径 → 相对项目根；其余盘符路径 → <绝对路径> 占位。"""
    text = str(text or "").replace(os.path.abspath(pd), ".")
    # 前面不能紧跟字母/数字/+/./-：否则 http:// 里的 p:/ 会被当成盘符掩掉（实测踩过）。
    return re.sub(r"(?<![A-Za-z0-9+.\-])[A-Za-z]:[\\/][^\s'\"]*", "<绝对路径>", text)


def build_context(args) -> dict:
    contracts_dir = os.path.abspath(args.contracts)
    chapter, table_contract = load_contracts(contracts_dir)
    pd = resolve_project_dir(args)
    template = resolve_template(args, pd, chapter)
    report = resolve_report(args, pd)
    docx = resolve_docx(args, pd, report)
    mapping = mapping_data(pd)
    source_text = resolve_source(args, pd, mapping, template)
    directions = get_directions(chapter)
    direction = resolve_direction(args, pd, directions)
    lines = lib.read_text(report).splitlines()
    heads = lib.heading_scope(lines, lib.iter_headings(lines))
    try:
        tpl_heads = read_docx_structure(template)["headings"]
    except RuntimeError as exc:
        raise InputError(f"模板无法回读（--template 是否是 .docx？）: {exc}") from exc
    docx_struct, docx_error = None, ""
    if docx:
        try:
            docx_struct = read_docx_structure(docx)
        except RuntimeError as exc:
            docx_error = safe_error_text(f"docx 回读失败: {exc}", pd)
    max_level = int((chapter.get("template") or {}).get("max_level") or 2)
    tpl = [(lv, t) for lv, t in tpl_heads if lv <= max_level]
    matched, owner, ranges = engine.build_mapping(tpl, heads)
    md_tables = lib.iter_tables(lines)
    table_list = table_contract.get("tables") or []
    return {
        "chapter": chapter, "table_contract": table_contract,
        "chapter_path": os.path.join(contracts_dir, "章节契约.yaml"),
        "table_path": os.path.join(contracts_dir, "表格契约.yaml"),
        "materials_path": sibling_rule("必备材料清单.yaml"),
        "project_dir": pd, "template": template, "report": report, "docx": docx,
        "direction": direction, "directions": directions, "lines": lines, "heads": heads,
        "md_tables": md_tables, "tpl": tpl, "matched": matched,
        "owner": owner, "ranges": ranges,
        "docx_struct": docx_struct, "docx_error": docx_error,
        "table_matches": match_tables(md_tables, table_list),
        "table_map": {str(t.get("id")): t for t in table_list},
        "source_text": source_text,
        # 检索台账（可选）：CHK-CITE 与 CHK-NUM 的台账数字来源池都从这里取
        "ledger": _ledger.load(pd, str(getattr(args, "ledger", "") or "")),
        "facts_path": os.path.join(pd, "05_过程", "现状事实库", "事实库.json"),
        # 缺口版：篇幅类要求降级为提示（缺料节写不满字数不是写作缺陷）
        "gapMode": bool(getattr(args, "gap_mode", False)),
        "edition": (getattr(args, "edition", "") or
                    ("gap" if getattr(args, "gap_mode", False) else "clean")),
    }


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass
    args = build_parser().parse_args(argv)
    try:
        ctx = build_context(args)
    except InputError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return engine.E_INPUT
    try:
        checks, manual, code = engine.run_checks(ctx)
    except InputError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return engine.E_INPUT
    except Exception as exc:  # noqa: BLE001
        # 脚本自身异常（契约引擎崩溃）按「需人工确认」返回 3，不当业务不合格 1。
        print(f"ERROR: 契约检查脚本自身异常: {type(exc).__name__}: {exc}", file=sys.stderr)
        return engine.E_ASK
    # 数字溯源被跳过（无来源转储）= 没验完，不许当"全过"：升级为需用户确认（exit 3）
    if code == engine.E_OK and any(c.id == "CHK-NUM-1" and c.status == "skip" for c in checks):
        code = engine.E_ASK
    payload = build_payload(ctx, checks, manual, code)
    out_dir = args.out or os.path.join(ctx["project_dir"], "05_过程", "契约检查")
    try:
        md_path, _ = write_outputs(out_dir, payload)
    except OSError as exc:
        print("ERROR: 无法写入契约检查产物（可能是只读目录、无权限或输出路径被占用）："
              f"{lib.rel_posix(os.path.abspath(out_dir), ctx['project_dir'])}", file=sys.stderr)
        print(f"       原因：{type(exc).__name__}: {safe_error_text(str(exc), ctx['project_dir'])}",
              file=sys.stderr)
        print("       未写出差异清单；请换可写目录（--out）或调整权限后重跑。", file=sys.stderr)
        return engine.E_FAIL
    s = payload["summary"]
    print(f"[contracts] exit={code} fail={s['fail']} hint={s['hint']} skip={s['skip']} "
          f"manual={s['manual']} direction={ctx['direction'] or '-'}")
    print(f"[contracts] out={lib.rel_posix(md_path, ctx['project_dir'])}")
    for c in checks:
        if c.status == "fail":
            print(f"[contracts] FAIL {c.id}: {c.items[0]}")
    if args.json:
        import json
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    return code


if __name__ == "__main__":
    sys.exit(main())
