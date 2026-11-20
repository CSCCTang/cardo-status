#!/usr/bin/env python3
"""把分类清单与产物台账写入 00_项目索引.md 的固定锚点区块（幂等，不重复追加）。

用法（T20 起，推荐用户目录口径；--root/--project 只用于校准/回归）：
  python update_index.py --project-dir "<用户自选项目目录>"
  python update_index.py --root "<项目库根>" --project "<项目名>"      # 标准项目库布局
  python update_index.py ... --manifest "05_过程/分类清单/分类清单.json"
  python update_index.py ... --artifact "标准现状分析报告|04_成果报告/xxx.docx|2026-10-21|已产出"
  python update_index.py ... --dry-run        # 只打印将写入的区块

规则：
- 只改 <!-- BEGIN:xxx --> 与 <!-- END:xxx --> 之间的内容；区块外的文字一律不动。
- 不传 --manifest 就不动材料台账，不传 --artifact 就不动产物台账，便于分步更新。
- 内容无变化时不写文件（重复执行不追加、不产生无意义改动）。
- 退出码：0 成功；3 索引/锚点缺失（需先补索引或问用户）；1 其他错误（含只读目录）。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _common import (  # noqa: E402
    RULES_FILE,
    UserProjectError,
    direction_names,
    disp_path,
    gate,
    load_yaml_file,
    read_text,
    resolve_project_dir,
    resolve_project_name,
    resolve_root,
    resolve_user_project_dir,
    setup_stdout,
    write_text,
)

INDEX_NAME = "00_项目索引.md"
MANIFEST_REL = "05_过程/分类清单/分类清单.json"
BLOCKS = {
    "材料台账": ("<!-- BEGIN:材料台账 -->", "<!-- END:材料台账 -->"),
    "产物台账": ("<!-- BEGIN:产物台账 -->", "<!-- END:产物台账 -->"),
    "待确认事项": ("<!-- BEGIN:待确认事项 -->", "<!-- END:待确认事项 -->"),
    "材料齐套": ("<!-- BEGIN:材料齐套 -->", "<!-- END:材料齐套 -->"),
}


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="把分类清单/产物写入项目索引的固定区块（幂等）")
    parser.add_argument("--project-dir", help="用户自选项目目录（最高优先级；任意路径）")
    parser.add_argument("--root", help="[校准/回归] 标准项目库根目录；缺省读环境变量 CARDO_LIB_ROOT 或 项目库配置.yaml")
    parser.add_argument("--project", help="[校准/回归] 项目文件夹名")
    parser.add_argument("--direction", help="[校准/回归] 方向目录名（用户目录模式不参与路径）")
    parser.add_argument(
        "--pending-item",
        action="append",
        default=[],
        help='待确认事项条目，格式 "编号|现象|候选/建议|影响"，可重复；与分类清单的待确认合并去重',
    )
    parser.add_argument(
        "--gap-file",
        help="材料齐套表 JSON 路径（check_materials.py --format json 输出），写入「材料齐套」区块",
    )
    parser.add_argument("--config", help=f"分类规则文件（用于读取清单默认目录），默认 {RULES_FILE}")
    parser.add_argument("--manifest", help=f"分类清单 JSON 路径（相对项目根或绝对），默认 {MANIFEST_REL}")
    parser.add_argument(
        "--artifact",
        action="append",
        default=[],
        help='产物台账行，格式 "名称|路径|日期|状态"，可重复；不给则产物台账保持原样',
    )
    parser.add_argument("--dry-run", action="store_true", help="只打印将写入的区块，不写文件")
    return parser.parse_args(argv)


def manifest_default(project_dir: Path, arg_config: str | None) -> Path:
    rel = MANIFEST_REL
    config_path = Path(arg_config).expanduser().resolve() if arg_config else RULES_FILE
    if config_path.is_file():
        try:
            rules = load_yaml_file(config_path)
            configured = (rules.get("dirs") or {}).get("manifest_dir")
            if isinstance(configured, str) and configured.strip():
                rel = configured.strip().strip("/\\") + "/分类清单.json"
        except SystemExit:
            pass
    return project_dir / rel


def esc(text: object) -> str:
    return str(text).replace("|", "\\|").replace("\n", " ")


def material_ledger(manifest: dict) -> str:
    summary = manifest.get("summary") or {}
    executed = {a.get("file"): a.get("status", "") for a in manifest.get("actions_executed") or []}
    files = manifest.get("files") or []
    lines = [
        f"> 投递区：`{esc(manifest.get('delivery', ''))}` ｜ 扫描时间：{manifest.get('generated_at', '')} ｜ "
        f"文件 {summary.get('total', len(files))}（自动 {summary.get('auto', 0)} / 待确认 {summary.get('pending', 0)}）",
        "",
        "| # | 文件（相对投递区） | 大小 | 初判类别 | 置信度 | 建议去向 | 状态 |",
        "|---|---|---|---|---|---|---|",
    ]
    for idx, item in enumerate(files, start=1):
        if item.get("pending"):
            status = "待确认"
        elif executed.get(item.get("rel")):
            status = executed[item["rel"]]
        else:
            status = "待归档"
        confidence = f"{float(item.get('confidence') or 0):.2f}" if item.get("category_label") else "—"
        lines.append(
            f"| {idx} | {esc(item.get('rel'))} | {esc(item.get('size_h'))} | {esc(item.get('category_label') or '待定')} | "
            f"{confidence} | {esc(item.get('dest_rel') or '（待确认）')} | {status} |"
        )
    if not files:
        lines.append("| — | 暂无（未扫描或投递区为空） | — | — | — | — | — |")
    return "\n".join(lines)


def product_ledger(artifacts: list[str]) -> str:
    lines = [
        "| 产物 | 路径 | 日期 | 状态 |",
        "|---|---|---|---|",
    ]
    for entry in artifacts:
        parts = [part.strip() for part in entry.split("|")]
        parts += [""] * (4 - len(parts))
        name, path, date, status = parts[:4]
        lines.append(
            f"| {esc(name)} | {esc(path)} | {esc(date or '—')} | {esc(status or '待产出')} |"
        )
    if not artifacts:
        lines.append("| 标准现状分析报告 docx | `04_成果报告\\`（待产出） | — | 预留 |")
    return "\n".join(lines)


def gap_ledger(data: dict) -> str:
    """材料齐套表（check_materials.py 的 JSON）→ 索引区块。"""
    summary = data.get("summary") or {}
    items = data.get("items") or []
    req_t = summary.get("required_total", 0)
    req_h = summary.get("required_hit", 0)
    sug_t = summary.get("suggested_total", 0)
    sug_h = summary.get("suggested_hit", 0)
    raw_rate = summary.get("rate")
    rate = f"{float(raw_rate) * 100:.0f}%" if isinstance(raw_rate, (int, float)) else "—"
    lines = [
        f"> 方向：`{esc(data.get('direction') or '—')}` ｜ 投递区：`{esc(data.get('delivery') or '—')}` ｜ "
        f"检查时间：{esc(data.get('generated_at') or '')} ｜ "
        f"必需 {req_h}/{req_t} 齐、建议 {sug_h}/{sug_t} 齐（齐套率 {rate}）",
        "",
    ]
    missing_req = [i for i in items if i.get("level") == "必需" and i.get("status") != "已齐"]
    if missing_req:
        lines.append("**缺必需材料（补齐前不宜直接进入拆解，或由用户确认「先出稿、缺项入报告」）：**")
        lines.append("")
    lines += [
        "| 编号 | 材料 | 等级 | 影响章节 | 状态 | 建议来源 |",
        "|---|---|---|---|---|---|",
    ]
    for item in items:
        chapters = "、".join(item.get("chapters") or []) or "—"
        matched = item.get("matched") or []
        status = item.get("status") or "缺失"
        if status == "已齐" and matched:
            status = f"已齐（{len(matched)} 份）"
        lines.append(
            f"| {esc(item.get('id'))} | {esc(item.get('name'))} | {esc(item.get('level') or '建议')} | "
            f"{esc(chapters)} | {esc(status)} | {esc(item.get('source') or '—')} |"
        )
    if not items:
        lines.append("| — | 未检查（先运行 check_materials.py） | — | — | — | — |")
    return "\n".join(lines)


def pending_ledger(manifest: dict, manual: list[str] | None = None) -> str:
    """待确认事项：分类清单里的 pending + 手工传入的 --pending-item，按编号去重。"""
    rows: list[tuple[str, str, str, str, str]] = []
    seen: set[str] = set()
    for item in manifest.get("pending") or []:
        no = str(item.get("no") or "")
        candidates = "；".join(str(c.get("label", "")) for c in (item.get("candidates") or [])[:3]) or "—"
        rows.append((no, str(item.get("file") or "—"), str(item.get("reason") or ""),
                     candidates, "需用户拍板后才能归档/成文"))
        seen.add(no)
    for entry in manual or []:
        parts = [p.strip() for p in entry.split("|")]
        parts += [""] * (4 - len(parts))
        no, phenomenon, candidate, impact = parts[:4]
        if no in seen:
            continue
        seen.add(no)
        rows.append((no or f"P{len(rows) + 1:02d}", "—", phenomenon, candidate or "—", impact or "需用户拍板"))
    if not rows:
        return "（当前无待确认事项。）"
    lines = [
        "| 编号 | 文件 | 现象 | 候选/建议 | 影响 |",
        "|---|---|---|---|---|",
    ]
    for no, file, reason, candidates, impact in rows:
        lines.append(f"| {esc(no)} | {esc(file)} | {esc(reason)} | {esc(candidates)} | {esc(impact)} |")
    return "\n".join(lines)


def parse_manual_pending(text: str) -> list[str]:
    """从索引现有的「待确认事项」区块里取回手工条目（文件列为 —），避免被下次运行清掉。"""
    begin, end = BLOCKS["待确认事项"]
    m = re.search(re.escape(begin) + r"(.*?)" + re.escape(end), text, re.S)
    if not m:
        return []
    out: list[str] = []
    for line in m.group(1).splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        if len(cells) < 5 or set(cells[0]) <= {"-", ":"} or cells[0] == "编号":
            continue
        if cells[1].replace("\\", "") not in ("—", "-", "－"):
            continue
        out.append("|".join([cells[0], cells[2], cells[3], cells[4]]))
    return out


def replace_block(text: str, begin: str, end: str, body: str) -> tuple[str, bool]:
    pattern = re.compile(re.escape(begin) + r"(.*?)" + re.escape(end), re.S)
    match = pattern.search(text)
    if not match:
        return text, False
    new_segment = f"{begin}\n{body.strip()}\n{end}"
    if match.group(0) == new_segment:
        return text, True
    return text[: match.start()] + new_segment + text[match.end():], True


def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    args = parse_args(argv)
    # 用户目录模式：显式 --project-dir，或两者都不给（当前目录像项目就用它，否则 exit 2）
    user_mode = bool(str(args.project_dir or "").strip()) or not (args.root or args.project)
    if user_mode:
        # ---- 用户目录模式（T20）：项目目录任意，不依赖标准项目库 ----
        try:
            project_dir = resolve_user_project_dir(args.project_dir)
        except UserProjectError as exc:
            print(f"错误：{exc}")
            return 2
        _ = direction_names  # 用户目录模式方向不参与路径
    else:
        root = resolve_root(args.root)
        project = resolve_project_name(args.project)
        project_dir = resolve_project_dir(root, project, args.direction)
        if project_dir is None:
            return gate(
                [
                    f"未找到项目目录：{project}（已按方向 {'/'.join(direction_names())} 与旧扁平布局查找）",
                    "请先确认项目名/方向，或运行 scaffold_project.py 建档。",
                ]
            )
    index_path = project_dir / INDEX_NAME
    index_display = disp_path(index_path) if user_mode else str(index_path)
    if not index_path.is_file():
        return gate(
            [
                f"缺少项目索引：{index_display}",
                "疑似新项目未建档或索引被删。请确认后运行 scaffold_project.py --init-index 初始化，再更新台账。",
            ]
        )

    blocks: dict[str, str] = {}
    manifest: dict = {}
    loaded_manifest = False
    existing_text = read_text(index_path)
    # 手工待确认条目（--pending-item 写入的）跨运行保留，只有显式再次给出同编号才覆盖
    kept_manual = parse_manual_pending(existing_text)
    manual_all = [entry for entry in kept_manual if entry.split("|")[0] not in
                  {e.split("|")[0] for e in args.pending_item}] + list(args.pending_item)
    if args.manifest:
        manifest_path = Path(args.manifest).expanduser()
        if not manifest_path.is_absolute():
            manifest_path = project_dir / manifest_path
    else:
        manifest_path = manifest_default(project_dir, args.config)
    if args.manifest or manifest_path.is_file():
        if not manifest_path.is_file():
            print(f"错误：分类清单不存在：{manifest_path}")
            print("      先运行 classify_materials.py 生成分类清单（默认 dry-run 也会生成）。")
            return 1
        try:
            manifest = json.loads(read_text(manifest_path))
        except json.JSONDecodeError as exc:
            print(f"错误：分类清单 JSON 解析失败：{exc}")
            return 1
        loaded_manifest = True
        blocks["材料台账"] = material_ledger(manifest)
        blocks["待确认事项"] = pending_ledger(manifest)
    else:
        print(f"[提示] 未找到分类清单 {manifest_path}，本次不更新材料台账。")

    if loaded_manifest or manual_all:
        blocks["待确认事项"] = pending_ledger(manifest if loaded_manifest else {}, manual_all)

    if args.gap_file:
        gap_path = Path(args.gap_file).expanduser()
        if not gap_path.is_absolute():
            gap_path = project_dir / gap_path
        if not gap_path.is_file():
            print(f"错误：材料齐套表不存在：{gap_path}")
            print("      先运行 check_materials.py --format json 生成。")
            return 1
        try:
            blocks["材料齐套"] = gap_ledger(json.loads(read_text(gap_path)))
        except json.JSONDecodeError as exc:
            print(f"错误：材料齐套表 JSON 解析失败：{exc}")
            return 1

    if args.artifact:
        blocks["产物台账"] = product_ledger(args.artifact)

    if not blocks:
        print("没有任何要更新的区块（既无分类清单，也无 --artifact / --pending-item）。")
        return 0

    text = existing_text
    missing_anchors = [name for name in blocks if BLOCKS[name][0] not in text or BLOCKS[name][1] not in text]
    if missing_anchors:
        return gate(
            [
                f"{INDEX_NAME} 缺少锚点区块：{'、'.join(missing_anchors)}",
                "请先把各区块的 <!-- BEGIN:xxx --> / <!-- END:xxx --> 锚点补进索引（模板里有），再重跑；不要手工往区块外追加台账。",
            ]
        )

    updated = text
    changed: list[str] = []
    for name, body in blocks.items():
        begin, end = BLOCKS[name]
        updated, found = replace_block(updated, begin, end, body)
        if not found:
            print(f"错误：锚点区块异常：{name}")
            return 1
        if args.dry_run:
            print(f"----- {name}（将写入） -----")
            print(body)
            print("")

    if args.dry_run:
        print("（dry-run：未写入文件）")
        return 0

    if updated == text:
        print(f"索引无变化（重复执行不会重复追加）：{index_display}")
        return 0

    try:
        write_text(index_path, updated)
    except OSError as exc:
        print(f"错误：无法写入项目索引（可能是只读目录、无权限或文件被占用）：{index_display}")
        print(f"      原因：{type(exc).__name__}: {exc}")
        return 1
    for name in blocks:
        print(f"[更新] {name} → {index_display}")
    return 0


if __name__ == "__main__":
    sys.exit(main())