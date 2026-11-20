#!/usr/bin/env python3
"""扫描项目投递区，按分类规则判定类别，生成归档计划（默认 dry-run，不动任何文件）。

用法（T20 起，推荐用户目录口径；--root/--project 只用于校准/回归）：
  python classify_materials.py --project-dir "<用户自选项目目录>"
  python classify_materials.py --root "<项目库根>" --project "<项目名>"      # 标准项目库布局
  python classify_materials.py ... --apply        # 复制到 02_分类材料/04_成果报告 等去向
  python classify_materials.py ... --apply --move # 移动原件（必须用户明确授权）

设计：
- 默认只读：扫描 + 判定 + 写 05_过程\\分类清单\\分类清单.{md,json}，不碰原始材料。
- 默认复制不移动；--move 才移动，且必须同时给 --apply。
- 不确定就标「待确认」，不自动归档；目标已存在且内容不同 → 不覆盖，记为冲突。
- 零硬编码：根目录/项目名/规则全部来自参数或配置；脚本内无盘符、无项目名。
- 产物不写盘符：用户目录模式下清单里的目录/规则一律写相对路径或 <技能目录>/… 占位。
"""
from __future__ import annotations

import argparse
import fnmatch
import json
import os
import re
import shutil
import sys
import unicodedata
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _common import (  # noqa: E402
    RULES_FILE,
    UserProjectError,
    direction_names,
    disp_path,
    format_size,
    gate,
    load_yaml_file,
    now_str,
    portable_display,
    read_text,
    rel_posix,
    resolve_project_dir,
    resolve_project_name,
    resolve_root,
    resolve_user_project_dir,
    setup_stdout,
    sha256_file,
    write_text,
    group_name_aliases,
)

INDEX_NAME = "00_项目索引.md"
MANIFEST_MD = "分类清单.md"
MANIFEST_JSON = "分类清单.json"
DELIVERY_LINE = "投递区实际路径"
MAX_SAMPLE_BYTES = 2_000_000
CONTENT_MAX_HITS = 3
WEIGHT_FILENAME = 3.0
WEIGHT_TAIL = 4.0
WEIGHT_PATH = 8.0
WEIGHT_EXT = 2.0
WEIGHT_CONTENT = 2.0


def _disp_width(text: str) -> int:
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)


def _pad(text: str, width: int) -> str:
    return text + " " * max(0, width - _disp_width(text))


def _truncate_disp(text: str, max_width: int) -> str:
    if _disp_width(text) <= max_width:
        return text
    out: list[str] = []
    used = 0
    for ch in text:
        ch_width = 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
        if used + ch_width > max_width - 1:
            break
        out.append(ch)
        used += ch_width
    return "".join(out) + "…"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="扫描投递区，按规则生成归档计划（默认 dry-run）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--project-dir", help="用户自选项目目录（最高优先级；任意路径）")
    parser.add_argument("--root", help="[校准/回归] 标准项目库根目录；缺省读环境变量 CARDO_LIB_ROOT 或 项目库配置.yaml")
    parser.add_argument("--project", help="[校准/回归] 项目文件夹名")
    parser.add_argument("--direction", help="方向目录名（仅 --root 布局参与路径；用户目录模式只做标注）")
    parser.add_argument("--config", help=f"分类规则文件，默认 {RULES_FILE}")
    parser.add_argument("--delivery", help="投递区目录（相对项目根）；缺省按索引记录，其次按候选名探测")
    parser.add_argument("--out", help="分类清单输出目录（相对项目根或绝对路径），默认 05_过程/分类清单")
    parser.add_argument("--apply", action="store_true", help="按计划复制文件（默认不复制）；原件保留")
    parser.add_argument("--move", action="store_true", help="移动原件（危险，必须用户明确授权；需同时给 --apply）")
    parser.add_argument("--dry-run", action="store_true", help="显式声明 dry-run（默认行为）")
    return parser.parse_args(argv)


# ---------------------------------------------------------------- 投递区解析


def resolve_delivery(project_dir: Path, arg_delivery: str | None, candidates: list[str]):
    """返回 (delivery_path, delivery_label, gate_msgs, note)。"""
    index_path = project_dir / INDEX_NAME
    index_text = read_text(index_path) if index_path.is_file() else None
    recorded = None
    if index_text:
        for pattern in (r"投递区实际路径[^\n]*?`([^`\n]+)`", r"投递区实际路径\s*[|:：]\s*([^\n|]+)"):
            match = re.search(pattern, index_text)
            if match:
                recorded = match.group(1).strip().strip("`").strip()
                if recorded:
                    break

    if arg_delivery:
        path = (project_dir / arg_delivery.strip().strip("/\\")).resolve()
        if not path.is_dir():
            return None, arg_delivery, [f"--delivery 指定的投递区不存在：{path}"], ""
        note = ""
        if recorded and recorded.strip("/\\") != arg_delivery.strip("/\\"):
            note = f"注意：命令行投递区 {arg_delivery} 与索引记录 {recorded} 不一致，本次以命令行为准。"
        return path, arg_delivery, [], note

    if not index_path.is_file():
        return None, "", [
            f"项目目录存在但缺少 {INDEX_NAME}，无法确认投递区与材料台账归属（疑似新项目未建档）。",
            "请先确认项目是否已建档：新项目先运行 scaffold_project.py 建档；已有项目补索引后再分类。",
        ], ""

    if recorded:
        path = (project_dir / recorded.strip("/\\")).resolve()
        if path.is_dir():
            return path, recorded, [], ""
        return None, recorded, [
            f"索引记录的投递区路径不存在：{recorded}",
            "请确认投递区是否被移动/改名，或索引记录写错；确认后再分类（不要擅自新建投递区）。",
        ], ""

    found = [c for c in candidates if (project_dir / c).is_dir()]
    groups = group_name_aliases(project_dir, found)
    if len(groups) == 1:
        names = groups[0]
        note = ""
        if len(names) > 1:
            note = (f"投递区别名：{'、'.join(names)} 指向同一目录，本次按 {names[0]} 处理"
                    "（编号位与用户原名等价，不必二选一）")
        return project_dir / names[0], names[0], [], note
    if len(groups) > 1:
        return None, "", [
            f"索引未记录投递区，且项目下同时存在多个候选（互不相同的目录）：{'、'.join(found)}",
            "无法判断哪个是用户真正投递区；请确认后用 --delivery <目录名> 指定，并写进索引。",
        ], ""
    return None, "", [
        f"索引未记录投递区，且找不到任何候选目录（{' / '.join(candidates)}）。",
        "请确认用户材料放在哪个目录；确认后补写索引或用 --delivery 指定。",
    ], ""


# ---------------------------------------------------------------- 扫描 / 取样


def iter_delivery_files(delivery: Path, ignore_globs: list[str]):
    for dirpath, dirnames, filenames in os.walk(delivery):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        for name in sorted(filenames):
            if any(fnmatch.fnmatch(name, pattern) for pattern in ignore_globs):
                continue
            path = Path(dirpath) / name
            if not path.is_file():
                continue
            rel = rel_posix(path, delivery)
            yield path, rel


def _xml_text(data: bytes) -> str:
    return re.sub(r"<[^>]+>", " ", data.decode("utf-8", "ignore"))


def read_content_sample(path: Path, ext: str) -> tuple[str, str]:
    """返回 (可读文本, 说明)。读不到不算错，说明会写进清单。"""
    try:
        if ext == "docx":
            with zipfile.ZipFile(path) as zf:
                names = [n for n in zf.namelist() if n in ("word/document.xml", "word/footnotes.xml")]
                if not names:
                    return "", "docx 内无 document.xml，未读正文"
                chunks = [_xml_text(zf.read(n)[:MAX_SAMPLE_BYTES]) for n in names]
            return " ".join(chunks), "docx 正文"
        if ext == "xlsx":
            with zipfile.ZipFile(path) as zf:
                names = [n for n in zf.namelist() if n == "xl/sharedStrings.xml"]
                if not names:
                    return "", "xlsx 无共享字符串，未读正文"
                return _xml_text(zf.read(names[0])[:MAX_SAMPLE_BYTES]), "xlsx 共享字符串"
        if ext == "pptx":
            with zipfile.ZipFile(path) as zf:
                names = sorted(n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n))[:20]
                return " ".join(_xml_text(zf.read(n)) for n in names), "pptx 幻灯片文本"
        if ext in ("txt", "md", "csv", "json"):
            return path.read_text(encoding="utf-8", errors="ignore")[:MAX_SAMPLE_BYTES], f"{ext} 文本"
        if ext == "pdf":
            return "", "pdf 未解析（本机无 pypdf/pdfminer），仅按文件名与路径判定；扫描件无法读文本"
        if ext in ("doc", "xls", "ppt", "wps", "dwg", "dxf"):
            return "", f"{ext} 旧版二进制格式，未读正文"
        return "", f"{ext or '无扩展名'} 未读取正文"
    except (zipfile.BadZipFile, OSError, RuntimeError) as exc:
        return "", f"读取失败（{type(exc).__name__}），仅按文件名与路径判定"


# ---------------------------------------------------------------- 判定


def category_map(categories: list[dict]) -> dict[str, dict]:
    return {str(cat.get("id")): cat for cat in categories if cat.get("id")}


def classify_one(path: Path, rel: str, rules: dict, categories: list[dict]) -> dict:
    ext = path.suffix.lower().lstrip(".")
    stem = path.stem
    size = path.stat().st_size
    sample, content_note = read_content_sample(path, ext)
    sample_lower = sample.lower()
    priorities = {str(c["id"]): float(c.get("priority") or 0) for c in categories}
    scored: list[dict] = []
    tail_end: dict[str, int] = {}
    tail_kw: dict[str, str] = {}

    for cat in categories:
        cat_id = str(cat.get("id"))
        ext_list = [str(e).lower().lstrip(".") for e in (cat.get("ext") or [])]
        if ext_list and ext not in ext_list:
            continue
        score = 0.0
        evidence: list[str] = []
        if ext_list:
            score += WEIGHT_EXT
            evidence.append(f"扩展名 .{ext}")

        fname_hits: list[str] = []
        for kw in [str(k) for k in (cat.get("filename_keywords") or []) if str(k).strip()]:
            if kw.lower() in stem.lower():
                fname_hits.append(kw)
        if fname_hits:
            score += WEIGHT_FILENAME * len(fname_hits)
            evidence.append("文件名命中：" + "、".join(fname_hits))
            end_map = {kw: stem.lower().rfind(kw.lower()) + len(kw) for kw in fname_hits}
            tail_end[cat_id] = max(end_map.values())
            tail_kw[cat_id] = max(end_map, key=lambda k: end_map[k])

        path_hits = [str(k) for k in (cat.get("path_keywords") or []) if str(k).strip() and str(k) in rel]
        if path_hits:
            score += WEIGHT_PATH * len(path_hits)
            evidence.append("路径命中：" + "、".join(path_hits))

        if sample:
            content_hits = []
            for kw in [str(k) for k in (cat.get("content_keywords") or []) if str(k).strip()]:
                if kw.lower() in sample_lower:
                    content_hits.append(kw)
                if len(content_hits) >= CONTENT_MAX_HITS:
                    break
            if content_hits:
                score += WEIGHT_CONTENT * len(content_hits)
                evidence.append("正文命中：" + "、".join(content_hits))

        if score > 0:
            scored.append(
                {
                    "id": cat_id,
                    "label": str(cat.get("label") or cat_id),
                    "dir": str(cat.get("dir") or ""),
                    "priority": priorities.get(cat_id, 0.0),
                    "score": score,
                    "evidence": evidence,
                    "preserve_subdir": bool(cat.get("preserve_subdir")),
                }
            )

    # 中心词加成：文件名里“最靠后”的命中词所在类别额外加分（中文文件名中心词多在末尾）
    if tail_end:
        best_cat = max(
            tail_end,
            key=lambda c: (tail_end[c], len(tail_kw.get(c, "")), priorities.get(c, 0.0)),
        )
        for item in scored:
            if item["id"] == best_cat:
                item["score"] += WEIGHT_TAIL
                item["evidence"].append(f"中心词「{tail_kw[best_cat]}」位于文件名末尾")

    scored.sort(key=lambda item: (-item["score"], -item["priority"]))
    thresholds = rules.get("thresholds") or {}
    min_score = float(thresholds.get("auto_min_score", 3.0))
    conf_threshold = float(thresholds.get("confidence_threshold", 0.55))
    gap_threshold = float(thresholds.get("confidence_gap", 0.15))

    result: dict = {
        "rel": rel,
        "abs": str(path),
        "size": size,
        "size_h": format_size(size),
        "ext": ext,
        "content_note": content_note,
        "score": 0.0,
        "confidence": 0.0,
        "gap": 0.0,
        "category_id": "",
        "category_label": "",
        "pending": False,
        "pending_reason": "",
        "pending_candidates": [],
        "evidence": [],
        "dest_rel": "",
    }

    if not scored:
        result.update(
            pending=True,
            pending_reason="没有任何类别命中（规则表可能缺这一类，或文件名太泛）",
            pending_candidates=[{"label": str(c.get("label")), "score": 0.0} for c in categories],
            evidence=["无匹配：扩展名/文件名/路径/正文均未命中规则"],
        )
        return result

    top = scored[0]
    second = scored[1] if len(scored) > 1 else None
    top_score = top["score"]
    second_score = second["score"] if second else 0.0
    confidence = top_score / (top_score + second_score) if second_score > 0 else 0.95
    gap = (top_score - second_score) / top_score if top_score > 0 else 0.0
    result.update(
        score=top_score,
        confidence=confidence,
        gap=gap,
        category_id=top["id"],
        category_label=top["label"],
        evidence=list(top["evidence"]),
        pending_candidates=[{"label": s["label"], "score": s["score"]} for s in scored[:3]],
    )

    reasons: list[str] = []
    if top_score < min_score:
        reasons.append(f"最高分 {top_score:g} < 阈值 {min_score:g}")
    if confidence < conf_threshold:
        reasons.append(f"置信度 {confidence:.2f} < 阈值 {conf_threshold:.2f}")
    if second and gap < gap_threshold:
        reasons.append(f"与次优类别「{second['label']}」差距 {gap:.2f} < 阈值 {gap_threshold:.2f}")
    if top["id"] == "other":
        reasons.append("只命中兜底类别「其他」")
    if reasons:
        result["pending"] = True
        result["pending_reason"] = "；".join(reasons)
    result["_top"] = top  # 供去向计算
    return result


# ---------------------------------------------------------------- 去向


def flatten_parent(rel: str) -> str:
    parent = Path(rel).parent.as_posix()
    if parent in ("", "."):
        return ""
    return parent.replace("/", "-")


def compute_destinations(project_dir: Path, results: list[dict], rules: dict) -> None:
    """就地填写 dest_rel；处理重名与目标已存在冲突。"""
    prelim: list[tuple[dict, str, str]] = []  # (result, dest_dir_rel, name)
    for item in results:
        top = item.get("_top")
        if not top or item["pending"]:
            continue
        dest_dir = top["dir"].strip("/")
        if top.get("preserve_subdir"):
            parent = flatten_parent(item["rel"])
            if parent:
                dest_dir = f"{dest_dir}/{parent}"
        prelim.append((item, dest_dir, Path(item["rel"]).name))

    counts: dict[tuple[str, str], int] = {}
    for _, dest_dir, name in prelim:
        key = (dest_dir, name)
        counts[key] = counts.get(key, 0) + 1

    for item, dest_dir, name in prelim:
        key = (dest_dir, name)
        if counts[key] > 1:
            parent = flatten_parent(item["rel"])
            name = f"{parent}__{name}" if parent else name
        dest_rel = f"{dest_dir}/{name}"
        dest_abs = project_dir / dest_rel
        item["dest_rel"] = dest_rel
        if dest_abs.exists():
            same = dest_abs.is_file() and dest_abs.stat().st_size == item["size"] and sha256_file(dest_abs) == sha256_file(item["abs"])
            if same:
                item["pending"] = True
                item["pending_reason"] = "目标已存在且内容相同（重复执行，无需再复制）"
                item["pending_candidates"] = [{"label": item["category_label"], "score": item["score"]}]
            else:
                item["pending"] = True
                item["pending_reason"] = f"目标已存在且内容不同，拒绝覆盖：{dest_rel}"
                item["pending_candidates"] = [{"label": item["category_label"], "score": item["score"]}]

    for item in results:
        item.pop("_top", None)


# ---------------------------------------------------------------- 产物


def build_summary(results: list[dict], categories: list[dict]) -> dict:
    by_category: dict[str, int] = {}
    pending = [r for r in results if r["pending"]]
    for item in results:
        if not item["pending"] and item["category_label"]:
            by_category[item["category_label"]] = by_category.get(item["category_label"], 0) + 1
    return {
        "total": len(results),
        "auto": len(results) - len(pending),
        "pending": len(pending),
        "by_category": by_category,
    }


def esc(text: str) -> str:
    return str(text).replace("|", "\\|").replace("\n", " ")


def write_manifest_md(path: Path, meta: dict, results: list[dict], summary: dict) -> None:
    mode_text = meta["mode_text"]
    lines: list[str] = []
    lines.append(f"# 分类清单 · {meta['project']}")
    lines.append("")
    lines.append("> 由 `cardo-project-material/scripts/classify_materials.py` 生成，可重复执行覆盖。")
    lines.append("> 这是**归档计划**，不是原始材料台账；原件始终留在投递区。")
    lines.append("")
    lines.append("## 概况")
    lines.append("")
    lines.append("| 项 | 值 |")
    lines.append("|---|---|")
    lines.append(f"| 生成时间 | {meta['generated_at']} |")
    lines.append(f"| 项目 | {esc(meta['project'])} |")
    lines.append(f"| 投递区 | `{esc(meta['delivery'])}` |")
    lines.append(f"| 规则文件 | {esc(meta['rules_path'])} |")
    lines.append(f"| 模式 | {mode_text} |")
    lines.append(f"| 文件总数 | {summary['total']} |")
    lines.append(f"| 自动归类 | {summary['auto']} |")
    lines.append(f"| 待确认 | {summary['pending']} |")
    by_cat = "；".join(f"{label} {count}" for label, count in summary["by_category"].items()) or "—"
    lines.append(f"| 按类别 | {by_cat} |")
    lines.append("")
    lines.append("## 归档计划")
    lines.append("")
    header = "| # | 文件（相对投递区） | 大小 | 初判类别 | 判定依据 | 置信度 | 建议去向 | 待确认？ |"
    lines.append(header)
    lines.append("|---|---|---|---|---|---|---|---|")
    for idx, item in enumerate(results, start=1):
        dest = item["dest_rel"] or "（未定）"
        pending = "是" if item["pending"] else ""
        conf = f"{item['confidence']:.2f}" if item["category_label"] else "—"
        lines.append(
            f"| {idx} | {esc(item['rel'])} | {item['size_h']} | {esc(item['category_label'] or '待定')} | "
            f"{esc('；'.join(item['evidence']))} | {conf} | {esc(dest)} | {pending} |"
        )
    lines.append("")
    lines.append("## 待确认（必须先问用户，不能自动归档）")
    lines.append("")
    pending_items = [r for r in results if r["pending"]]
    if pending_items:
        lines.append("| 编号 | 文件 | 现象 | 候选项 | 影响 | 建议 |")
        lines.append("|---|---|---|---|---|---|")
        for idx, item in enumerate(pending_items, start=1):
            candidates = "；".join(f"{c['label']}（{c['score']:g}）" for c in item["pending_candidates"][:3]) or "—"
            lines.append(
                f"| P{idx:02d} | {esc(item['rel'])} | {esc(item['pending_reason'])} | {esc(candidates)} | "
                f"文件无法自动归档或去向不确定 | 请用户确认类别；确认后更新规则或用 --delivery/手工归档 |"
            )
    else:
        lines.append("无。全部文件均已按规则自动判定。")
    lines.append("")
    lines.append("## 模式说明")
    lines.append("")
    lines.append("- 默认 **dry-run**：只生成计划，不复制、不移动任何文件。")
    lines.append("- `--apply`：按计划**复制**到建议去向，原件保留在投递区。")
    lines.append("- `--move`：移动原件（危险），必须同时给 `--apply`，且需用户明确授权。")
    lines.append("- 目标已存在且内容不同 → 拒绝覆盖，记为待确认。")
    lines.append("")
    lines.append("## 正文可读性说明")
    lines.append("")
    for item in results:
        if item["content_note"] and "未读" in item["content_note"] or "未解析" in item["content_note"]:
            lines.append(f"- `{item['rel']}`：{item['content_note']}")
    if not any(("未读" in r["content_note"] or "未解析" in r["content_note"]) for r in results):
        lines.append("- 全部文件均可读正文（或无需读正文）。")
    lines.append("")
    write_text(path, "\n".join(lines))


def write_manifest_json(path: Path, meta: dict, results: list[dict], summary: dict,
                        path_base: Path | None = None) -> None:
    def file_row(item: dict) -> dict:
        row = {k: v for k, v in item.items() if not k.startswith("_")}
        if path_base is not None:
            try:
                row["abs"] = rel_posix(Path(row["abs"]), path_base)
            except (OSError, ValueError):
                row["abs"] = item["rel"]
        return row

    payload = {
        "schema": 1,
        "generated_at": meta["generated_at"],
        "project": meta["project"],
        "project_dir": str(meta["project_dir"]),
        "delivery": meta["delivery"],
        "delivery_abs": str(meta["delivery_abs"]),
        "rules_path": str(meta["rules_path"]),
        "mode": meta["mode"],
        "thresholds": meta["thresholds"],
        "summary": summary,
        "files": [file_row(item) for item in results],
        "pending": [
            {
                "no": f"P{idx:02d}",
                "file": item["rel"],
                "reason": item["pending_reason"],
                "candidates": item["pending_candidates"][:3],
                "suggestion": item["dest_rel"] or "（待确认后确定）",
            }
            for idx, item in enumerate([r for r in results if r["pending"]], start=1)
        ],
        "actions_executed": meta.get("executed", []),
    }
    write_text(path, json.dumps(payload, ensure_ascii=False, indent=2) + "\n")


# ---------------------------------------------------------------- 执行


def execute_actions(project_dir: Path, results: list[dict], move: bool) -> list[dict]:
    executed: list[dict] = []
    for item in results:
        if item["pending"] or not item["dest_rel"]:
            continue
        src = Path(item["abs"])
        dest = project_dir / item["dest_rel"]
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            if move:
                shutil.move(str(src), str(dest))
                status = "已移动"
            else:
                shutil.copy2(src, dest)
                status = "已复制"
        except OSError as exc:
            status = f"失败：{exc}"
        executed.append({"file": item["rel"], "dest": item["dest_rel"], "status": status})
    return executed


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
        project = project_dir.name
    else:
        root = resolve_root(args.root)
        project = resolve_project_name(args.project)
        project_dir = resolve_project_dir(root, project, args.direction)
    rules_path = Path(args.config).expanduser().resolve() if args.config else RULES_FILE

    if args.move and not args.apply:
        print("错误：--move 必须与 --apply 同时使用（--move 的含义是移动原件，需用户明确授权）。")
        return 1
    if not user_mode and project_dir is None:
        return gate(
            [
                f"未找到项目目录：{project}"
                f"（已按方向 {'/'.join(direction_names())} 与旧扁平布局查找）",
                "确认项目名/方向是否正确；若是新项目，先运行 scaffold_project.py 建档（不会自动创建）。",
            ]
        )
    assert project_dir is not None
    if not rules_path.is_file():
        print(f"错误：分类规则文件不存在：{rules_path}")
        return 1

    rules = load_yaml_file(rules_path)
    categories = [c for c in (rules.get("categories") or []) if isinstance(c, dict)]
    if not categories:
        print("错误：分类规则文件里 categories 为空。")
        return 1
    for cat in categories:
        for key in ("id", "label", "dir"):
            if not cat.get(key):
                print(f"错误：类别配置缺少 {key}：{cat}")
                return 1

    candidates = [str(c) for c in (rules.get("delivery_candidates") or []) if str(c).strip()]
    if not candidates:
        print("错误：分类规则文件里 delivery_candidates 为空。")
        return 1
    thresholds = rules.get("thresholds") or {}
    dirs_cfg = rules.get("dirs") or {}
    ignore_globs = [str(p) for p in (rules.get("ignore") or [])]

    delivery_path, delivery_label, gate_msgs, delivery_note = resolve_delivery(project_dir, args.delivery, candidates)
    if gate_msgs:
        return gate(gate_msgs)
    assert delivery_path is not None
    if delivery_note:
        print(f"[提示] {delivery_note}")

    files = list(iter_delivery_files(delivery_path, ignore_globs))
    if not files:
        print(f"投递区为空：{delivery_path}")
        print("（没有可分类文件；请确认用户是否还没投递材料。）")
        return 0

    results = [classify_one(path, rel, rules, categories) for path, rel in files]
    compute_destinations(project_dir, results, rules)
    summary = build_summary(results, categories)

    mode = "apply-move" if args.move else ("apply-copy" if args.apply else "dry-run")
    mode_text = {
        "dry-run": "dry-run（未复制、未移动任何文件）",
        "apply-copy": "apply（复制，原件保留）",
        "apply-move": "apply + move（移动原件）",
    }[mode]

    out_dir = Path(args.out).expanduser() if args.out else Path(dirs_cfg.get("manifest_dir") or "05_过程/分类清单")
    if not out_dir.is_absolute():
        out_dir = project_dir / out_dir
    manifest_md = out_dir / MANIFEST_MD
    manifest_json = out_dir / MANIFEST_JSON

    meta = {
        "generated_at": now_str(),
        "project": project,
        "project_dir": project_dir if not user_mode else ".",
        "delivery": delivery_label,
        "delivery_abs": delivery_path if not user_mode else rel_posix(delivery_path, project_dir),
        "rules_path": rules_path if not user_mode else portable_display(rules_path),
        "mode": mode,
        "mode_text": mode_text,
        "thresholds": {
            "auto_min_score": float(thresholds.get("auto_min_score", 3.0)),
            "confidence_threshold": float(thresholds.get("confidence_threshold", 0.55)),
            "confidence_gap": float(thresholds.get("confidence_gap", 0.15)),
        },
    }

    if args.apply:
        executed = execute_actions(project_dir, results, move=args.move)
        meta["executed"] = executed
        summary = build_summary(results, categories)

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        write_manifest_md(manifest_md, meta, results, summary)
        # 用户目录模式：清单里的文件路径一律相对项目目录（不写盘符绝对路径）
        write_manifest_json(manifest_json, meta, results, summary,
                            path_base=project_dir if user_mode else None)
    except OSError as exc:
        print(f"错误：无法写入分类清单（可能是只读目录、无权限或文件被占用）：{disp_path(out_dir)}")
        print(f"      原因：{type(exc).__name__}: {exc}")
        return 1

    # ---- 控制台摘要 ----
    print(f"项目：{project}")
    print(f"投递区：{delivery_label}（{disp_path(delivery_path) if user_mode else delivery_path}）")
    print(f"模式：{mode_text}")
    print(f"文件总数 {summary['total']}：自动归类 {summary['auto']}，待确认 {summary['pending']}")
    if summary["by_category"]:
        print("按类别：" + "；".join(f"{k} {v}" for k, v in summary["by_category"].items()))
    print("")
    print(f"{'#':>2}  {_pad('文件', 44)} {'大小':>9}  {_pad('初判类别', 18)} {'置信度':>5}  去向")
    for idx, item in enumerate(results, start=1):
        name = _truncate_disp(item["rel"], 44)
        label = _truncate_disp(item["category_label"] or "待确认", 18)
        conf = f"{item['confidence']:.2f}" if item["category_label"] else "  — "
        print(
            f"{idx:>2}  {_pad(name, 44)} {item['size_h']:>9}  {_pad(label, 18)} {conf:>5}  "
            f"{item['dest_rel'] or '（待确认）'}"
        )
    if summary["pending"]:
        print("")
        print(f"[闸门] 有 {summary['pending']} 个文件无法安全自动归类，必须先问用户：")
        for item in [r for r in results if r["pending"]]:
            print(f"  - {item['rel']}：{item['pending_reason']}")
    if args.apply:
        print("")
        print(f"已执行 {len(meta.get('executed', []))} 个文件动作（{'移动' if args.move else '复制'}）。原件"
              + ("已移动，投递区对应文件不再存在。" if args.move else "保留在投递区。"))
    print("")
    print(f"分类清单：{disp_path(manifest_md) if user_mode else manifest_md}")
    print(f"机器可读：{disp_path(manifest_json) if user_mode else manifest_json}")
    if not args.apply:
        print("下一步：确认计划无误且用户授权后，再运行 --apply（复制）；移动原件需 --apply --move。")
    return 0


if __name__ == "__main__":
    sys.exit(main())