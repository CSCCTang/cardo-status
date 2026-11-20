#!/usr/bin/env python3
"""输入材料齐套性检查（只读）：投递区实际文件 ←→ 必备材料清单，输出齐套表与补齐提醒。

用法（T20 起，推荐用户目录口径；--root/--project 只用于校准/回归）：
  python check_materials.py --project-dir "<用户自选项目目录>" [--direction "<方向>"]
  python check_materials.py --root "<项目库根>" --project "<项目名>"      # 标准项目库布局
  python check_materials.py ... --format json                 # 控制台输出完整 JSON
  python check_materials.py ... --quiet                       # 只输出 JSON 摘要

只读承诺：
- 只扫描投递区（路径以 00_项目索引.md 的「投递区实际路径」为准）；
- 不移动/复制/改名/删除任何原始文件，也不往投递区写任何东西；
- 唯一写入位置是 <项目>\\05_过程\\齐套检查\\ 下的 材料齐套表.md / .json。
- 用户目录模式下齐套表内的目录/规则一律写相对路径或 <技能目录>/… 占位（不写盘符）。

产物 JSON（材料齐套表.json）的顶层结构固定为 cardo/material-gap/v1，
供 update_index.py --gap-file 直接消费：
  schema=cardo/material-gap/v1；project；direction；delivery（投递区相对路径）；
  generated_at；summary（required_total/required_hit/required_missing/
  suggested_total/suggested_hit/rate，rate 为 0~1 小数）；
  items[]（id/name/level/chapters/status/matched/source/why；
  status 只取「已齐 / 缺失 / 待确认」，顺序与清单一致：通用→本方向增补）。

判定口径：
- 材料清单见 规则/必备材料清单.yaml（通用 common + 方向增补 direction_extra）；
- 文件名命中 hints 每词 +3，父目录命中 hints 每词 +2，分类清单命中 match_category +6，
  弱线索 weak_hints 每词 +1；扩展名 ext 是硬门槛；
- 分类清单只认 01 建档与材料归集实际产出的那一份（过程文件/01-建档与材料归集/分类清单/分类清单.json），
  读不到才回落到旧布局；同一份输入只出一个日期口径的结论；
- 有强证据（文件名/父目录/分类清单）且得分 ≥3 → 已齐；只有弱线索 → 待确认-疑似；否则缺失；
- 「已齐」只说明「有这份材料」，不说明版本新、时效近（版本/时效另行确认）。

退出码：
  0 必需材料齐套；1 缺必需材料（含待确认，属提醒不是错误）；2 参数/输入错误；
  3 脚本自身异常（崩溃/环境问题，不是缺料判定）。
"""
from __future__ import annotations

import argparse
import fnmatch
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

# 兄弟模块导入失败属于执行错误：必须在模块级就以 exit 3 退出，
# 不能退化成 exit 1 被总控当成「缺材料」的业务结论。
try:
    import _common  # noqa: F401
    import classify_materials  # noqa: F401
except Exception as exc:
    import traceback
    traceback.print_exc()
    print(f"[齐套检查][内部错误] 无法导入同目录模块：{type(exc).__name__}: {exc}", file=sys.stderr)
    print("  这是脚本执行错误（打包缺失/环境问题），不是缺料判定。", file=sys.stderr)
    raise SystemExit(3) from None

from _common import (  # noqa: E402
    RULES_FILE,
    SKILL_DIR,
    UserProjectError,
    canon_direction,
    card_direction,
    default_direction,
    direction_names,
    disp_path,
    format_size,
    load_lib_config,
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
    write_text,
)
from classify_materials import resolve_delivery  # noqa: E402

INDEX_NAME = "00_项目索引.md"
MATERIALS_NAME = "必备材料清单.yaml"
CHECK_DIR_NAME = "齐套检查"
TABLE_MD = "材料齐套表.md"
TABLE_JSON = "材料齐套表.json"

LEVELS = ("必需", "建议", "可选")
STATUS_OK = "已齐"
STATUS_MISSING = "缺失"
STATUS_SUSPECT = "待确认-疑似"

# 模板六章（与标准模板一致；只用作「影响哪章」的展示，不含模板文件名）
CHAPTERS = (
    "项目概述",
    "片区基础概况分析",
    "现状详细分析",
    "核心问题汇总",
    "现状综合研判与更新思路",
    "附件",
)
CN_NUM = ("一", "二", "三", "四", "五", "六")
MAX_EVIDENCE_FILES = 6

WEIGHT_FILENAME = 3
WEIGHT_PATH = 2
WEIGHT_CATEGORY = 6
WEIGHT_WEAK = 1
STRONG_MIN_SCORE = 3


# ---------------------------------------------------------------- 参数 / 错误


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="输入材料齐套性检查（只读扫描投递区，输出齐套表与补齐提醒）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--project-dir", help="用户自选项目目录（最高优先级；任意路径）")
    parser.add_argument("--root", help="[校准/回归] 标准项目库根目录；缺省读环境变量 CARDO_LIB_ROOT 或 项目库配置.yaml")
    parser.add_argument("--project", help="[校准/回归] 项目文件夹名（不含路径）")
    parser.add_argument("--project-name", default="",
                        help="项目全称（如「××市××片区老旧小区改造项目」）；"
                             "缺省用项目目录名。仅用于报告抬头，不影响判定")
    parser.add_argument("--direction", help="方向（可选；缺省读 .cardo/project.json 或按项目所在方向目录推断）")
    parser.add_argument("--format", choices=("text", "json"), default="text", help="控制台输出格式，默认 text")
    parser.add_argument("--quiet", action="store_true", help="只输出 JSON 摘要（不打印补齐提醒长文）")
    parser.add_argument("--materials", help=f"必备材料清单文件，默认 规则/{MATERIALS_NAME}")
    parser.add_argument("--classification", default="",
                        help="01 建档产出的分类清单 JSON（相对项目目录或绝对路径）；缺省按 01 实际产物路径定位")
    parser.add_argument("--out", help=f"输出目录（相对项目根或绝对路径），默认 <process>/{CHECK_DIR_NAME}")
    return parser.parse_args(argv)


def fail(message: str) -> int:
    print(f"错误：{message}")
    return 2


def warn(message: str) -> None:
    print(f"[提示] {message}")


# ---------------------------------------------------------------- 配置 / 清单


def materials_path(arg_materials: str | None) -> Path:
    if arg_materials:
        return Path(arg_materials).expanduser().resolve()
    cfg = load_lib_config()
    configured = cfg.get("required_materials_file")
    if isinstance(configured, str) and configured.strip():
        candidate = SKILL_DIR / configured.strip()
        if candidate.is_file():
            return candidate
    return SKILL_DIR / "规则" / MATERIALS_NAME


def _as_list(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [str(v).strip() for v in value if str(v).strip()]
    text = str(value).strip()
    return [text] if text else []


def normalize_item(raw: dict, extra: dict) -> dict:
    item = dict(extra)
    item.update(
        {
            "id": str(raw.get("id") or "").strip(),
            "name": str(raw.get("name") or "").strip(),
            "level": str(raw.get("level") or "").strip(),
            "chapters": [str(c).strip() for c in (raw.get("chapters") or []) if str(c).strip()],
            "why": str(raw.get("why") or "").strip(),
            "hints": [h.lower() for h in _as_list(raw.get("hints"))],
            "weak_hints": [h.lower() for h in _as_list(raw.get("weak_hints"))],
            "exclude_hints": [h.lower() for h in _as_list(raw.get("exclude_hints"))],
            "ext": [e.lower().lstrip(".") for e in _as_list(raw.get("ext"))],
            "source": str(raw.get("source") or "").strip(),
            # 落位子目录：补齐后放到投递区下的哪个文件夹（按提供方归组）。
            # 缺了会在导出时按材料名兜底推导。
            "delivery_dir": str(raw.get("delivery_dir") or "").strip(),
            "categories": _as_list(raw.get("match_category")),
        }
    )
    return item


def collect_items(spec: dict, direction: str) -> tuple[list[dict], list[str], list[str]]:
    """返回 (items, errors, notes)。"""
    errors: list[str] = []
    notes: list[str] = []
    common_raw = spec.get("common")
    if not isinstance(common_raw, list) or not common_raw:
        errors.append("必备材料清单里 common 为空或格式不对（应为列表）")
        return [], errors, notes

    extra_map = spec.get("direction_extra") or {}
    if not isinstance(extra_map, dict):
        errors.append("必备材料清单里 direction_extra 格式不对（应为映射）")
        extra_map = {}
    direction_raw = extra_map.get(direction) or []
    if direction not in extra_map:
        notes.append(f"方向「{direction}」未配置 direction_extra，本次只检查通用清单。")

    items: list[dict] = []
    seen: dict[str, int] = {}
    for raw, extra in [(r, {"scope": "通用"}) for r in common_raw] + [
        (r, {"scope": direction}) for r in direction_raw
    ]:
        if not isinstance(raw, dict):
            errors.append(f"清单条目不是映射：{raw!r}")
            continue
        item = normalize_item(raw, extra)
        if not item["id"]:
            errors.append(f"清单条目缺少 id：{item['name'] or raw}")
            continue
        if not item["name"]:
            errors.append(f"清单条目 {item['id']} 缺少 name")
        if item["level"] not in LEVELS:
            errors.append(f"清单条目 {item['id']} 的 level 必须是 {'/'.join(LEVELS)}，实际 {item['level']!r}")
        if item["id"] in seen:
            errors.append(f"清单编号重复：{item['id']}")
        seen[item["id"]] = seen.get(item["id"], 0) + 1
        unknown_chapters = [c for c in item["chapters"] if c not in CHAPTERS]
        if unknown_chapters:
            errors.append(f"清单条目 {item['id']} 的 chapters 含未知章节：{'、'.join(unknown_chapters)}")
        if item["level"] == "必需" and not (item["hints"] or item["categories"]):
            errors.append(f"必需条目 {item['id']} 既无 hints 也无 match_category，永远无法匹配到材料")
        items.append(item)
    return items, errors, notes


# ---------------------------------------------------------------- 文件扫描 / 分类清单


def scan_delivery(delivery: Path, ignore_globs: list[str]) -> list[dict]:
    files: list[dict] = []
    for dirpath, dirnames, filenames in os.walk(delivery):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        for name in sorted(filenames):
            if any(fnmatch.fnmatch(name, pattern) for pattern in ignore_globs):
                continue
            path = Path(dirpath) / name
            if not path.is_file():
                continue
            rel = rel_posix(path, delivery)
            files.append(
                {
                    "rel": rel,
                    "name": name,
                    "parent": Path(rel).parent.as_posix(),
                    "ext": path.suffix.lower().lstrip("."),
                    "size": path.stat().st_size,
                }
            )
    return files


def resolve_classification_path(project_dir: Path, dirs_cfg: dict, explicit: str = "") -> tuple[Path | None, list[str]]:
    """定位 01 建档与材料归集实际产出的分类清单；返回 (路径或 None, notes)。

    优先级：--classification 显式指定 > 现行 01 产物路径 > 规则里的 manifest_dir > 旧布局。
    显式指定但文件不存在时不静默回落——由调用方判定为输入错误，避免「一份输入两个结论」。
    """
    notes: list[str] = []
    if explicit:
        path = Path(explicit).expanduser()
        if not path.is_absolute():
            path = project_dir / path
        return path, notes
    candidates: list[Path] = [
        project_dir / "过程文件" / "01-建档与材料归集" / "分类清单" / "分类清单.json",
    ]
    configured = str(dirs_cfg.get("manifest_dir") or "").strip().strip("/\\")
    if configured:
        candidates.append(project_dir / Path(configured) / "分类清单.json")
    candidates.append(project_dir / "05_过程" / "分类清单" / "分类清单.json")  # 旧布局兜底
    for cand in candidates:
        if cand.is_file():
            return cand, notes
    return None, notes


def load_classification(project_dir: Path, dirs_cfg: dict, explicit: str = "") -> tuple[dict[str, dict], str, list[str]]:
    """读 01 产出的分类清单，返回 ({rel: file_entry}, 相对路径, notes)。读不到不是错误。"""
    notes: list[str] = []
    manifest_path, notes = resolve_classification_path(project_dir, dirs_cfg, explicit)
    if manifest_path is None:
        notes.append("未找到 01 建档产出的分类清单（过程文件/01-建档与材料归集/分类清单/分类清单.json），"
                     "本次只按文件名/路径匹配（不读正文）。")
        return {}, "", notes
    try:
        rel_manifest = manifest_path.resolve().relative_to(project_dir.resolve()).as_posix()
    except (OSError, ValueError):
        rel_manifest = manifest_path.as_posix()
    if not manifest_path.is_file():
        notes.append(f"分类清单不存在：{rel_manifest}，本次只按文件名/路径匹配（不读正文）。")
        return {}, "", notes
    try:
        payload = json.loads(read_text(manifest_path))
    except (OSError, json.JSONDecodeError) as exc:
        notes.append(f"分类清单读取失败（{type(exc).__name__}），本次只按文件名/路径匹配。")
        return {}, "", notes
    entries: dict[str, dict] = {}
    for entry in payload.get("files") or []:
        if not isinstance(entry, dict):
            continue
        rel = str(entry.get("rel") or "").strip()
        if not rel:
            continue
        entries[rel] = entry
    if not entries:
        notes.append(f"分类清单里没有文件记录：{rel_manifest}，本次只按文件名/路径匹配。")
    recorded_dir = str(payload.get("project_dir") or "").strip()
    # 用户目录模式下分类清单把 project_dir 记成 "."（避免产物带盘符），不是可比对的绝对路径，
    # 直接拿它和项目目录比会永远「不一致」→ 误报提示。只在记录了绝对路径时才做比对。
    if recorded_dir and recorded_dir != ".":
        recorded_path = Path(recorded_dir)
        if not recorded_path.is_absolute():
            recorded_path = (project_dir / recorded_path).resolve()
        try:
            same = os.path.normcase(os.path.realpath(str(recorded_path))) == \
                os.path.normcase(os.path.realpath(str(project_dir)))
        except OSError:
            same = False
        if not same:
            notes.append("分类清单记录的 project_dir 与当前项目目录不一致（可能旧布局），已按文件相对路径复用类别判定。")
    return entries, rel_manifest, notes


# ---------------------------------------------------------------- 匹配


def match_item(item: dict, files: list[dict]) -> dict:
    hits: list[dict] = []
    for f in files:
        if item["ext"] and f["ext"] not in item["ext"]:
            continue
        name = f["name"].lower()
        parent = f["parent"].lower()
        if any(h and h in name for h in item["exclude_hints"]):
            continue
        strong_hits = [h for h in item["hints"] if h and h in name]
        path_hits = [h for h in item["hints"] if h and h in parent]
        weak_hits = [h for h in item["weak_hints"] if h and (h in name or h in parent)]

        category_hit = False
        category_label = ""
        if item["categories"]:
            cat_id = str(f.get("category_id") or "")
            cat_label = str(f.get("category_label") or "")
            if cat_id in item["categories"] or cat_label in item["categories"]:
                category_hit = True
                category_label = cat_label or cat_id

        score = (
            WEIGHT_FILENAME * len(strong_hits)
            + WEIGHT_PATH * len(path_hits)
            + WEIGHT_CATEGORY * (1 if category_hit else 0)
            + WEIGHT_WEAK * len(weak_hits)
        )
        if score <= 0:
            continue
        evidence: list[str] = []
        if strong_hits:
            evidence.append("文件名命中「" + "、".join(strong_hits) + "」")
        if path_hits:
            evidence.append("路径命中「" + "、".join(path_hits) + "」")
        if category_hit:
            evidence.append(f"分类清单判定为「{category_label}」")
        if weak_hits:
            evidence.append("弱线索「" + "、".join(weak_hits) + "」")
        hits.append(
            {
                "file": f["rel"],
                "score": score,
                "evidence": evidence,
                "confirmed": bool(strong_hits or path_hits or category_hit) and score >= STRONG_MIN_SCORE,
                "category": category_label,
            }
        )

    hits.sort(key=lambda h: (-h["score"], h["file"]))
    confirmed = [h for h in hits if h["confirmed"]]
    if confirmed:
        status = STATUS_OK
    elif hits:
        status = STATUS_SUSPECT
    else:
        status = STATUS_MISSING

    result = dict(item)
    result["status"] = status
    result["hits"] = hits[:MAX_EVIDENCE_FILES]
    result["hit_count"] = len(hits)
    result["confirmed_hits"] = [h for h in hits if h["confirmed"]]
    return result


# ---------------------------------------------------------------- 汇总


def rate_text(ok: int, total: int) -> str:
    if total <= 0:
        return "—"
    return f"{ok / total * 100:.1f}%"


def summarize(items: list[dict]) -> dict:
    by_level: dict[str, dict] = {}
    for level in LEVELS:
        subset = [i for i in items if i["level"] == level]
        ok = sum(1 for i in subset if i["status"] == STATUS_OK)
        suspect = sum(1 for i in subset if i["status"] == STATUS_SUSPECT)
        missing = sum(1 for i in subset if i["status"] == STATUS_MISSING)
        by_level[level] = {
            "total": len(subset),
            "ok": ok,
            "suspect": suspect,
            "missing": missing,
            "rate": round(ok / len(subset) * 100, 1) if subset else 0.0,
        }
    required = by_level["必需"]
    summary = {
        "items_total": len(items),
        "by_level": by_level,
        "required_total": required["total"],
        "required_ok": required["ok"],
        "required_suspect": required["suspect"],
        "required_missing": required["missing"],
        "required_rate": required["rate"],
        "suggested_total": by_level["建议"]["total"],
        "suggested_ok": by_level["建议"]["ok"],
        "suggested_missing": by_level["建议"]["missing"],
        "optional_total": by_level["可选"]["total"],
        "optional_ok": by_level["可选"]["ok"],
        "missing_required_count": len(
            [i for i in items if i["level"] == "必需" and i["status"] == STATUS_MISSING]
        ),
        "suspect_required_count": len(
            [i for i in items if i["level"] == "必需" and i["status"] == STATUS_SUSPECT]
        ),
    }
    summary["exit_code"] = 1 if (required["missing"] or required["suspect"]) else 0
    return summary


def chapter_ref(item: dict) -> str:
    parts: list[str] = []
    for name in item["chapters"]:
        idx = CHAPTERS.index(name)
        parts.append(f"第{CN_NUM[idx]}章 {name}")
    return "、".join(parts)


def clip(text: str, limit: int = 42) -> str:
    return text if len(text) <= limit else text[: limit - 1] + "…"


def esc(text: str) -> str:
    return str(text).replace("|", "\\|").replace("\n", " ").replace("\r", " ")


# ---------------------------------------------------------------- 输出：md


def build_payload(
    project: str,
    direction: str,
    delivery_label: str,
    delivery: Path,
    files: list[dict],
    items: list[dict],
    summary: dict,
    meta: dict,
) -> dict:
    def item_brief(item: dict) -> dict:
        return {
            "id": item["id"],
            "name": item["name"],
            "level": item["level"],
            "scope": item["scope"],
            "chapters": item["chapters"],
            "why": item["why"],
            "source": item["source"],
            "delivery_dir": item.get("delivery_dir", ""),
            "status": item["status"],
            "hits": item["hits"],
            "confirmed_hits": item["confirmed_hits"],
            "suggestion": "" if item["status"] == STATUS_OK else "补齐后重跑本命令",
        }

    return {
        "schema": 1,
        "generated_at": now_str(),
        "readonly": True,
        "project": project,
        "project_name": meta.get("project_name") or project,
        "direction": direction,
        "delivery": {
            "label": delivery_label,
            "abs": meta.get("delivery_display") or str(delivery),
            "file_count": len(files),
            "total_size": sum(f["size"] for f in files),
            "total_size_h": format_size(sum(f["size"] for f in files)),
        },
        "materials_file": meta["materials_file"],
        "classification_manifest": meta["manifest_rel"] or None,
        "summary": summary,
        "items": [item_brief(i) for i in items],
        "missing_required": [item_brief(i) for i in items if i["level"] == "必需" and i["status"] == STATUS_MISSING],
        "suspect_required": [item_brief(i) for i in items if i["level"] == "必需" and i["status"] == STATUS_SUSPECT],
        "missing_suggested": [item_brief(i) for i in items if i["level"] == "建议" and i["status"] != STATUS_OK],
        "missing_optional": [item_brief(i) for i in items if i["level"] == "可选" and i["status"] != STATUS_OK],
        "files": [{k: f[k] for k in ("rel", "ext", "size")} for f in files],
        "notes": meta["notes"],
    }


def build_gap_json(payload: dict) -> dict:
    """按 cardo/material-gap/v1 生成 材料齐套表.json（供 update_index.py --gap-file 消费）。"""
    summary = payload["summary"]
    by_level = summary["by_level"]
    required = by_level["必需"]
    suggested = by_level["建议"]
    optional = by_level["可选"]

    def json_status(status: str) -> str:
        return "待确认" if status == STATUS_SUSPECT else status

    items: list[dict] = []
    for item in payload["items"]:
        if item["status"] == STATUS_OK:
            matched = [h["file"] for h in item["confirmed_hits"]] or [h["file"] for h in item["hits"]]
        elif item["status"] == STATUS_SUSPECT:
            matched = [h["file"] for h in item["hits"]]
        else:
            matched = []
        items.append(
            {
                "id": item["id"],
                "name": item["name"],
                "level": item["level"],
                "chapters": item["chapters"],
                "status": json_status(item["status"]),
                "matched": matched,
                "source": item["source"],
                # 落位目录：补齐后放到投递区下的哪个子目录（按提供方归组）
                "delivery_dir": item.get("delivery_dir", ""),
                "why": item["why"],
            }
        )

    return {
        "schema": "cardo/material-gap/v1",
        "project": payload["project"],
        "project_name": payload.get("project_name") or payload["project"],
        "direction": payload["direction"],
        "delivery": payload["delivery"]["label"],
        "generated_at": payload["generated_at"],
        "readonly": True,
        "materials_file": payload["materials_file"],
        "summary": {
            "required_total": required["total"],
            "required_hit": required["ok"],
            "required_missing": required["total"] - required["ok"],
            "suggested_total": suggested["total"],
            "suggested_hit": suggested["ok"],
            "rate": round(required["ok"] / required["total"], 2) if required["total"] else 0.0,
            "required_suspect": required["suspect"],
            "suggested_missing": suggested["total"] - suggested["ok"],
            "optional_total": optional["total"],
            "optional_hit": optional["ok"],
            "items_total": summary["items_total"],
            "files_scanned": payload["delivery"]["file_count"],
            # 落盘口径与控制台一致：缺必需（含待确认）→ 1（提醒，不是错误）
            "exit_code": summary["exit_code"],
        },
        "items": items,
        "notes": payload.get("notes") or [],
    }


def write_table_md(path: Path, payload: dict) -> None:
    summary = payload["summary"]
    items = payload["items"]
    # 抬头用项目全称；目录名单独列出，便于追溯来源目录
    display = payload.get("project_name") or payload["project"]
    lines: list[str] = []
    lines.append(f"# 材料齐套表 · {display}")
    lines.append("")
    lines.append("> 由 `cardo-project-material/scripts/check_materials.py` 只读检查生成，可重复执行覆盖。")
    lines.append("> 本表不改动投递区任何文件；「已齐」只表示有这份材料，版本与时效需另行确认。")
    lines.append("")
    lines.append("## 概况")
    lines.append("")
    lines.append("| 项 | 值 |")
    lines.append("|---|---|")
    lines.append(f"| 生成时间 | {payload['generated_at']} |")
    lines.append(f"| 项目 | {esc(display)} |")
    if display != payload["project"]:
        lines.append(f"| 工作区目录 | {esc(payload['project'])} |")
    lines.append(f"| 方向 | {esc(payload['direction'])} |")
    lines.append(f"| 投递区 | `{esc(payload['delivery']['label'])}`（{payload['delivery']['file_count']} 个文件，"
                 f"{payload['delivery']['total_size_h']}） |")
    lines.append(f"| 材料清单 | {esc(payload['materials_file'])} |")
    lines.append(f"| 分类清单 | {esc(payload['classification_manifest'] or '（未使用）')} |")
    lines.append("")
    lines.append("## 齐套率")
    lines.append("")
    lines.append("| 等级 | 已齐 | 待确认-疑似 | 缺失 | 合计 | 齐套率（已齐/合计） |")
    lines.append("|---|---|---|---|---|---|")
    for level in LEVELS:
        row = summary["by_level"][level]
        if row["total"] == 0:
            continue
        lines.append(
            f"| {level} | {row['ok']} | {row['suspect']} | {row['missing']} | {row['total']} | "
            f"{rate_text(row['ok'], row['total'])} |"
        )
    lines.append("")
    lines.append(
        f"**必需材料：已齐 {summary['required_ok']}/{summary['required_total']}"
        f"（{rate_text(summary['required_ok'], summary['required_total'])}），"
        f"待确认 {summary['required_suspect']}，缺失 {summary['required_missing']}。**"
    )
    lines.append("")
    lines.append("## 明细")
    lines.append("")
    lines.append("| 编号 | 材料名 | 等级 | 影响章节 | 状态 | 依据 | 建议来源 |")
    lines.append("|---|---|---|---|---|---|---|")
    for item in items:
        if item["status"] == STATUS_OK:
            detail = "命中：" + "；".join(h["file"] for h in item["confirmed_hits"]) or "命中"
            if len(item["confirmed_hits"]) > 1:
                detail += f"（共 {len(item['confirmed_hits'])} 份，任一份算齐）"
        elif item["status"] == STATUS_SUSPECT:
            detail = "疑似：" + "；".join(f"{h['file']}（{'、'.join(h['evidence'])}）" for h in item["hits"])
        else:
            detail = f"未找到；缺失影响：{item['why']}"
        chapters = "、".join(item["chapters"])
        lines.append(
            f"| {item['id']} | {esc(item['name'])} | {item['level']} | {esc(chapters)} | {item['status']} | "
            f"{esc(detail)} | {esc(item['source'])} |"
        )
    lines.append("")
    lines.append("## 缺失清单（必需）")
    lines.append("")
    if payload["missing_required"]:
        lines.append("| 编号 | 材料名 | 影响章节 | 为什么需要 | 向谁要 |")
        lines.append("|---|---|---|---|---|")
        for item in payload["missing_required"]:
            lines.append(
                f"| {item['id']} | {esc(item['name'])} | {esc('、'.join(item['chapters']))} | "
                f"{esc(item['why'])} | {esc(item['source'])} |"
            )
    else:
        lines.append("无。必需材料已齐套。")
    lines.append("")
    lines.append("## 待确认清单（疑似）")
    lines.append("")
    if payload["suspect_required"]:
        lines.append("| 编号 | 材料名 | 疑似文件 | 判定依据 | 向谁确认 |")
        lines.append("|---|---|---|---|---|")
        for item in payload["suspect_required"]:
            files = "；".join(h["file"] for h in item["hits"])
            evidence = "；".join("；".join(h["evidence"]) for h in item["hits"])
            lines.append(
                f"| {item['id']} | {esc(item['name'])} | {esc(files)} | {esc(evidence)} | {esc(item['source'])} |"
            )
    else:
        lines.append("无。")
    lines.append("")
    lines.append("## 判定口径与说明")
    lines.append("")
    lines.append("- 只读：扫描投递区，不移动/复制/改名/删除任何原始文件，也不往投递区写东西。")
    lines.append("- 匹配：文件名/父目录命中 hints，或已有分类清单命中 match_category 即为强证据；"
                 "weak_hints 只算疑似。扩展名 ext 为硬门槛。")
    lines.append("- 已齐 ≠ 质量合格：齐套率只回答「有没有」，不回答「新不新、全不全、能不能直接用」，"
                 "版本与时效需另行确认（属闸门③）。")
    lines.append("- 待确认-疑似按未齐处理（退出码 1），因为「像但不是」不能直接当依据用。")
    if payload["notes"]:
        lines.append("")
        lines.append("### 过程提示")
        for note in payload["notes"]:
            lines.append(f"- {note}")
    lines.append("")
    write_text(path, "\n".join(lines))


# ---------------------------------------------------------------- 输出：控制台


def build_reminder(payload: dict) -> str:
    summary = payload["summary"]
    display = payload.get("project_name") or payload["project"]
    lines: list[str] = []
    lines.append(f"【材料补齐提醒】{display}")
    lines.append(
        f"方向：{payload['direction']}；投递区已收到 {payload['delivery']['file_count']} 个文件。"
    )
    lines.append("")
    lines.append(
        f"必需材料：已齐 {summary['required_ok']}/{summary['required_total']}"
        f"（{rate_text(summary['required_ok'], summary['required_total'])}），"
        f"待确认 {summary['required_suspect']}，缺失 {summary['required_missing']}"
    )
    suggested = summary["by_level"]["建议"]
    if suggested["total"]:
        lines.append(
            f"建议材料：已齐 {suggested['ok']}/{suggested['total']}，缺失 {suggested['missing']}"
        )
    lines.append("")

    def one_line(idx: int, item: dict) -> str:
        return (
            f"{idx}. {item['name']} —— {item['why']}；影响：{chapter_ref(item)}；"
            f"向「{item['source']}」要"
        )

    if payload["missing_required"]:
        lines.append("【必需·缺失】缺这些，对应章节会没有第一手依据，请优先补：")
        for idx, item in enumerate(payload["missing_required"], start=1):
            lines.append(one_line(idx, item))
        lines.append("")
    else:
        lines.append("【必需·缺失】无，必需材料已齐。")
        lines.append("")

    if payload["suspect_required"]:
        lines.append("【必需·待确认】有相近文件，但不能确认就是这一项，请确认是否可用：")
        for idx, item in enumerate(payload["suspect_required"], start=1):
            hit = item["hits"][0] if item["hits"] else None
            if hit:
                lines.append(
                    f"{idx}. {item['name']}：疑似「{clip(hit['file'])}」（{'、'.join(hit['evidence'])}）"
                    f"→ 影响：{chapter_ref(item)}；请向「{item['source']}」确认"
                )
            else:
                lines.append(f"{idx}. {item['name']}→ 影响：{chapter_ref(item)}；请向「{item['source']}」确认")
        lines.append("")

    if payload["missing_suggested"]:
        lines.append("【建议·缺失】有则更扎实，没有可先出稿：")
        for idx, item in enumerate(payload["missing_suggested"], start=1):
            lines.append(
                f"{idx}. {item['name']} —— {item['why']}；影响：{chapter_ref(item)}；"
                f"向「{item['source']}」要"
            )
        lines.append("")

    lines.append(
        "以上只列缺口，已齐材料不重复。补齐后把文件放回投递区、重跑本命令即可；"
        "原件不用改名、不用移动。"
    )
    return "\n".join(lines)


def print_text(payload: dict, outputs: dict) -> None:
    summary = payload["summary"]
    print("=== 齐套检查结果（只读） ===")
    print(f"项目：{payload['project']}")
    print(f"方向：{payload['direction']}")
    print(
        f"投递区：{payload['delivery']['label']}（{payload['delivery']['file_count']} 个文件，"
        f"{payload['delivery']['total_size_h']}）"
    )
    print(
        f"必需：已齐 {summary['required_ok']}/{summary['required_total']}"
        f"（{rate_text(summary['required_ok'], summary['required_total'])}）；"
        f"待确认 {summary['required_suspect']}；缺失 {summary['required_missing']}"
    )
    print(
        f"建议：已齐 {summary['by_level']['建议']['ok']}/{summary['by_level']['建议']['total']}；"
        f"缺失 {summary['by_level']['建议']['missing']}"
    )
    print("")
    print("──────── 以下为「材料补齐提醒」，可直接发给用户 ────────")
    print(build_reminder(payload))
    print("──────── 提醒结束 ────────")
    print("")
    print(f"齐套表（内部留档）：{outputs['md']}")
    print(f"机器可读（供缺项映射）：{outputs['json']}")
    if summary["exit_code"] == 0:
        print("结论：必需材料已齐，可以进入拆解。")
    else:
        print("结论：必需材料未齐（退出码 1，属提醒不是错误）。")


def print_quiet(payload: dict, outputs: dict) -> None:
    gap = build_gap_json(payload)
    brief = {
        "schema": gap["schema"],
        "project": gap["project"],
        "direction": gap["direction"],
        "delivery": gap["delivery"],
        "summary": gap["summary"],
        "missing_required_ids": [i["id"] for i in payload["missing_required"]],
        "suspect_required_ids": [i["id"] for i in payload["suspect_required"]],
        "exit_code": payload["summary"]["exit_code"],
        "outputs": outputs,
    }
    print(json.dumps(brief, ensure_ascii=False, indent=2))


# ---------------------------------------------------------------- 主流程


def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    args = parse_args(argv)
    # 用户目录模式：显式 --project-dir，或两者都不给（当前目录像项目就用它，否则 exit 2）
    user_mode = bool(str(args.project_dir or "").strip()) or not (args.root or args.project)

    # 1. 项目目录：--project-dir（用户目录）优先；--root/--project 走校准/回归布局
    if user_mode:
        try:
            project_dir = resolve_user_project_dir(args.project_dir)
        except UserProjectError as exc:
            return fail(str(exc))
        project = project_dir.name
    else:
        try:
            root = resolve_root(args.root)
        except SystemExit as exc:
            print(str(exc))
            return 2
        if not root.is_dir():
            return fail(f"项目库根目录不存在：{root}")

        try:
            project = resolve_project_name(args.project)
        except SystemExit as exc:
            print(str(exc))
            return 2

        try:
            project_dir = resolve_project_dir(root, project, args.direction)
        except SystemExit as exc:
            print(str(exc))
            return 2
        if project_dir is None:
            if args.direction:
                available = "、".join(direction_names())
                return fail(
                    f"项目「{project}」在方向「{args.direction}」下不存在"
                    f"（应位于 <项目库根>\\{args.direction}\\{project}）。可选方向：{available}。"
                )
            hits = [f"{d}/{project}" for d in direction_names() if (root / d / project).is_dir()]
            hint = f"；同名目录出现在：{'、'.join(hits)}" if hits else ""
            return fail(f"找不到项目目录：{project}（已按方向层与扁平布局查找）{hint}")

    # 2. 方向（可选）：命令行 > .cardo/project.json（只读）> --root 布局的方向目录 > 默认方向
    if user_mode:
        raw_arg = str(args.direction or "").strip()
        raw_card = card_direction(project_dir)
        if raw_arg:
            canonical = canon_direction(raw_arg)
            if canonical is None:
                return fail(f"未知方向 {raw_arg!r}；可选：{'、'.join(direction_names())}")
            direction = canonical
        else:
            canonical = canon_direction(raw_card) if raw_card else None
            if canonical:
                direction = canonical
            else:
                direction = default_direction()
                hint = "（.cardo/project.json 的 direction 不可识别）" if raw_card else ""
                warn(f"未判定方向{hint}，按默认方向「{direction}」检查（方向属闸门 G2，应由用户确认）。")
    else:
        if args.direction:
            direction = args.direction.strip()
        else:
            parent = project_dir.parent.name
            if parent in direction_names():
                direction = parent
            else:
                direction = default_direction()
                warn(f"项目不在方向层目录下，按默认方向「{direction}」检查。")

    # 2b. 项目全称：--project-name 优先（总控传工作区参数里的真实项目名），
    #     缺省用目录名。仅影响报告抬头，不参与齐套判定。
    project_name = str(args.project_name or "").strip() or project
    if project_name != project:
        print(f"提示：项目全称采用「{project_name}」（工作区参数），目录名为「{project}」。")

    # 3. 规则与材料清单
    rules_path = RULES_FILE
    if not rules_path.is_file():
        return fail(f"分类规则文件不存在：{rules_path}")
    rules = load_yaml_file(rules_path)
    dirs_cfg = rules.get("dirs") or {}
    ignore_globs = [str(p) for p in (rules.get("ignore") or [])]
    candidates = [str(c) for c in (rules.get("delivery_candidates") or []) if str(c).strip()]
    if not candidates:
        return fail("分类规则文件里 delivery_candidates 为空，无法定位投递区。")

    spec_path = materials_path(args.materials)
    if not spec_path.is_file():
        return fail(f"必备材料清单不存在：{spec_path}")
    try:
        spec = load_yaml_file(spec_path)
    except SystemExit as exc:
        print(str(exc))
        return 2
    items_raw, errors, notes = collect_items(spec, direction)
    if errors:
        print("错误：必备材料清单内容有误：")
        for err in errors:
            print(f"  - {err}")
        return 2
    if not items_raw:
        return fail(f"必备材料清单没有任何条目：{spec_path}")

    # 4. 投递区（索引记录优先，回落候选名；找不到/多个候选 → 输入错误，不猜）
    delivery_path, delivery_label, gate_msgs, delivery_note = resolve_delivery(project_dir, None, candidates)
    if gate_msgs:
        print("错误：无法确定投递区（不做猜测、不新建目录）：")
        for msg in gate_msgs:
            print(f"  - {msg}")
        return 2
    assert delivery_path is not None
    if delivery_note:
        notes.append(delivery_note)

    # 5. 扫描 + 分类清单 + 匹配
    files = scan_delivery(delivery_path, ignore_globs)
    if args.classification:
        cls_path, cls_notes = resolve_classification_path(project_dir, dirs_cfg, args.classification)
        notes.extend(cls_notes)
        if cls_path is None or not cls_path.is_file():
            if files:
                return fail(f"--classification 指定的分类清单不存在：{args.classification}"
                            "（01 建档应已产出；不静默回落到别的清单，避免同一份输入出两个结论）")
            # 投递区为空时 01 不会产出分类清单，这不是错误：按空清单继续，全部判缺
            notes.append("投递区为空，01 未产出分类清单；本次按无材料处理。")
    manifest, manifest_rel, manifest_notes = load_classification(project_dir, dirs_cfg, args.classification)
    notes.extend(manifest_notes)
    if manifest:
        matched_rels = {f["rel"] for f in files}
        missing_in_delivery = [rel for rel in manifest if rel not in matched_rels]
        if missing_in_delivery:
            notes.append(
                f"分类清单里有 {len(missing_in_delivery)} 条记录在投递区已找不到"
                "（文件可能被移动/改名），本次以投递区实际文件为准。"
            )
    for f in files:
        entry = manifest.get(f["rel"]) or {}
        f["category_id"] = str(entry.get("category_id") or "")
        f["category_label"] = str(entry.get("category_label") or "")

    items = [match_item(item, files) for item in items_raw]
    summary = summarize(items)

    # 6. 输出（唯一写入：<项目>\\<process>\\齐套检查\\）
    out_dir = Path(args.out).expanduser() if args.out else Path(dirs_cfg.get("process") or "05_过程") / CHECK_DIR_NAME
    if not out_dir.is_absolute():
        out_dir = project_dir / out_dir
    try:
        if out_dir.resolve().is_relative_to(delivery_path.resolve()):
            return fail(f"输出目录落在投递区内，拒绝写入（投递区只读）：{out_dir}")
    except AttributeError:  # pragma: no cover - 旧 Python 无 is_relative_to
        pass

    meta = {
        "materials_file": str(spec_path) if not user_mode else portable_display(spec_path),
        "manifest_rel": manifest_rel,
        "notes": notes,
        "project_name": project_name,
        "delivery_display": str(delivery_path) if not user_mode else rel_posix(delivery_path, project_dir),
    }
    payload = build_payload(
        project, direction, delivery_label, delivery_path, files, items, summary, meta
    )
    gap_json = build_gap_json(payload)
    md_path = out_dir / TABLE_MD
    json_path = out_dir / TABLE_JSON
    try:
        write_table_md(md_path, payload)
        write_text(json_path, json.dumps(gap_json, ensure_ascii=False, indent=2) + "\n")
    except OSError as exc:
        print(f"错误：无法写入齐套表（可能是只读目录、无权限或文件被占用）：{disp_path(out_dir)}")
        print(f"      原因：{type(exc).__name__}: {exc}")
        print("      检查未完成；请换可写目录（--out）或调整权限后重跑。")
        return 3  # 写盘失败是执行错误，不是「缺材料」
    outputs = {"md": str(md_path), "json": str(json_path)}
    if user_mode:
        outputs = {"md": disp_path(md_path), "json": disp_path(json_path)}

    # 7. 控制台
    if args.quiet:
        print_quiet(payload, outputs)
    elif args.format == "json":
        print(json.dumps(gap_json, ensure_ascii=False, indent=2))
    else:
        print_text(payload, outputs)
    return summary["exit_code"]


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception as exc:  # 崩溃与「缺料」必须分开出口，不能把程序错误翻译成业务结论
        import traceback
        traceback.print_exc()
        print(f"[齐套检查][内部错误] {type(exc).__name__}: {exc}", file=sys.stderr)
        print("  这是脚本执行错误（环境/代码问题），不是缺料判定；请按上面的堆栈定位。", file=sys.stderr)
        sys.exit(3)