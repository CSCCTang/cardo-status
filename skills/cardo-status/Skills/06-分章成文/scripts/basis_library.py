# -*- coding: utf-8 -*-
"""编制依据库读取与成文渲染。

依据库文件：包内 assets/依据库/依据库.json（可由 CARDO_BASIS_LIB 覆盖）。
运行时核验记录：过程文件/05-.../包内依据/依据核验记录.json（可由 CARDO_BASIS_VERIFY 覆盖）。

设计口径：
- 分类固定为 法律法规 / 政策文件 / 技术标准 / 编制办法 / 项目资料；
- 条目文字按公文层次序数排布（（一）类目 + 1. 2. 条目），不使用项目符号；
- 文号只取库内原值，库内没有就写“公开文本未载文号”，不臆造；
- 库内“待核实”条目（C02、C04）退出正式依据表，按核验结论单独说明；
- 本次运行未完成联网复核的条目统一加“†”，并在表后说明，不冒充已核验。
"""
from __future__ import annotations

import json
import os
from pathlib import Path

CATEGORIES = ["法律法规", "政策文件", "技术标准", "编制办法", "项目资料"]
LEVEL_ORDER = {"国家": 0, "某省": 1, "某市": 2, "项目": 3}
# 退出正式依据表、按“待核实”单独处理的条目
PENDING_IDS = {"C02", "C04"}
# 只作方向性泛引、不写文号的条目
DIRECTIONAL_ONLY = {"C04"}


def _package_root() -> Path:
    # Skills/06-分章成文/scripts/basis_library.py -> 包根
    return Path(__file__).resolve().parents[3]


def library_path() -> Path:
    p = os.environ.get("CARDO_BASIS_LIB")
    return Path(p) if p else _package_root() / "assets" / "依据库" / "依据库.json"


def verification_path() -> Path | None:
    p = os.environ.get("CARDO_BASIS_VERIFY")
    if p:
        return Path(p)
    return None


def load_library(path: Path | None = None) -> dict:
    path = Path(path) if path else library_path()
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def verify_ledger(path: Path | None = None) -> dict:
    path = Path(path) if path else verification_path()
    if not path or not Path(path).exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _level_rank(level: str) -> int:
    """层级排序权重。库内层级可能是具体地名（如“某省”“某市”），
    先按固定层级查，查不到再按行政区后缀兜底：省/自治区→1、市→2、其他→9。"""
    level = (level or "").strip()
    if level in LEVEL_ORDER:
        return LEVEL_ORDER[level]
    if level.endswith("省") or level.endswith("自治区"):
        return 1
    if level.endswith("市"):
        return 2
    return 9


def classify(entry: dict) -> str:
    typ = (entry.get("类型") or "").strip()
    name = (entry.get("名称") or "").strip()
    if typ == "编制办法":
        return "编制办法"
    if typ == "技术标准":
        return "技术标准"
    # 行政法规与地方性法规按“条例”归入法律法规
    if name.endswith("条例"):
        return "法律法规"
    return "政策文件"


def _fmt_date(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    parts = s.replace("/", "-").split("-")
    if len(parts) >= 3 and all(x.isdigit() for x in parts[:3]):
        return f"{int(parts[0])}年{int(parts[1])}月{int(parts[2])}日"
    return s


def _number_text(entry: dict) -> str:
    no = (entry.get("文号") or "").strip()
    if not no or no.startswith("未") or no.startswith("未见"):
        return ""
    return no


def _entry_line(entry: dict, idx: int, verified: bool, typ: str) -> str:
    name = (entry.get("名称") or "").strip()
    no = _number_text(entry)
    org = (entry.get("发布机关") or "").strip()
    pub = _fmt_date(entry.get("发布日期", ""))
    eff = _fmt_date(entry.get("施行日期", ""))
    bits = []
    if no:
        bits.append(no)
    elif typ == "技术标准" and not no:
        bits.append("标准号未载")
    if org:
        bits.append(org)
    if pub:
        bits.append(f"{pub}发布")
    if eff and eff != pub:
        # 库内施行日期本身可能已是“自印发之日起施行”，不再补“实施”二字
        bits.append(eff if eff.endswith(("施行", "实施", "执行")) else f"{eff}实施")
    head = f"{idx}. 《{name}》"
    body = "，".join([b for b in bits if b])
    if not no and head.startswith(f"{idx}. 《"):
        body = ("公开文本未载文号" if typ != "项目资料" else "项目资料，无文号") + (("，" + body) if body else "")
    line = head + (f"（{body}）" if body else "")
    if not verified:
        line += "†"
    return line


def _project_material_lines(items: list[dict], start_idx: int) -> list[str]:
    lines = []
    for i, it in enumerate(items, start=start_idx):
        name = it.get("name") or "项目资料"
        src = it.get("source") or ""
        note = it.get("note") or ""
        bits = ["项目资料，无文号"]
        if src:
            bits.append(f"来源：{src}")
        if note:
            bits.append(note)
        lines.append(f"{i}. 《{name}》（{'，'.join(bits)}）")
    return lines


def render_basis_section(
    *,
    project: str,
    ledger_name: str,
    verify: dict | None = None,
    project_materials: list[dict] | None = None,
    lib: dict | None = None,
) -> tuple[list[str], dict]:
    """返回 (markdown 行, 统计元数据)。库缺失时返回带“待补文号”的骨架，不编造条目。"""
    lib = lib if lib is not None else load_library()
    verify = verify if verify is not None else verify_ledger()
    entries = list((lib or {}).get("条目") or [])
    results = {x.get("id"): x for x in (verify.get("results") or [])}
    checked_at = verify.get("checked_at") or ""

    lines: list[str] = []
    meta = {"library": str(library_path()), "entries": 0, "categories": {}, "unverified": 0,
            "pending": sorted(PENDING_IDS), "checked_at": checked_at, "library_found": bool(entries)}

    if not entries:
        lines.append(
            "编制依据库未接入（assets/依据库/依据库.json 缺失）。本节按“待补文号”登记，"
            "凡未取得公开文号的文件一律不写文号、不引用具体条款，待依据库接入后统一替换。"
        )
        lines.append("")
        for cat in CATEGORIES[:-1]:
            lines.append(f"**（{_cn_num(CATEGORIES.index(cat) + 1)}）{cat}**")
            lines.append("")
            lines.append(f"1. 《{cat}依据待补》（文号待补，发布机关待补）")
            lines.append("")
        lines.append(f"**（{_cn_num(len(CATEGORIES))}）项目资料**")
        lines.append("")
        lines.extend(_project_material_lines(project_materials or [], 1))
        meta["unverified"] = 1
        return lines, meta

    grouped: dict[str, list[dict]] = {c: [] for c in CATEGORIES}
    for e in entries:
        if e.get("id") in PENDING_IDS:
            continue
        grouped[classify(e)].append(e)

    for c in grouped:
        grouped[c].sort(key=lambda e: (_level_rank(e.get("层级", "")), e.get("id", "")))

    counters = {c: len(grouped[c]) for c in CATEGORIES[:-1]}
    counters["项目资料"] = len(project_materials or [])
    total = sum(counters.values())
    meta["entries"] = total
    meta["categories"] = counters

    lines.append(
        f"编制依据按法律法规、政策文件、技术标准、编制办法、项目资料五类归集，共 {total} 项："
        f"法律法规 {counters['法律法规']} 项、政策文件 {counters['政策文件']} 项、"
        f"技术标准 {counters['技术标准']} 项、编制办法 {counters['编制办法']} 项、"
        f"项目资料 {counters['项目资料']} 项。各类之内按国家、某省、某市层级和文号顺序排布；"
        "文号一律照录依据库原值，库内未载文号的不补造。"
    )
    lines.append("")

    for ci, cat in enumerate(CATEGORIES, start=1):
        lines.append(f"**（{_cn_num(ci)}）{cat}**")
        lines.append("")
        if cat == "项目资料":
            lines.extend(_project_material_lines(project_materials or [], 1))
        else:
            n = 0
            for e in grouped[cat]:
                n += 1
                r = results.get(e.get("id"))
                verified = bool(r and r.get("status") == "一致")
                if not verified:
                    meta["unverified"] += 1
                lines.append(_entry_line(e, n, verified, cat))
        lines.append("")

    # 待核实条目的处理说明（C02 已失效、改引条例；C04 只作方向性泛引）
    lines.append(
        "另有两条库内条目不宜作为现行依据，处理如下，不计入上表：其一，"
        "《某市人民政府办公厅关于印发某市城市有机更新实施办法的通知》（某办发〔2020〕43号）"
        "经某市官方「失效或废止行政规范性文件」目录核实，已于 2025 年 5 月 27 日有效期届满失效，"
        "本报告不再引用，城市更新实施机制改引 2026 年施行的《某市城市更新条例》，"
        "涉及沿革时注明 43 号文已失效；其二，"
        "《某市城镇老旧院落（小区）改造技术导则（试行）》未核到印发文号，"
        "本报告仅作方向性泛引，不写文号、不引用其具体条款与指标。"
    )
    lines.append("")

    if checked_at:
        ok = sum(1 for x in results.values() if x.get("status") == "一致")
        bad = len(results) - ok
        lines.append(
            f"依据核验：本次生成于 {checked_at} 对可直连原文页的条目逐条复核，通过 {ok} 条、"
            f"未能联网 {bad} 条；带“†”的条目为本次未完成联网复核，按离线基线值引用，"
            "其库内登记核验状态与原文链接见内部件《依据核对表》。"
        )
    else:
        lines.append(
            "依据核验：本次生成未执行联网复核，上表全部条目按离线基线库引用（标“†”），"
            "库内登记核验状态与原文链接见内部件《依据核对表》。"
        )
    lines.append("")
    return lines, meta


def _cn_num(n: int) -> str:
    nums = "零一二三四五六七八九十"
    if 1 <= n <= 10:
        return nums[n]
    if n < 20:
        return "十" + nums[n - 10]
    return str(n)
