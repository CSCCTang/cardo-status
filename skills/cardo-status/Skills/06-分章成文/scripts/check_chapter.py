#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""06 分章成文统一自检引擎（cardo/chapter-check/v1）。

背景：A–F 六章原先各持一份 check_chapter.py，实为两套互不一致的实现——
A/B 一套（23 条检查，输出 id/level/message 的 CHK-* 结构并写 --out 文件），
C 是 A 系 + 规则副本在位检查（44 条），D/E/F 另一套（id/status/detail，--out 被当目录、
只有 --json 才落盘）。总控 step06 按后者调用时读不到 JSON，D/E/F 自检形同虚设。

本引擎把两套判据合并成一份，按 `--unit A|B|C|D|E|F` 选本章 CONFIG：
  * 检查项取并集，并按 CONFIG 门控（例如 D/E/F 没有小节字数阈值，就不发 CHK-LEN-1）；
  * 同一语义只查一次（来源行/证据密度/字数/禁句统一走 A 系口径，D/E 的 S/P 编号、
    映射率、必答要点等专属项另发新 id）；
  * A/B/C 系 CHK-* id 原样保留——总控 规则/门禁.json 的 blockingChecks 依赖它；
    D/E/F 专属检查（CHK-ITEM-1/2、CHK-TBL-2、CHK-MAP-1、CHK-KW-1、CHK-GAP-4/5、
    CHK-SCOPE-1）只作 FAIL/WARN 展示，不参与阻断判定；
  * 输出统一为 {"schema","chapter","draft","direction","checks":[{"id","level","message"}],
    "summary","exit_code"}；--out 一律按“文件路径”直写该文件（不再当目录）。

用法：
    python <技能目录>/Skills/06-分章成文/scripts/check_chapter.py --unit C --project-dir . \
        --draft 过程文件/06-分章成文/分章初稿/C_现状详细分析.md \
        --direction 老旧街区 --out 过程文件/06-分章成文/自检/C.json
    python <技能目录>/Skills/06-分章成文/scripts/check_chapter.py --unit D --draft D_核心问题汇总.md
    python <技能目录>/Skills/06-分章成文/scripts/check_chapter.py --unit A --root <项目库根> --project <项目名>

退出码：0 无 FAIL；1 有 FAIL；2 参数/输入错误。只读草稿，只写 --out；仅标准库。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

__version__ = "1.0.0"
SCHEMA = "cardo/chapter-check/v1"

# ── 六章 CONFIG（数值逐字取自原 06-A…06-F/scripts/check_chapter.py，只做字段归一）──────
# 说明：原 A/B/C 用 chapter_min_chars/chapter_max_chars/section_*，原 D/E/F 用 min_chars/max_chars；
# 这里统一成 chapter_min_chars/chapter_max_chars，数值不变。section_* 只有 A/B/C 有，故 D/E/F 置 None
# （它们原本只查章总字数，不上小节阈值）。
# rules_skill 只有 C 需要（原 C 脚本的 CHK-RULE-1 查本章 skill 的 规则/ 副本是否齐全）。
# 政策/文件引用的识别特征：带文号、书名号文件名、成文时间或政策发布机关。
# 带这些的句子属可核验依据，其中的空洞措辞按原文保留，不判禁句。
_CITE_RE = re.compile(
    r"[〔\[]\s*\d{4}\s*\]\s*\d+\s*号"      # 文号：〔2024〕17号
    r"|《[^》]{2,}》"                            # 文件名（"《通知》"也算）
    r"|\d{4}\s*年\d{1,2}\s*月"                 # 成文时间
    r"|国办发|川办发|发改|住建"
    r"|《?[^》\s]{0,6}?(通知|办法|意见|方案|导则|标准|规定|批复|纪要)》?"
)


def _is_citable(text):
    """判断一句话是否为可核验的政策/文件引用。"""
    return bool(_CITE_RE.search(text or ""))


CONFIGS = {
    "A": {
        "unit": "A", "skill": "cardo-status-chapter-overview",
        "chapter_no": 1, "chapter_id": "ch01", "key": "overview",
        "h1": "项目概述", "draft_name": "A_项目概述.md",
        "sections": ["项目概况", "项目编制依据", "项目调研范围与期限"],
        "optional_sections": [], "allow_merged_sections": False, "lenient_directions": (),
        "chapter_min_chars": 1200, "chapter_max_chars": 9000,
        "section_min_chars": 300, "section_soft_chars": 900, "section_hard_chars": 1800,
        "item_prefix": "", "item_label": "", "item_elements": [], "item_element_policy": "warn",
        "tables": [], "mapping": None, "mapping_min": 0.8,
        "section_keywords": {}, "scope_patterns": [],
        "gap_numbering": False, "gap_closure": False, "rules_skill": "",
    },
    "B": {
        "unit": "B", "skill": "cardo-status-chapter-area",
        "chapter_no": 2, "chapter_id": "ch02", "key": "area",
        "h1": "片区基础概况分析", "draft_name": "B_片区基础概况分析.md",
        "sections": ["区位条件", "规划要求", "历史沿革分析"],
        "optional_sections": ["现状指标体系与测算结果"], "allow_merged_sections": False, "lenient_directions": (),
        "chapter_min_chars": 1200, "chapter_max_chars": 9000,
        "section_min_chars": 300, "section_soft_chars": 900, "section_hard_chars": 1800,
        "item_prefix": "", "item_label": "", "item_elements": [], "item_element_policy": "warn",
        "tables": [], "mapping": None, "mapping_min": 0.8,
        "section_keywords": {}, "scope_patterns": [],
        "gap_numbering": False, "gap_closure": False, "rules_skill": "",
    },
    "C": {
        "unit": "C", "skill": "cardo-status-chapter-current",
        "chapter_no": 3, "chapter_id": "ch03", "key": "current",
        "h1": "现状详细分析", "draft_name": "C_现状详细分析.md",
        "sections": [
            "土地与空间利用现状", "建筑现状与风貌保护现状", "人口与社会人居现状", "街区业态与经营现状",
            "交通与通行现状", "市政基础设施现状", "公共服务配套现状", "生态与人居环境现状",
        ],
        "optional_sections": [], "allow_merged_sections": False, "lenient_directions": (),
        "chapter_min_chars": 4000, "chapter_max_chars": 9000,
        "section_min_chars": 300, "section_soft_chars": 900, "section_hard_chars": 1800,
        "item_prefix": "", "item_label": "", "item_elements": [], "item_element_policy": "warn",
        "tables": [], "mapping": None, "mapping_min": 0.8,
        "section_keywords": {}, "scope_patterns": [],
        "gap_numbering": False, "gap_closure": False, "rules_skill": "06-C-现状详细分析",
        # 原 C 脚本 CHK-RULE-1 查的五份规范副本
        "rule_files": ["章节契约.yaml", "表格契约.yaml", "写作规范.yaml",
                       "chapter_skills.yaml", "质量评分表.yaml"],
    },
    "D": {
        "unit": "D", "skill": "cardo-status-chapter-problems",
        "chapter_no": 4, "chapter_id": "ch04", "key": "problems",
        "h1": "核心问题汇总", "draft_name": "D_核心问题汇总.md",
        "sections": ["街巷空间与肌理问题", "建筑安全与风貌破损问题", "文脉保护与活化不足问题",
                     "业态老旧与活力不足问题", "交通通行与消防隐患突出问题", "基础设施老旧滞后问题",
                     "民生配套与人居品质偏低", "街区治理与运营短板"],
        "optional_sections": ["成因分析", "核心问题清单"], "allow_merged_sections": True,
        "lenient_directions": ("老旧小区", "老旧厂区", "其他"),
        "chapter_min_chars": 3000, "chapter_max_chars": 12000,
        "section_min_chars": None, "section_soft_chars": None, "section_hard_chars": None,
        "item_prefix": "P", "item_label": "问题", "item_element_policy": "fail",
        "check_lead_source": True,  # 原 D 脚本把「本章导语」也纳入来源行检查
        "item_elements": [("表现", ("表现", "表征")), ("证据", ("证据",)), ("成因", ("成因",)),
                          ("影响", ("影响",)), ("对应更新方向", ("对应更新方向", "更新方向"))],
        "tables": [{"id": "TBL-02", "name": "核心问题清单", "required": True,
                    "columns": ["序号", "问题类别", "问题表征", "关键事实与指标"]}],
        "mapping": None, "mapping_min": 0.8,
        "section_keywords": {}, "scope_patterns": [],
        "gap_numbering": False, "gap_closure": False, "rules_skill": "",
    },
    "E": {
        "unit": "E", "skill": "cardo-status-chapter-strategy",
        "chapter_no": 5, "chapter_id": "ch05", "key": "strategy",
        "h1": "现状综合研判与更新思路", "draft_name": "E_现状综合研判与更新思路.md",
        "sections": ["综合现状研判", "总体更新思路", "核心更新改造方向"],
        "optional_sections": ["更新潜力与问题识别"], "allow_merged_sections": False, "lenient_directions": (),
        "chapter_min_chars": 1500, "chapter_max_chars": 14000,
        "section_min_chars": None, "section_soft_chars": None, "section_hard_chars": None,
        "item_prefix": "S", "item_label": "对策", "item_element_policy": "fail",
        "check_lead_source": True,  # 原 E 脚本把「本章导语」也纳入来源行检查
        "item_elements": [("对应问题", ("对应问题", "对应 P", "回应问题")),
                          ("更新动作", ("更新动作", "改造动作", "更新内容", "更新措施", "更新方向")),
                          ("预期成效", ("预期成效", "预期效果", "目标效果", "成效"))],
        "tables": [], "mapping": {"ref_prefix": "P"}, "mapping_min": 0.8,
        "section_keywords": {
            "综合现状研判": [("评价", "综合判断"), ("潜力",), ("问题总结", "问题归纳", "问题归并"), ("差距",)],
            "总体更新思路": [("原则",), ("保护",), ("微改", "微改造", "小尺度"), ("大拆大建",)],
        },
        "scope_patterns": ["工程量清单", "施工图", "管径\\s*\\d", "路面结构", "断面尺寸",
                           "投资估算", "投资额"],
        "gap_numbering": False, "gap_closure": False, "rules_skill": "",
    },
    "F": {
        "unit": "F", "skill": "cardo-status-chapter-attachments",
        "chapter_no": 6, "chapter_id": "ch06", "key": "attachments",
        "h1": "附件", "draft_name": "F_附件.md",
        "sections": ["附件一 现有基础资料清单", "附件二 图件与影像资料",
                     "附件三 专项台账",
                     "附件四 数据缺口清单"],
        # 附件四 由 08 assemble_chapters 从各章「本节缺项表」汇总生成，分章稿可不手写
        "optional_sections": ["附件四 数据缺口清单"],
        "allow_merged_sections": False, "lenient_directions": (),
        "chapter_min_chars": 600, "chapter_max_chars": 4000,
        "section_min_chars": None, "section_soft_chars": None, "section_hard_chars": None,
        "item_prefix": "", "item_label": "缺项", "item_elements": [], "item_element_policy": "fail",
        "tables": [{"id": "TBL-03", "name": "附件一 现有基础资料清单", "required": True,
                    "columns": ["序号", "文件名称", "类型", "主要用途", "本次使用情况"]},
                   {"id": "TBL-04", "name": "附件二 图件与影像资料", "required": True,
                    "columns": ["序号", "图件/影像", "应表达的内容", "状态"]},
                   {"id": "TBL-05", "name": "附件三 专项台账", "required": True,
                    "columns": ["序号", "台账名称", "应包含字段", "状态"]},
                   {"id": "TBL-06", "name": "附件四 数据缺口清单", "required": False,
                    "columns": ["编号", "缺项内容", "建议来源", "现有替代性判断"]}],
        "mapping": None, "mapping_min": 0.8,
        "section_keywords": {
            "附件一 现有基础资料清单": [("使用", "登记")],
            # v2 干净版：附件二/三只列本次实际收到的材料，不再要求「待补」字样
            "附件二 图件与影像资料": [("已提供", "已引用", "已收到", "以文字描述替代")],
            "附件三 专项台账": [("已收到", "已提供", "引用", "覆盖范围")],
        },
        "scope_patterns": [],
        "gap_numbering": True, "gap_closure": True, "rules_skill": "",
        "check_lead_source": True,  # 原 F 脚本把「本章导语」也纳入来源行检查
    },
}

# ── 内置判据（禁句来自 规则/禁句表.yaml 单一来源；规范是唯一真相）──────────────
# 运行时是 python313._pth 模式，脚本目录不进 sys.path：兄弟模块必须显式加进来
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from _banned import TODO_PHRASES as BANNED_PHRASES  # noqa: E402
except Exception as exc:  # 导入失败＝执行错误，不能退化成「本章不合格」的业务判定
    import traceback
    traceback.print_exc()
    print(f"[单章自检][内部错误] 无法导入同目录模块 _banned：{type(exc).__name__}: {exc}", file=sys.stderr)
    print("  这是脚本执行错误（打包缺失/环境问题），不是本章自检结论。", file=sys.stderr)
    raise SystemExit(3) from None

PROCESS_SECTIONS = ("本节缺项表", "来源映射")
DIRECTION_MARKERS = ["方向二", "方向三"]
STYLE_WARN = ["我们认为", "我们可以看到", "笔者", "显而易见"]
UNITS = ("万平方米|万平方公里|平方公里|平方米|万元|亿元|万人|公里|公顷|"
         "米|亩|吨|个|户|人|处|家|栋|座|条|项|间|层|套|盏|辆|%")
NUM_RE = re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)\s*(" + UNITS + ")")
DOCNO_RE = re.compile(r"〔\d{4}〕\s*\d+\s*号|GB(?:/T)?\s*\d{4,5}(?:[—\-]\d{4})?|第\s*\d+\s*号")
GAP_RE = re.compile(r"缺项-([A-Z])(\d{2})")
GAP_ANY_RE = re.compile(r"缺项-([A-F])(\d{1,3})")
HEAD_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*$")
TABLE_ROW_RE = re.compile(r"^\s*\|")
PROJECT_MARKERS = ("00_项目索引.md", ".cardo")
DRAFT_DIR = Path("03_拆解成果") / "分章初稿"

SKILLS_DIR = Path(__file__).resolve().parents[2]  # Skills/


class BadInput(Exception):
    """参数/输入错误 → exit 2。"""


# ── 通用工具 ─────────────────────────────────────────────────────────────
def read_text(path) -> str:
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        return f.read()


def rel_display(path, base=None) -> str:
    """对外只写相对路径（不出现盘符）：优先相对项目目录，其次相对当前目录，最后只留文件名。"""
    p = Path(path)
    for b in (base, Path.cwd()):
        if b is None:
            continue
        try:
            return p.resolve().relative_to(Path(b).resolve()).as_posix()
        except (ValueError, OSError):
            continue
    try:
        return Path(os.path.relpath(p, Path.cwd())).as_posix()
    except ValueError:
        return p.name


def norm_cell(text) -> str:
    return re.sub(r"\s+", "", text)


def char_count(text) -> int:
    return len(re.sub(r"\s", "", text))


def resolve_project_dir(args) -> Path:
    """项目定位（与总控 --project-dir / --root+--project 口径一致）。"""
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
        direction = str(args.direction or "").strip()
        candidates = ([root / direction / project] if direction else
                      sorted(d / project for d in root.iterdir() if d.is_dir() and (d / project).is_dir()))
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


def find_draft(project: Path, explicit: str, cfg) -> Path:
    """按 H1 标题定位本章草稿：显式路径 > <unit>_<H1>.md > 文件名含本章 H1 > 文件内有一级标题。"""
    if str(explicit or "").strip():
        path = Path(str(explicit).strip()).expanduser()
        path = path if path.is_absolute() else project / path
        if not path.is_file():
            raise BadInput(f"--draft 不存在：{path.name}")
        return path.resolve()
    base = project / DRAFT_DIR
    if not base.is_dir():
        raise BadInput(f"分章初稿目录不存在：{DRAFT_DIR.as_posix()}")
    named = base / cfg["draft_name"]
    if named.is_file():
        return named.resolve()
    same_h1 = [p for p in sorted(base.glob("*.md")) if p.stem.split("_", 1)[-1] == cfg["h1"]]
    if len(same_h1) == 1:
        return same_h1[0].resolve()
    if len(same_h1) > 1:
        raise BadInput(f"多份草稿命中「{cfg['h1']}」，请用 --draft 指定")
    fallback = []
    for p in sorted(base.glob("*.md")):
        try:
            head = read_text(p).splitlines()
        except OSError:
            continue
        if any(ln.strip() == f"# {cfg['h1']}" for ln in head):
            fallback.append(p)
    if len(fallback) == 1:
        return fallback[0].resolve()
    if len(fallback) > 1:
        raise BadInput(f"多份文件含一级标题「{cfg['h1']}」，请用 --draft 指定")
    raise BadInput(f"未找到本章草稿（期望 {cfg['draft_name']}，或文件名形如 <unit>_{cfg['h1']}.md）")


def load_titles(args, cfg):
    """--titles 覆盖节标题：逗号分隔，或 @文件（每行一个）。"""
    raw = str(args.titles or "").strip()
    if not raw:
        return list(cfg["sections"])
    if raw.startswith("@"):
        path = raw[1:]
        if not os.path.isfile(path):
            raise BadInput(f"--titles 文件不存在：{Path(path).name}")
        titles = [t.strip().lstrip("-* ").strip() for t in read_text(path).splitlines()]
        titles = [t for t in titles if t and not t.startswith("#")]
    else:
        titles = [t.strip() for t in re.split(r"[,，\n]", raw) if t.strip()]
    if not titles:
        raise BadInput("--titles 解析为空")
    return titles


# ── 文本解析 ─────────────────────────────────────────────────────────────
def parse_doc(text: str):
    lines = text.splitlines()
    heads = []
    for i, line in enumerate(lines):
        m = HEAD_RE.match(line)
        if m:
            heads.append((i, len(m.group(1)), m.group(2).strip()))
    return lines, heads


def section_bodies(lines, heads, start: int, end: int):
    """返回 [(title, [行...])]：H2 切节，H3+ 并入所属节，H1 截断。"""
    out, cur_title, cur_lines = [], None, []
    h2_at = {i: t for (i, lv, t) in heads if lv == 2 and start <= i < end}
    h1_at = {i for (i, lv, _t) in heads if lv == 1 and start <= i < end}
    for i in range(start, end):
        if i in h2_at:
            if cur_title is not None:
                out.append((cur_title, cur_lines))
            cur_title, cur_lines = h2_at[i], []
        elif i in h1_at:
            if cur_title is not None:
                out.append((cur_title, cur_lines))
            cur_title, cur_lines = None, []
        elif cur_title is not None:
            cur_lines.append(lines[i])
    if cur_title is not None:
        out.append((cur_title, cur_lines))
    return out


def prose_lines(body_lines):
    """正文行：去掉表格行、来源行、〔锚点〕依据块、空行（表格与依据不占篇幅）。"""
    return [ln for ln in body_lines
            if ln.strip() and not TABLE_ROW_RE.match(ln)
            and not ln.strip().startswith("〔锚点")
            and "（数据来源" not in ln and "(数据来源" not in ln]


def source_lines(body_lines):
    return [ln.strip() for ln in body_lines
            if ln.strip().startswith("（数据来源") or ln.strip().startswith("(数据来源")]


def count_evidence(text: str) -> dict:
    nums, docs, gaps = NUM_RE.findall(text), DOCNO_RE.findall(text), GAP_RE.findall(text)
    return {"unit_numbers": len(nums), "doc_numbers": len(docs),
            "gap_refs": len(gaps), "total": len(nums) + len(docs) + len(gaps)}


def is_source_format_ok(line: str) -> bool:
    """规范格式：（数据来源：<类型>，来源：<源>，第 N 页/段 N）／网络资料带网址与获取日期。"""
    return "数据来源" in line and bool(re.search(r"[，,]\s*来源\s*[：:]", line))


def parse_tables(body: str):
    """返回 [(header_cells, data_rows)]；跳过表头下的分隔行。"""
    tables, cur = [], None
    for line in str(body or "").splitlines():
        s = line.strip()
        if not (s.startswith("|") and s.endswith("|") and s.count("|") >= 3):
            cur = None
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if cur and all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
            continue
        if cur is None:
            cur = [cells, []]
            tables.append(cur)
        else:
            cur[1].append(cells)
    return tables


def item_blocks(text: str, prefix: str):
    """按行首 **P-NN / **S-NN 切分块，返回 [(id, block_text)]。"""
    if not prefix:
        return []
    marks = list(re.finditer(r"(?m)^\s*(?:[-*+]\s*)?\*\*(" + prefix + r"-\d{2})\b", text))
    return [(m.group(1), text[m.end():marks[i + 1].start() if i + 1 < len(marks) else len(text)])
            for i, m in enumerate(marks)]


# ── 各检查项（统一 add(id, level, message)）─────────────────────────────────
def checklist(draft: Path, text: str, args, cfg) -> list:
    checks = []

    def add(cid, level, msg):
        checks.append({"id": cid, "level": level, "message": msg})

    lines, heads = parse_doc(text)
    h1 = cfg["h1"]
    h1s = [(i, t) for (i, lv, t) in heads if lv == 1]
    if not h1s:
        add("CHK-HEAD-1", "FAIL", f"缺一级标题「{h1}」（模板标题逐字，层级别改成 ##）")
        return checks
    matched = [i for (i, t) in h1s if t == h1]
    if not matched:
        add("CHK-HEAD-1", "FAIL",
            f"一级标题应为「{h1}」，实际「{'、'.join(t for _i, t in h1s)}」（标题唯一真相=模板）")
        return checks
    if len(h1s) == 1:
        add("CHK-HEAD-1", "PASS", f"一级标题「{h1}」逐字一致")
    elif getattr(args, "allow_merged", False):
        add("CHK-HEAD-1", "WARN",
            f"一个文件里有 {len(h1s)} 个一级标题（旧写作单元合并稿）；本章按「{h1}」定位，合成前应拆成一章一文件")
    else:
        add("CHK-HEAD-1", "FAIL",
            f"一个文件里出现 {len(h1s)} 个一级标题；本章应只写「{h1}」（合成按章取文件）")

    start = matched[0]
    later = [i for (i, _t) in h1s if i > start]
    end = later[0] if later else len(lines)
    deep = [(i, lv, t) for (i, lv, t) in heads if start <= i < end and lv >= 3]
    add("CHK-HEAD-3", "FAIL" if deep else "PASS",
        ("出现三级及以下标题（模板最多二级）：" + "、".join(t for _i, _lv, t in deep[:3]))
        if deep else "无三级及以下标题（模板 max_level=2）")

    titles = load_titles(args, cfg)
    bodies = section_bodies(lines, heads, start, end)
    body_map = {t: b for t, b in bodies}
    optional = list(cfg["optional_sections"])
    required = [t for t in titles if t not in optional]
    found = [t for t, _b in bodies if t not in PROCESS_SECTIONS]
    lenient = (str(args.direction or "") and cfg["allow_merged_sections"]
               and any(d in str(args.direction) for d in cfg["lenient_directions"]))
    if lenient:
        add("CHK-HEAD-2", "INFO" if not found else "WARN",
            "非街区方向且无分类小节：节标题不作逐字校验（归并口径人工确认）" if not found else
            f"非街区方向：{len(found)} 个分类小节未逐字校验（须能归并到八类：{required[0] if required else '—'}…）")
    else:
        missing = [t for t in required if t not in body_map]
        extra = [t for t in found if t not in required and t not in optional]
        order_bad = ([found.index(t) for t in required if t in found] !=
                     sorted(found.index(t) for t in required if t in found))
        if missing:
            add("CHK-HEAD-2", "FAIL", "缺节标题：" + "、".join(missing))
        elif extra:
            add("CHK-HEAD-2", "FAIL", "出现模板外二级节：" + "、".join(extra))
        elif order_bad:
            add("CHK-HEAD-2", "FAIL", "节顺序与模板不一致：期望 " + " / ".join(required))
        else:
            add("CHK-HEAD-2", "PASS", f"{len(required)} 节标题逐字、顺序一致")

    h2_positions = [i for (i, lv, _t) in heads if lv == 2 and start <= i < end]
    head_text = "\n".join(lines[start + 1:(h2_positions[0] if h2_positions else end)])
    # 导语来源行：原 D/E/F 会把「本章导语」也纳入来源行检查，A/B/C 不查；按 CONFIG 门控。
    # 成稿口径下导语可用〔锚点〕依据块替代来源行（两者都算出处已交代）。
    if cfg.get("check_lead_source") and head_text.strip():
        head_lines = [ln.strip() for ln in head_text.strip().splitlines() if ln.strip()]
        lead_has_anchor = any(ln.startswith("〔锚点") for ln in head_lines)
        if head_lines and ("|" in head_text or len(head_text) >= 40) \
                and not head_lines[-1].startswith("（数据来源：") and not lead_has_anchor:
            add("CHK-SRC-1", "FAIL", "「本章导语」缺来源行（节末须独立一行「（数据来源：…）」或〔锚点〕依据块）")
        else:
            add("CHK-SRC-1", "PASS", "「本章导语」出处已交代（来源行或依据块）")

    content_titles = [t for t in titles if t in body_map and t not in PROCESS_SECTIONS]
    total_prose, inline_gap_ids = 0, []
    for title in content_titles:
        body = body_map[title]
        raw, ptext = "\n".join(body), "\n".join(prose_lines(body))
        nchars = char_count(ptext)
        total_prose += nchars
        # 「编制依据」是按文号逐条罗列的清单式小节，条目本身即证据，
        # 不适用 1800 字硬上限（实测 31 条依据 3400 余字，压缩只能删条目）。
        excused = "不涉及" in ptext[:120] or "编制依据" in title
        srcs = source_lines(body)
        if any(ln.strip().startswith("〔锚点") for ln in body):
            # 成稿口径：出处写进逐段依据块，由《依据对照表》承载，正文不背来源行
            n_ev = sum(1 for ln in body if ln.strip().startswith("〔锚点"))
            add("CHK-SRC-1", "PASS", f"「{title}」标注源成文（依据块 {n_ev} 条）")
        elif not srcs:
            add("CHK-SRC-1", "FAIL", f"「{title}」缺来源行（每节至少 1 条「（数据来源：…）」）")
        elif len(srcs) > 1:
            add("CHK-SRC-1", "WARN", f"「{title}」有 {len(srcs)} 条来源行（建议合并为 1 条）")
        else:
            add("CHK-SRC-1", "PASS", f"「{title}」来源行 1 条")
        if srcs and not all(is_source_format_ok(s) for s in srcs):
            add("CHK-SRC-2", "FAIL" if getattr(args, "strict_source_format", False) else "WARN",
                f"「{title}」来源行格式未按规范「（数据来源：<类型>，来源：<源>，第 N 页/段 N）」")
        ev = count_evidence(raw)
        if ev["total"] >= 2:
            add("CHK-NUM-1", "PASS",
                f"「{title}」证据 {ev['total']} 处（带单位数字 {ev['unit_numbers']}／"
                f"文号标准号 {ev['doc_numbers']}／缺项编号 {ev['gap_refs']}）")
        elif excused:
            add("CHK-NUM-1", "WARN", f"「{title}」声明「不涉及」，证据仅 {ev['total']} 处（按方向豁免，无需补数）")
        elif ev["gap_refs"]:
            # 本节已按规范登记缺项编号：缺什么、找谁补都写清楚了，
            # 证据不足是**材料缺口**而非写作缺陷，判 FAIL 会让缺口版永远跑不过。
            add("CHK-NUM-1", "WARN",
                f"「{title}」证据 {ev['total']} 处（含缺项编号 {ev['gap_refs']} 处），"
                f"已按缺项登记，无需补数")
        else:
            add("CHK-NUM-1", "FAIL", f"「{title}」带单位数字/可核验事实仅 {ev['total']} 处，要求 ≥2"
                                     "（缺数据写缺项编号，见「数据与证据要求」）")
        # 小节字数阈值只有 A/B/C 配了；D/E/F 只查章总字数，不发 CHK-LEN-1
        if cfg["section_min_chars"] is not None:
            if nchars < cfg["section_min_chars"] and not excused:
                add("CHK-LEN-1", "FAIL",
                    f"「{title}」正文 {nchars} 字 < 下限 {cfg['section_min_chars']} 字")
            elif nchars > (cfg["section_hard_chars"] or 10 ** 9) and not excused:
                add("CHK-LEN-1", "FAIL",
                    f"「{title}」正文 {nchars} 字 > 硬上限 {cfg['section_hard_chars']} 字（压缩空话，不删证据）")
            elif nchars > (cfg["section_soft_chars"] or 10 ** 9) and not excused:
                add("CHK-LEN-1", "WARN",
                    f"「{title}」正文 {nchars} 字 > 软目标 {cfg['section_soft_chars']} 字（超长请先压缩空话）")
            else:
                add("CHK-LEN-1", "PASS", f"「{title}」正文 {nchars} 字")
        inline_gap_ids += [f"{a}{b}" for a, b in GAP_RE.findall(raw)]

    low, high = cfg["chapter_min_chars"], cfg["chapter_max_chars"]
    if total_prose < low:
        add("CHK-LEN-2", "FAIL", f"本章正文 {total_prose} 字 < 章下限 {low} 字")
    elif total_prose > high:
        add("CHK-LEN-2", "FAIL", f"本章正文 {total_prose} 字 > 章上限 {high} 字")
    else:
        add("CHK-LEN-2", "PASS", f"本章正文 {total_prose} 字（{low}–{high}）")

    scan = text
    # 禁句按「句」判定，并豁免**可核验的政策/文件引用**：
    # 政策原文里的「大力推进」「进一步加强」是原文表述、可核验，不是空洞套话；
    # 没有出处的才是问题。口径与 06-分章成文/scripts/gen_chapters.py 的
    # is_todo_phrase / is_empty_phrase 保持一致（两端口径必须相同）。
    # 判定粒度是**段落**而非单句：政策引用往往整段都是原文
    # （如「《通知》部署，要有序推进……。大力推进有条件的楼栋加装电梯。」），
    # 按句切会把末句单独拎出来判成空洞套话。按段判定，段内有出处则整段豁免。
    hits = []
    for para in re.split(r"\n\s*\n", scan):
        if not para.strip():
            continue
        if _is_citable(para):
            continue
        for sent in re.split(r"(?<=[。；！？])", para):
            for bp in BANNED_PHRASES:
                if bp in sent:
                    hits.append(bp)
    # 同一禁句只记一次，避免同一句里多次出现刷屏
    seen_bp = set()
    uniq_hits = []
    for bp in hits:
        if bp not in seen_bp:
            seen_bp.add(bp)
            uniq_hits.append((bp, hits.count(bp)))
    add("CHK-BAN-1", "FAIL" if uniq_hits else "PASS",
        ("命中禁句：" + "、".join(f"{p}×{c}" for p, c in uniq_hits)) if uniq_hits
        else f"无禁句（清单 {len(BANNED_PHRASES)} 条；政策引用已豁免）")
    dm = [p for p in DIRECTION_MARKERS if p in scan]
    add("CHK-BAN-2", "FAIL" if dm else "PASS",
        ("出现模板方向标记：" + "、".join(dm) + "（按本项目方向直接写）") if dm else "无方向标记残留")
    st = [p for p in STYLE_WARN if p in scan]
    if st:
        add("CHK-STYLE-1", "WARN", "出现第一人称/口语表述：" + "、".join(st))

    # 缺项编号：前缀 / 重复 / 与「本节缺项表」闭环
    prefix_expect = cfg["unit"]
    m = re.match(r"^([A-F])(?:_|$)", draft.stem)
    if m:
        prefix_expect = m.group(1)
    letters = sorted({a for a, _b in GAP_RE.findall(text)})
    if not letters:
        add("CHK-GAP-1", "PASS", "无内联缺项编号（如本章确有缺项请补编号）")
    elif len(letters) > 1:
        # 按章编号（D/E/F…）是契约规定做法，而 08 装配会把六章合成一份文档，
        # 合成后必然同时出现多个前缀。原判据要求「全篇单一前缀」，
        # 会把规范写法判成 FAIL。改为：每个前缀各自成组、编号连续，
        # 且都必须是合法的章字母；前缀多于一个时不再判失败。
        #
        # **跨章引用例外**：D 章引用 C 章的缺项编号时，若该行显式说明
        # （「详见第三章缺项登记」），这个 C03 不算 C 组自己的编号，
        # 不参与 C 组的连续性检查——否则会逼作者删掉证据来凑连续。
        cross_ref_re = re.compile(r"见第\s*[一二三四五六七八九十\d]+\s*章|详见|参见|另见")
        own_pairs: dict[str, set[int]] = {}
        for ln in text.splitlines():
            if cross_ref_re.search(ln):
                continue
            for a, b in GAP_RE.findall(ln):
                own_pairs.setdefault(a, set()).add(int(b))
        bad = [a for a in letters if a not in "ABCDEF"]
        if bad:
            add("CHK-GAP-1", "FAIL", "缺项编号前缀非法：" + "、".join(bad))
        else:
            groups = {}
            for a, b in GAP_RE.findall(text):
                groups.setdefault(a, set()).add(int(b))
            broken = []
            for a, nums in sorted(groups.items()):
                ns = sorted(nums)
                own = sorted(own_pairs.get(a, set()))
                # 本章自己的编号要连续；只有跨章引用（无自有编号）的组不判
                if not own:
                    continue
                if ns != list(range(1, len(ns) + 1)):
                    broken.append(a)
            if broken:
                add("CHK-GAP-1", "FAIL",
                    "缺项编号不连续（应从 01 起）：" + "、".join(broken)
                    + "（确为跨章引用请写「详见第三章缺项登记」）")
            else:
                add("CHK-GAP-1", "PASS",
                    "缺项编号按章成组且连续：" + "、".join(f"缺项-{a}xx" for a in letters))
    elif letters[0] != prefix_expect:
        # 单前缀且与本章不符：可能是跨章引用。显式说明过的按 WARN 放行。
        if re.search(r"见第\s*[一二三四五六七八九十\d]+\s*章|详见|参见|另见", text):
            add("CHK-GAP-1", "WARN",
                f"缺项编号前缀「缺项-{letters[0]}xx」与本章 {prefix_expect} 不一致，"
                f"但已显式说明跨章引用，放行")
        else:
            add("CHK-GAP-1", "FAIL" if getattr(args, "strict_gap_prefix", False) else "WARN",
                f"缺项编号前缀「缺项-{letters[0]}xx」与本章写作单元 {prefix_expect}"
                f"（缺项-{prefix_expect}xx）不一致")
    else:
        add("CHK-GAP-1", "PASS", f"缺项编号前缀统一为「缺项-{prefix_expect}xx」")

    dup = sorted({f"{a}{b}" for a, b in GAP_RE.findall(text) if text.count(f"缺项-{a}{b}") > 2})
    if dup:
        add("CHK-GAP-2", "WARN", "同一缺项编号被引用超过 2 次：" + "、".join("缺项-" + d for d in dup))
    else:
        add("CHK-GAP-2", "PASS", "缺项编号无异常重复")

    table_ids, gap_sec = [], body_map.get("本节缺项表")
    if gap_sec is not None:
        for ln in gap_sec:
            if TABLE_ROW_RE.match(ln):
                table_ids += [f"{a}{b}" for a, b in GAP_RE.findall(ln)]
        dup_table = sorted({i for i in table_ids if table_ids.count(i) > 1})
        if dup_table:
            add("CHK-GAP-2", "FAIL", "本节缺项表编号重复：" + "、".join("缺项-" + d for d in dup_table))
        orphan = sorted(set(inline_gap_ids) - set(table_ids))
        if orphan:
            add("CHK-GAP-3", "WARN", "内联缺项未进本节缺项表：" + "、".join("缺项-" + d for d in orphan))
        elif inline_gap_ids:
            add("CHK-GAP-3", "PASS", "内联缺项编号均已在缺项表登记")
    elif inline_gap_ids:
        add("CHK-TBL-1", "FAIL", "正文有缺项编号但缺「## 本节缺项表」")
    else:
        add("CHK-TBL-1", "WARN", "未列「## 本节缺项表」（本章无缺项可接受，有缺项必须汇总）")

    for proc in PROCESS_SECTIONS:
        add("CHK-PROC-1", "PASS" if proc in body_map else "WARN",
            f"{'有' if proc in body_map else '缺'}过程段「{proc}」")

    # ── F 专属：缺项编号组内连续 / 附件四与正文闭环（新 id，只展示不阻断）──
    if cfg["gap_numbering"]:
        ids = sorted({(a, int(b)) for a, b in GAP_ANY_RE.findall(text)})
        if not ids:
            add("CHK-GAP-4", "WARN", "未发现缺项编号（材料齐备时可写“本项目无数据缺口”，否则须逐条编号）")
        else:
            groups = {}
            for letter, num in ids:
                groups.setdefault(letter, []).append(num)
            bad = [f"缺项-{l}：{'、'.join('%02d' % n for n in groups[l])}" for l in sorted(groups)
                   if groups[l] != list(range(1, len(groups[l]) + 1))]
            add("CHK-GAP-4", "FAIL" if bad else "PASS",
                ("缺项编号不连续/重号：" + "；".join(bad)) if bad else
                f"缺项编号 {len(ids)} 个，{len(groups)} 组字母各自 01…N 连续")
    if cfg["gap_closure"]:
        inline = {m2.group(0) for m2 in re.finditer(r"缺项-[A-F]\d{1,3}", text)}
        closure_ids = {r[0].strip() for s in bodies if s[0].startswith("附件四")
                       for hdr, rows in parse_tables("\n".join(s[1]))
                       if [norm_cell(c) for c in hdr][:2] == ["编号", "缺项内容"]
                       for r in rows if r and r[0].strip()}
        if not closure_ids:
            add("CHK-GAP-5", "INFO", "分章稿未手写附件四（由 assemble_chapters.py 汇总生成，合成后校验闭环）")
        else:
            gap_missing, orphan = sorted(closure_ids - inline), sorted(inline - closure_ids)
            add("CHK-GAP-5", "WARN" if (gap_missing or orphan) else "PASS",
                (f"附件四 与正文内联缺项不闭合：表内多 {'、'.join(gap_missing[:4]) or '无'}；"
                 f"正文多 {'、'.join(orphan[:4]) or '无'}（合成期以全篇为准）")
                if (gap_missing or orphan) else f"附件四 {len(closure_ids)} 条与正文内联缺项闭合")

    # ── D/E 专属：编号连续与要素齐全 ──
    prefix = cfg["item_prefix"]
    if prefix:
        ids = [i for t in content_titles
               for i, _b in item_blocks("\n".join(body_map[t]), prefix)]
        if not ids:
            add("CHK-ITEM-1", "FAIL",
                f"未发现 {prefix}-NN 编号（每条{cfg['item_label']}须以独立行 **{prefix}-NN 名称** 开头）")
        elif ids != [f"{prefix}-{k:02d}" for k in range(1, len(ids) + 1)]:
            add("CHK-ITEM-1", "FAIL", "编号不连续或乱序：实际 " + "、".join(ids))
        else:
            add("CHK-ITEM-1", "PASS", f"{prefix}-01…{prefix}-{len(ids):02d} 共 {len(ids)} 条连续")

        total, bad = 0, []
        for t in content_titles:
            for iid, block in item_blocks("\n".join(body_map[t]), prefix):
                total += 1
                lack = [n for n, alts in cfg["item_elements"] if not any(a in block for a in alts)]
                if lack:
                    bad.append(f"{iid} 缺 {'/'.join(lack)}")
        if not total:
            add("CHK-ITEM-2", "SKIP", f"无 {prefix} 编号块，跳过要素检查")
        elif bad:
            add("CHK-ITEM-2", "FAIL" if cfg["item_element_policy"] == "fail" else "WARN",
                f"{len(bad)}/{total} 条缺要素：" + "；".join(bad[:5]))
        else:
            add("CHK-ITEM-2", "PASS",
                f"{total} 条均含 " + "/".join(n for n, _a in cfg["item_elements"]))

    # ── E 专属：S→P 映射覆盖 ──
    if cfg["mapping"]:
        blocks = [b for t in content_titles
                  for b in item_blocks("\n".join(body_map[t]), prefix)]
        if not blocks:
            add("CHK-MAP-1", "FAIL", f"无 {prefix} 编号块，映射率无法计算")
        else:
            known = set()
            prob = next((p for p in sorted(draft.parent.glob("*.md"))
                         if p.is_file() and "核心问题汇总" in p.name), None)
            if prob is not None:
                try:
                    lines_p, _heads_p = parse_doc(read_text(prob))
                    known = {i for i, _b in item_blocks("\n".join(lines_p),
                                                       cfg["mapping"]["ref_prefix"])}
                except OSError:
                    known = set()
            refs = [re.findall(cfg["mapping"]["ref_prefix"] + r"-\d{2}", b) for _i, b in blocks]
            mapped = sum(1 for r in refs if r)
            rate = mapped / float(len(blocks))
            dangling = [f"{blocks[i][0]}→{r}" for i, rs in enumerate(refs)
                        for r in rs if known and r not in known]
            detail = (f"S→P 映射率 {mapped}/{len(blocks)}={rate * 100:.0f}%"
                      f"（阈值 {cfg['mapping_min'] * 100:.0f}%）")
            if dangling:
                add("CHK-MAP-1", "FAIL", detail + "；悬空引用：" + "、".join(dangling[:5]))
            elif rate + 1e-9 < cfg["mapping_min"]:
                add("CHK-MAP-1", "FAIL", detail + "；未映射：" +
                    "、".join([blocks[i][0] for i, r in enumerate(refs) if not r][:6]))
            else:
                add("CHK-MAP-1", "PASS", detail)

    # ── E/F 专属：核心节必答要点 ──
    if cfg["section_keywords"]:
        bad = []
        for t in content_titles:
            lack = [alts[0] for alts in cfg["section_keywords"].get(t, [])
                    if not any(a in "\n".join(body_map[t]) for a in alts)]
            if lack:
                bad.append(f"{t} 缺 {'/'.join(lack)}")
        add("CHK-KW-1", "FAIL" if bad else "PASS",
            (f"{len(bad)} 节未覆盖必答要点：" + "；".join(bad[:5])) if bad else
            f"核心节必答要点齐全（{len(cfg['section_keywords'])} 节受检）")

    # ── D/F 专属：表格契约 ──
    if cfg["tables"]:
        found_tables = []
        for title, body in [("本章导语", head_text)] + [(t, "\n".join(body_map[t])) for t in content_titles]:
            found_tables += [(title, hdr, rows) for hdr, rows in parse_tables(body)]
        fails, oks = [], []
        for spec in cfg["tables"]:
            hit = [f for f in found_tables
                   if [norm_cell(c) for c in f[1]] == [norm_cell(c) for c in spec["columns"]]]
            if not hit:
                if spec.get("required", True):
                    fails.append(f"{spec['id']} 未找到，表头须恰为「{'|'.join(spec['columns'])}」；实际："
                                 + ("；".join(f"{f[0]}({'|'.join(f[1])})" for f in found_tables[:3]) or "无表格"))
                continue
            empty = sum(1 for r in hit[0][2] if any(not c.strip() for c in r))
            if empty:
                fails.append(f"{spec['id']} 有 {empty} 行空单元格（契约 empty_cell_ratio=0）")
            else:
                oks.append(f"{spec['id']}@{hit[0][0]} {len(hit[0][2])} 行")
        if fails:
            add("CHK-TBL-2", "FAIL", "；".join(fails))
        else:
            add("CHK-TBL-2", "PASS", "；".join(oks) if oks else "表格齐全")

    # ── E 专属：越界候选（设计做法/工程量/投资额属方案阶段）──
    if cfg["scope_patterns"]:
        scope_scan = re.sub(r"《[^》]*》", "", text)
        scope = [f"越界「{m3.group(0)}」" for p in cfg["scope_patterns"]
                 for m3 in re.finditer(p, scope_scan)]
        add("CHK-SCOPE-1", "WARN" if scope else "PASS",
            ("疑似越界（设计做法/工程量/投资额属方案阶段）：" + "；".join(scope[:6])) if scope
            else "越界候选 0")

    # ── C 专属：本章规范副本在位（缺只 WARN，不改判据）──
    if cfg["rules_skill"]:
        base = SKILLS_DIR / cfg["rules_skill"] / "规则"
        miss = [n for n in cfg.get("rule_files", []) if not (base / n).is_file()]
        add("CHK-RULE-1", "WARN" if miss else "PASS",
            ("本地规范副本缺失：" + "、".join(miss) + "（请从 cardo-status-spec/规则 拷入 规则/）") if miss
            else f"本地规范副本齐全（规则/ 共 {len(cfg.get('rule_files', []))} 份）")

    return checks


def summarize(checks) -> dict:
    result = {"pass": 0, "warn": 0, "fail": 0, "info": 0, "skip": 0}
    for c in checks:
        key = str(c["level"]).lower()
        result[key] = result.get(key, 0) + 1
    result["total"] = len(checks)
    return result


def build_parser():
    ap = argparse.ArgumentParser(
        prog="check_chapter.py",
        description="06 分章成文统一自检：--unit A|B|C|D|E|F 选本章判据（只读草稿，只写 --out）")
    ap.add_argument("--unit", required=True, type=str.upper, choices=sorted(CONFIGS),
                    help="写作单元/章：A–F（大小写均可）")
    ap.add_argument("--project-dir", default="", help="项目目录（任意路径；最高优先级）")
    ap.add_argument("--root", default="", help="[校准/回归] 项目库根（与 --project 搭配）")
    ap.add_argument("--project", default="", help="[校准/回归] 项目名（与 --root 搭配）")
    ap.add_argument("--direction", default="", help="方向：老旧街区/老旧厂区/老旧小区")
    ap.add_argument("--draft", default="", help="分章稿路径（缺省按 H1 在 03_拆解成果/分章初稿 下定位）")
    ap.add_argument("--titles", default="", help="覆盖节标题：逗号分隔或 @文件")
    ap.add_argument("--min-chars", type=int, default=0, help="覆盖章总字数下限")
    ap.add_argument("--max-chars", type=int, default=0, help="覆盖章总字数上限")
    ap.add_argument("--out", default="", help="把检查结果写成 JSON（按文件路径直写该文件）")
    ap.add_argument("--json", action="store_true", help="把 JSON 结果打印到 stdout（默认打印人读清单）")
    ap.add_argument("--allow-merged", action="store_true", help="兼容旧写作单元合并稿（多个一级标题时只 WARN）")
    ap.add_argument("--strict-source-format", action="store_true", help="来源行格式不合规即 FAIL（默认 WARN）")
    ap.add_argument("--strict-gap-prefix", action="store_true", help="缺项前缀与本章写作单元不一致即 FAIL（默认 WARN）")
    ap.add_argument("--no-timestamp", action="store_true", help="JSON 不带生成时间（幂等）")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return ap


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    cfg = dict(CONFIGS[str(args.unit).upper()])  # 副本：--min-chars/--max-chars 只改本次覆盖值
    if args.min_chars:
        cfg["chapter_min_chars"] = args.min_chars
    if args.max_chars:
        cfg["chapter_max_chars"] = args.max_chars

    try:
        project = resolve_project_dir(args)
        draft = find_draft(project, args.draft, cfg)
        try:
            text = read_text(draft)
        except UnicodeDecodeError as exc:
            raise BadInput(f"草稿不是 UTF-8 文本：{exc}") from exc
        if not text.strip():
            raise BadInput("草稿是空文件")
    except (BadInput, OSError) as exc:
        print(f"[chapter-check] 输入错误：{exc}", file=sys.stderr)
        return 2

    checks = checklist(draft, text, args, cfg)
    stats = summarize(checks)
    exit_code = 1 if stats["fail"] else 0
    payload = {
        "schema": SCHEMA,
        "skill": cfg["skill"],
        "unit": cfg["unit"],
        "chapter": {"no": cfg["chapter_no"], "id": cfg["chapter_id"],
                    "key": cfg["key"], "h1": cfg["h1"]},
        "draft": rel_display(draft, project),
        "direction": str(args.direction or "").strip() or "(未指定)",
        "checks": checks,
        "summary": dict(stats, exit_code=exit_code),
        "exit_code": exit_code,
    }
    if not args.no_timestamp:
        payload["generated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")

    if args.out:
        out = Path(str(args.out).strip()).expanduser()
        out = out if out.is_absolute() else (project / out)
        try:
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            print(f"[chapter-check] JSON 写入失败：{exc}", file=sys.stderr)
            return 2

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return exit_code

    print(f"[chapter-check] {cfg['skill']} v{__version__}（unit {cfg['unit']}／第 {cfg['chapter_no']} 章）")
    print(f"[chapter-check] 项目：{project.name}　草稿：{payload['draft']}　方向：{payload['direction']}")
    for c in checks:
        print(f"  {c['level']:<4} {c['id']:<12} {c['message']}")
    tail = "PASS" if exit_code == 0 else "FAIL"
    print(f"[chapter-check] 结果：{tail}（失败 {stats['fail']}／警告 {stats['warn']}／通过 {stats['pass']}"
          f"／提示 {stats['info']}／跳过 {stats['skip']}）")
    if exit_code == 1:
        print("[chapter-check] 修 FAIL 项后重跑；只有 06 阻断集内的 CHK-* 才会卡总控，专属检查的 FAIL 属提示级")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
