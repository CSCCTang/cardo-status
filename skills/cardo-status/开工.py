#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Cardo 现状分析 · 一条命令开工。

依次做五件事：
  1. 建工作区目录（输入文件 / 过程文件 / 成果文件 …）
  2. 释放标准模板（sha 锁定，被改动即停）
  3. 依赖自检；缺就在包根建 .venv，按 阿里云 → 清华 → USTC → 官方源 回退安装
  4. 校验输入与 工作区参数.json（缺料就停，不编造）
  5. 调 01→09 流水线，按退出码给下一步

退出码：0 全过；3 缺料/待人工确认；1 流水线硬失败；2 环境或结构问题。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import traceback

# 运行期不写 __pycache__：Skills/ 是只读交付物，写进 pyc 会污染交付库
# （实测交付库里混进过 19 个 pyc）。这里在导入任何本地模块之前就关掉。
sys.dont_write_bytecode = True

import zipfile
from datetime import datetime
from pathlib import Path

SKILL = Path(__file__).resolve().parent      # 技能包：只读，运行期不写入
RUNTIME_PY = SKILL / "runtime" / "python" / ("python.exe" if os.name == "nt" else "python")
ROOT = SKILL.parent
CACHE_BASE = Path(os.environ.get("CARDO_CACHE_HOME") or os.environ.get("LOCALAPPDATA") or str(Path.home()))
def venv_dir() -> Path:
    """派生仓库：虚拟环境放用户缓存，不写安装目录。"""
    return CACHE_BASE / "cardo" / "venv-status"
                          # 工作区：技能包所在的上一级目录
# 工作区是否已成功解析（resolve_workspace 会抛 SystemExit；未解析时不能拿默认 ROOT 去写看板状态）
WORKSPACE_RESOLVED = False

# 看板状态注入标记（开工.py 每次运行刷新这一段；看板读不到工作区状态）
STATUS_BEGIN = "<!-- CARDO-STATUS:BEGIN -->"
STATUS_END = "<!-- CARDO-STATUS:END -->"

TITLE = "Cardo 现状分析"
PY_MIN = (3, 9)

TEMPLATE_DST = "提资数据/模板/1 城市更新项目现状分析报告（模板）-补充数据来源（包含三类项目）.docx"
DIRS = ["输入文件", "过程文件", "成果文件", "记忆与规则", "提资数据/模板", "输入文件-模板"]
ASSETS = [
    {
        "src": "assets/templates/1 城市更新项目现状分析报告（模板）-补充数据来源（包含三类项目）.docx",
        "dst": TEMPLATE_DST,
        "sha256": "758289C670056EE75987DBC2880A05D7FFF9241A3B24742298F8BDAC61497102",
        "label": "现状分析报告模板",
    },
]
INPUT_TEMPLATES = []
# 运行期依赖。fitz（PyMuPDF）用于 PDF 材料**归类采样**（判文书/图纸）
# 与**有文本层 PDF 的逐页全文抽取**（只供事实引用，不进正文、不作章节结构来源）。
IMPORTS = ["docx", "openpyxl", "yaml", "fitz"]

# 技能包完整性自检：每步的**必需脚本**与**必需规则文件**。
# 为什么要自检：技能包被复制不完整时（实测有人只拷了 开工.py + assets + runtime，
# Skills/00~09 整目录缺失），流水线会一路跑到 09 才失败，报错还不指向根因。
# 故在开工第一步就查：缺哪个目录/脚本/规则，从哪补。
SKILL_STEPS = [
    ("00-Cardo现状分析总控", ["pipeline.py", "detect_project.py"], []),
    ("01-建档与材料归集", ["classify_materials.py", "confirm_and_merge.py"], ["分类规则.yaml"]),
    ("02-齐套检查与缺料拍板", ["check_materials.py", "export_gap_xlsx.py",
                              "make_material_request.py"], ["必备材料清单.yaml"]),
    ("03-模板与源报告转储", ["assemble_materials.py", "dump_docx_outline.py"],
     ["章节契约.yaml"]),
    ("04-章节映射", ["fill_mapping.py", "make_mapping_skeleton.py"], ["章节契约.yaml"]),
    ("05-事实抽取与事实库", ["extract_facts.py"], []),
    ("06-分章成文", ["build_write_tasks.py", "gen_chapters.py", "check_chapter.py"],
     ["写作规范.yaml", "禁句表.yaml", "章节契约.yaml", "表格契约.yaml"]),
    ("07-表3.1口径计算", ["calc_metrics.py", "build_table31.py"], ["计算口径.yaml"]),
    ("08-组装与模板渲染", ["assemble_chapters.py", "build_report_docx.py",
                           "export_evidence_pack.py", "pdf_render.py",
                           "gongwen_layout.py"], ["表格契约.yaml"]),
    ("09-三重自检与边界", ["check_contracts.py", "build_trace.py", "check_trace.py",
                            "boundary_check.py", "check_evidence.py"],
     ["章节契约.yaml", "表格契约.yaml"]),
]

PARAMS = "工作区参数.json"
SCHEMA = "cardo/workspace-params/v1"
DEFAULTS = {"template_file": TEMPLATE_DST, "delivery": "输入文件", "gap_mode": "auto"}
REQUIRED = ["project", "direction", "template_file", "source_report"]
DIRECTIONS = ["老旧街区", "老旧小区", "老旧厂区"]
# 「散材料自动汇编」约定标记（与 Skills/00-Cardo现状分析总控/scripts/pipeline.py 保持一致）：
# source_report 取此值表示「没有成篇现状稿，由 03 步按材料自动汇编成主线」。
ASSEMBLE_MARK = "__assemble__"

PIPELINE = "Skills/00-Cardo现状分析总控/scripts/pipeline.py"
DETECT = "Skills/00-Cardo现状分析总控/scripts/detect_project.py"
DELIVERABLES = [
    "成果文件/*现状分析报告*.docx",
    "成果文件/输入材料补齐确认单.docx",
    "成果文件/数据缺口清单.xlsx",
]
GUIDES = {
    "01": (
        "人工分类确认：读 过程文件/01-建档与材料归集/分类清单/分类清单.json，"
        "取 pending 的文件逐条问用户归到哪类，写 过程文件/01-建档与材料归集/分类确认.yaml"
        "（decisions[].file / category_id / note）。类别：feasibility 可研、planning 片区策划、"
        "checkup 体检、project_list 清单/投资表、approval 批复、mapping 测绘图纸、media 影像、"
        "minutes 纪要、output_report 既有成果、other 其他。"
    ),
    "02": "缺口拍板：问用户「先出稿（缺口版继续）/ 先补料（停下等补齐）」，选择后者时加 --material-choice 先补料。",
    "04": (
        "章节映射确认：读 过程文件/04-章节映射/映射统计.json（或 章节映射.yaml 中 status=medium 的节点），"
        "逐条问「保留该标题 / 判为缺项」，写 过程文件/04-章节映射/人工确认.yaml"
        "（nodes[].id + sources，判缺项写 sources: []）。"
    ),
}

# ---------------------------------------------------------------- 收料引导
INTAKE_ITEMS = [{'key': 'any', 'name': '【必需】任意一份项目材料', 'required': True, 'dst': '输入文件/（任意层级）', 'patterns': ['*'], 'how': 'docx / pdf / xlsx / dwg / rar 都算；文件夹与压缩包也行，放进来就会被 01 自动分类', 'why': '没有材料无法开工：流水线不从别处找材料，材料由你放到投递区'}, {'key': 'source', 'name': '成篇材料（现状底稿 / 旧报告，有就更好）', 'required': False, 'dst': '输入文件/（任意层级）', 'patterns': ['*现状*.docx', '*报告*.docx', '*说明*.docx', '*调研*.docx'], 'how': '有就放进 输入文件/，会被直接当主线；没有也行，散材料会按模板章节自动归位', 'why': '有成篇材料时主线更完整；没有则由 03 自动汇编'}, {'key': 'approval', 'name': '批复文件', 'required': False, 'dst': '输入文件/（任意层级）', 'patterns': ['*批复*.docx', '*批复*.pdf'], 'how': '立项/可研批复原件；有文本层的 PDF 会抽全文作事实证据（带页码，不进正文）', 'why': '项目性质、规模、投资的第一手依据'}, {'key': 'list', 'name': '项目清单 / 投资表', 'required': False, 'dst': '输入文件/（任意层级）', 'patterns': ['*清单*.xlsx', '*清单*.docx'], 'how': '改造项目清单、投资计划表', 'why': '工程量与投资口径的原始来源'}, {'key': 'feasibility', 'name': '可研 / 片区策划 / 体检报告', 'required': False, 'dst': '输入文件/（任意层级）', 'patterns': ['*可研*.docx', '*策划*.docx', '*体检*.docx'], 'how': '已有的前期成果；没有就登记缺口', 'why': '背景、必要性、片区定位的补充依据'}, {'key': 'mapping', 'name': '测绘图纸 / 影像', 'required': False, 'dst': '输入文件/（任意层级）', 'patterns': ['*.dwg', '*.jpg', '*.png', '*.pdf'], 'how': '图纸、航拍、现场照片', 'why': '现状描述与图件附件'}, {'key': 'minutes', 'name': '会议纪要 / 部门意见', 'required': False, 'dst': '输入文件/（任意层级）', 'patterns': ['*纪要*.docx', '*意见*.docx'], 'how': '座谈、评审、部门征求意见的书面记录', 'why': '问题与诉求的佐证'}, {'key': 'others', 'name': '其他材料（有什么放什么）', 'required': False, 'catch_all': True, 'dst': '输入文件/（任意层级）', 'patterns': ['*'], 'how': '统计年鉴、上级规划、新闻剪报、访谈录音转写、现场台账……凡是可能用得上的都能放，文件夹和 zip 也行', 'why': '01 会自动归类；归不了的会问你怎么归。这些材料会作为事实来源和附件参考'}]


def intake_scan():
    """逐项探测落位情况；返回 [(item, [命中文件])]。

    catch_all 项（「其他材料」）只统计没被任何具体项认领的文件，避免把已知材料重复计数。
    """
    import fnmatch

    base = ROOT / "输入文件"
    pool = [p for p in base.rglob("*")
            if p.is_file() and not p.name.startswith("~$") and "空白" not in p.name] if base.is_dir() else []
    claimed, rows = set(), []

    for item in INTAKE_ITEMS:
        if item.get("catch_all"):
            continue
        hits = [p for p in pool if any(fnmatch.fnmatch(p.name.lower(), pat.lower())
                                       for pat in item["patterns"])]
        if item.get("key") == "any":
            # 「任意一份材料」是汇总项：命中全部文件，但不独占——
            # 具体项（批复/清单/图纸…）仍各自认领自己的文件，否则它们会全部显示未到位。
            rows.append((item, sorted(set(pool))))
            continue
        claimed.update(hits)
        rows.append((item, sorted(set(hits))))

    for item in INTAKE_ITEMS:
        if item.get("catch_all"):
            rows.append((item, [p for p in pool if p not in claimed]))

    order = {id(item): index for index, item in enumerate(INTAKE_ITEMS)}
    return sorted(rows, key=lambda pair: order[id(pair[0])])


def write_input_hints(pdf_count: int, ratio: float, message: str,
                      text_layer: int = 0, scanned: int = 0) -> None:
    """把「输入提示」落成文件：pipeline 每次重写运行报告都会带上它，不依赖会话记忆。

    message 为空表示本次没有 PDF 口径问题 —— 仍然覆盖写，把上一次的提示清掉。
    text_layer / scanned 记录 PDF 文本层与扫描件的份数，便于事后核对。
    """
    out = ROOT / "过程文件/00-总控/输入提示.json"
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps({"schema": "cardo/input-hints/v1",
                                   "pdfCount": pdf_count,
                                   "pdfRatio": round(ratio, 4),
                                   "pdfTextLayer": text_layer,
                                   "pdfScanned": scanned,
                                   "messages": [message] if message else []},
                                  ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError as exc:
        say(f"  （输入提示写入失败，不影响流程：{exc}）")


def pdf_layer_stats(paths) -> tuple[int, int, int]:
    """统计 PDF 是否有文本层：返回 (总数, 有文本层, 无文本层)。

    只看前 3 页；本机没有解析库时无法判断，按「有文本层」计入并交给 01 步再判，
    避免只因探测失败就误报「全部是扫描件」。
    """
    total = text_layer = 0
    for path in paths:
        total += 1
        try:
            import pymupdf
            doc = pymupdf.open(str(path))
            try:
                if any(doc[i].get_text().strip() for i in range(min(len(doc), 3))):
                    text_layer += 1
            finally:
                doc.close()
        except ImportError:
            try:
                from pypdf import PdfReader
                reader = PdfReader(str(path))
                if any((p.extract_text() or "").strip() for p in reader.pages[:3]):
                    text_layer += 1
            except Exception:
                text_layer += 1
        except Exception:
            text_layer += 1
    return total, text_layer, total - text_layer


def write_intake_list(rows, params: dict | None = None,
                      detected: dict | None = None) -> Path:
    """生成人读的《放料清单.md》：要什么、放哪（绝对路径）、怎么命名、现在到没到。

    带 params/detected 时多一节「本次替你定的值」——把自动识别出来的项目名、
    方向、源报告摊开，写清依据与改法。用户不用猜哪里是猜的。
    """
    target = ROOT / "放料清单.md"
    ready = sum(1 for item, hits in rows if hits)
    lines = ["# 放料清单 · {}".format(SKILL.name), "",
             "> 本清单由 开工.py 自动生成，每次运行都会刷新。工作区：`{}`".format(norm(ROOT)),
             "> 放好文件后不用改本文件，重跑一次开工.py 即可。", "",
             "## 1. 现在到哪一步", "",
             "| 项 | 是否必需 | 状态 | 落位（绝对路径） |", "|---|---|---|---|"]
    for item, hits in rows:
        if hits:
            where = "、".join(norm(path.relative_to(ROOT)) if ROOT in path.parents else norm(path)
                              for path in hits[:2])
            status = "已到位（{}）".format(where)
        else:
            status = "**缺**" if item["required"] else "未提供（可选）"
        lines.append("| {} | {} | {} | `{}` |".format(
            item["name"], "必需" if item["required"] else "可选", status,
            norm(ROOT / item["dst"])))
    lines += ["", "已到位 {} / {} 项。".format(ready, len(rows)), ""]

    # ---------------- 本次替你定的值 ----------------
    if params:
        sug = []
        proj = (detected or {}).get("project") or {}
        direc = (detected or {}).get("direction") or {}
        src_type = params.get("project_source", "")
        if params.get("project"):
            if src_type == "auto-detect":
                mark = "（从材料里自动识别；依据：{}）".format(proj.get("basis") or "材料正文")
            elif src_type == "user":
                mark = "（你用 --project 指定的）"
            else:
                mark = "（你指定的）"
            sug.append(("项目全称", params["project"], mark))
        else:
            sug.append(("项目全称", "**待填**",
                        "（材料里没识别出来。用 --project \"正式全称\" 指定，"
                        "或让 AI 直接问你要）"))
        if params.get("direction"):
            if (detected or {}).get("direction", {}).get("value") == params.get("direction") \
                    and params.get("direction_source") == "auto-detect":
                mark = "（自动识别；依据：{}）".format(
                    "、".join(direc.get("evidence") or []) or "材料正文")
            else:
                mark = "（你指定的）"
            sug.append(("项目方向", params["direction"], mark))
        src_val = params.get("source_report")
        if src_val == ASSEMBLE_MARK:
            sug.append(("主线材料", "散材料自动汇编",
                        "（输入文件/ 里没有成篇现状分析稿，第 03 步按模板章节自动归位）"))
        elif src_val:
            sug.append(("主线材料", src_val, "（成篇 Word，直接当主线）"))
        if sug:
            lines += ["## 本次替你定的值（都能改）", "",
                      "| 项 | 取值 | 说明 |", "|---|---|---|"]
            for name, val, note in sug:
                lines.append("| {} | {} | {} |".format(name, val, note))
            lines += [""]
        # 另有项目名：只摊开事实与归属状态，不替用户判断并入还是排除
        cands = (detected or {}).get("scopeCandidates") or []
        if cands:
            lines += ["## 材料里还有这些项目名（归属按证据判，判不出就问你）", "",
                      "| 项目名 | 覆盖材料 | 归属 | 依据 |", "|---|---|---|---|"]
            for c in cands:
                d = c.get("decision") or "ask"
                by = c.get("decidedBy") or ""
                state = {"area": "本项目的一个区域", "separate": "独立项目，不入正文"}.get(d)
                if not state:
                    state = "**待你定**"
                elif by == "user":
                    state += "（你已定）"
                elif by == "evidence":
                    state += "（凭包内证据）"
                lines.append("| {} | {} 份 | {} | {} |".format(
                    c.get("name"), c.get("hits"), state,
                    "；".join(c.get("signals") or []) or "—"))
            lines += ["",
                      "判不了会当场问 A/B（证据摘要 + 建议），答复只对本次运行有效，"
                      "下一轮按当次材料重新判；不会按名字片段自动归入或排除。", ""]

    lines += ["## 2. 每一项怎么准备", ""]
    for index, item in enumerate(INTAKE_ITEMS, start=1):
        lines += ["### {}. {}".format(index, item["name"]),
                  "- 放到：`{}`".format(norm(ROOT / item["dst"])),
                  "- 怎么来：{}".format(item["how"]),
                  "- 干什么用：{}".format(item["why"]),
                  "- 缺失后果：{}".format("不开工，提示补料" if item["required"]
                                          else "照常开工，产物给初步判定（推断＋置信度），缺口进数据缺口清单"), ""]
    lines += ["## 3. 下一步", "",
              "```text",
              "python <技能目录>/scripts/cardo.py            # 直接校验并开工（参数会自动识别）".format(SKILL.name),
              "python <技能目录>/scripts/cardo.py --intake   # 一项一项问，确认落位".format(SKILL.name),
              "```", ""]
    target.write_text("\n".join(lines), encoding="utf-8")
    return target


def run_intake(rows) -> int:
    """一项一问：给绝对路径、给准备方法、确认是否已放好。返回 0 继续，3 待补料。"""
    if not sys.stdin or not sys.stdin.isatty():
        say("  当前不是交互终端，只输出清单；请在终端里重跑 --intake，或直接按清单放料。")
        return 3
    say("")
    def scan(item):
        out = []
        for pattern in item["patterns"]:
            out.extend(path for path in (ROOT / "输入文件").rglob(pattern)
                       if path.is_file() and not path.name.startswith("~$") and "空白" not in path.name)
        return sorted(set(out))

    say("")
    say("  ---- 收料引导：一项一项来。放好文件后按【回车】，我会重新扫一遍确认 ----")
    for item, hits in rows:
        if hits:
            say("  [已到位] {}：{}".format(item["name"], norm(hits[0].relative_to(ROOT))))
            continue
        say("")
        say("  需要：{}（{}）".format(item["name"], "必需" if item["required"] else "可选"))
        say("  放到：{}".format(norm(ROOT / item["dst"])))
        say("  怎么来：{}".format(item["how"]))
        say("  用途：{}".format(item["why"]))
        while True:
            try:
                answer = input("  → 放好后按【回车】确认；这份暂时没有就输入 s；退出输入 q：").strip().lower()
            except EOFError:
                # stdin 被重定向（如 < /dev/null）而 isatty() 判断漏了这种情形：
                # 没有输入可读，视为「暂时没有」，按该项的缺料口径处理，不要崩栈。
                answer = "s"
            if answer in ("", "y", "yes"):
                found = scan(item)
                if found:
                    say("  ✓ 扫到了 {} 个，例如：{}".format(len(found), norm(found[0].relative_to(ROOT))))
                    break
                say("  ✗ 还没扫到。核对两点：文件名里带上关键词（{}）、确实放在上面那个目录。".format(
                    "、".join(item["patterns"])))
                say("    改好后按【回车】再扫一次；确实没有就输入 s。")
                continue
            if answer in ("s", "n", "no", "没有"):
                if item["required"]:
                    say("  ! 这一项是硬门槛，缺了开不了工。先放料，再重跑。")
                    return 3
                say("  · 记为缺口：产物照样给初步判定，缺口进数据缺口清单，不影响开工。")
                break
            if answer in ("q", "quit", "exit"):
                say("  已退出收料引导；清单已保存到 放料清单.md")
                return 3
            say("  只认 回车 / s / q。")

    say("")
    say("  收料完成，继续校验与开工。")
    return 0


MIRRORS = [
    ("aliyun", "https://mirrors.aliyun.com/pypi/simple/", "mirrors.aliyun.com"),
    ("tsinghua", "https://pypi.tuna.tsinghua.edu.cn/simple", "pypi.tuna.tsinghua.edu.cn"),
    ("ustc", "https://mirrors.ustc.edu.cn/pypi/simple", "mirrors.ustc.edu.cn"),
    ("pypi", "https://pypi.org/simple", "pypi.org"),
]



# ---------------------------------------------------------------- 小工具

def say(msg: str = "") -> None:
    print(msg, flush=True)


def norm(p: Path) -> str:
    return str(p).replace("\\", "/")


def read_json(path: Path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def missing_mods(py: str, mods: list[str]) -> list[str]:
    code = ("import importlib.util,json,sys\n"
            "mods=json.loads(sys.argv[1])\n"
            "print(json.dumps([m for m in mods if importlib.util.find_spec(m) is None]))")
    try:
        r = subprocess.run([py, "-c", code, json.dumps(mods)], capture_output=True,
                           text=True, timeout=90)
        return json.loads(r.stdout.strip().splitlines()[-1]) if r.returncode == 0 else list(mods)
    except Exception:
        return list(mods)


def check_skill_package() -> list[str]:
    """技能包完整性自检：返回问题清单（空 = 完整）。

    为什么要有：技能包被复制不完整时（实测有人只拷了 开工.py + assets + runtime，
    Skills/00~09 整目录缺失），流水线会一路跑到 09 才失败，报错还不指向根因。
    这里在开工第一步就查清楚：**缺哪个目录、缺哪个脚本、缺哪个规则文件**，
    并给出补法。

    只读检查，不改任何文件。
    """
    problems: list[str] = []
    skills_root = SKILL / "Skills"
    if not skills_root.is_dir():
        return ["Skills/ 目录整体缺失（技能包的核心）。"
                "从交付包完整复制 CardoStatus/Skills/ 到 {}".format(norm(skills_root))]
    for step, scripts, rules in SKILL_STEPS:
        step_dir = skills_root / step
        if not step_dir.is_dir():
            problems.append("Skills/{}/ 目录缺失。".format(step))
            continue
        scripts_dir = step_dir / "scripts"
        if not scripts_dir.is_dir():
            problems.append("Skills/{}/scripts/ 目录缺失。".format(step))
        else:
            for s in scripts:
                if not (scripts_dir / s).is_file():
                    problems.append("Skills/{}/scripts/{} 缺失。".format(step, s))
        for r in rules:
            if not (step_dir / "规则" / r).is_file():
                problems.append("Skills/{}/规则/{} 缺失。".format(step, r))
    # 06 的六个章节规则副本（D/E/F 章的节标题与必答要点从中读取）
    for letter, name in (("A", "06-A-项目概述"), ("B", "06-B-片区基础概况分析"),
                         ("C", "06-C-现状详细分析"), ("D", "06-D-核心问题汇总"),
                         ("E", "06-E-现状综合研判与更新思路"), ("F", "06-F-附件")):
        if not (skills_root / name / "规则" / "chapter_skills.yaml").is_file():
            problems.append("Skills/{}/规则/chapter_skills.yaml 缺失。".format(name))
    return problems


# 便携环境必须能 import 的包（缺一个就有一类材料读不进来）。
# 这里的清单按「缺了会怎样」排：
#   docx/openpyxl   —— Word/Excel 材料读不出正文
#   yaml            —— 契约与规则文件读不出（06 步会静默降级成只剩 D/F 两章）
#   fitz            —— PDF 材料分不出「文书 / 图纸」（只影响归类，不影响正文）
#   numpy/cv2/PIL   —— 归类采样与识别依赖，缺了按文件名归类（项目信息卡读不出来）
RUNTIME_MODULES = ["docx", "openpyxl", "yaml", "fitz",
                   "numpy", "cv2", "PIL", "rapidocr_onnxruntime"]


def check_runtime() -> list[str]:
    """便携环境完整性自检：返回问题清单（空 = 完整）。

    为什么要有：实测有人复制工作区时**runtime 只拷了一部分**——
    site-packages 从 50 个包变成 12 个（numpy/onnxruntime/opencv/PyYAML 全丢），
    于是 import yaml 失败、按文件名归类失效、PDF 分不出类型。
    症状出现在很后面（06 步只生成 D/F 两章、OCR 静默失效），
    根因却在复制这一步。故在这里查，并说清「重新完整复制 runtime」。

    只读检查，不改任何文件。
    """
    problems: list[str] = []
    if not RUNTIME_PY.is_file():
        # 派生仓库不带 runtime：有系统 Python 时按 requirements.txt 建虚拟环境即可
        if Path(sys.executable).is_file():
            return problems
        problems.append("runtime/python/ 缺失（便携 Python 解释器），且未找到系统 Python。"
                        "从交付包完整复制 CardoStatus/runtime/ 到 {}".format(norm(SKILL / "runtime")))
        return problems
    miss = missing_mods(str(RUNTIME_PY), RUNTIME_MODULES)
    if miss:
        site = RUNTIME_PY.parent / "Lib" / "site-packages"
        problems.append(
            "runtime 的 site-packages 不完整，缺 {}（共 {} 个包）。"
            "多半是复制时漏了：从交付包**重新完整复制** CardoStatus/runtime/ "
            "覆盖到 {}".format("、".join(miss), len(RUNTIME_MODULES), norm(SKILL / "runtime")))
        if not (site / "yaml").is_dir():
            problems.append("特别提示：缺 PyYAML 时 06 步会**静默降级**"
                            "（章节映射读成空，只生成 D/F 两章且不报错）。")
    return problems


def report_runtime_problems(problems: list[str]) -> None:
    """把便携环境问题说清楚：缺什么、怎么补。"""
    say("[自检] 便携运行环境不完整：")
    for p in problems:
        say("  · {}".format(p))
    say("")
    say("  补法：从交付包重新完整复制整个 runtime 目录覆盖过去（推荐）：")
    say("     xcopy /E /I /Y <交付包>\\CardoStatus\\runtime <工作区>\\CardoStatus\\runtime")
    say("  补齐后重跑本命令即可；不需要重装依赖、不需要重建工作区。")


def report_skill_problems(problems: list[str]) -> None:
    """把完整性问题说清楚：缺什么、从哪补、怎么补。"""
    say("[自检] 技能包不完整，缺 {} 项：".format(len(problems)))
    for p in problems[:12]:
        say("  · {}".format(p))
    if len(problems) > 12:
        say("  · …另有 {} 项".format(len(problems) - 12))
    say("")
    say("  这些是交付包 CardoStatus/Skills/ 下的文件。补法：")
    say("  1) 从完整的交付包重新复制整个 Skills/ 目录覆盖过去（推荐）：")
    say("     xcopy /E /I /Y <交付包>\\CardoStatus\\Skills <工作区>\\CardoStatus\\Skills")
    say("  2) 或只补上面列出的缺失文件。")
    say("  补齐后重跑本命令即可；不需要重装依赖、不需要重建工作区。")


def venv_python() -> Path:
    name = "python.exe" if os.name == "nt" else "python"
    sub = "Scripts" if os.name == "nt" else "bin"
    return venv_dir() / sub / name


DELIVERY_MARKER = ".cardo_delivery_marker"


def _is_delivery_root(ws: Path) -> bool:
    """判断工作区是否为「交付库根」：交付库只读出包，不能在里面开工。

    唯一判据：**当前目录本身**存在标记文件 .cardo_delivery_marker。
    为什么不看「根下有 ≥2 个 Cardo* 目录」：用户经常把几个技能包并列放在
    同一个项目目录（如 测试-全流程贯通\）里使用，那是正常用法，不能拦。
    """
    try:
        return (ws / DELIVERY_MARKER).exists()
    except OSError:
        return False


def _has_multi_packages(ws: Path) -> bool:
    """并列多个 Cardo* 目录只作提示，不作拒绝条件。"""
    try:
        pkgs = [d for d in ws.iterdir() if d.is_dir() and d.name.startswith("Cardo")]
        return len(pkgs) >= 2
    except OSError:
        return False


def resolve_workspace(value: str) -> Path:
    """工作区 = 技能包所在的上一级目录（可 --workspace 覆盖）；技能包只读。"""
    ws = Path(value).resolve() if value else SKILL.parent
    if _is_delivery_root(ws):
        say(f"[错误] 这是交付库根目录，不能在这里开工：{ws}")
        say("        交付库是只读的发布产物，开工会在库里生成 过程文件/成果文件。")
        say("        正确做法：把技能包复制到项目目录后再运行，例如")
        say("          <项目目录> 下放 CardoStatus\\，再在该目录运行开工.py")
        raise SystemExit(2)
    if ws == SKILL or SKILL in ws.parents:
        say(f"[错误] 工作区不能是技能包目录或其子目录：{ws}")
        raise SystemExit(2)
    if not ws.is_dir():
        say(f"[错误] 工作区不存在：{ws}")
        raise SystemExit(2)
    if _has_multi_packages(ws):
        say("[提示] 当前目录并列了多个 Cardo* 技能包；未发现 .cardo_delivery_marker，按普通工作区继续。")
    if len(ws.parts) <= 1:
        # 只有当工作区本身就是盘符根（如 D://）时才拦；判断看自身层级，
        # 不能用 ws.parent.name —— 那样 D://任意文件夹 的父目录也是 D://，会被误拦。
        say("[错误] 技能包不能直接放在盘符根目录；请放进一个项目文件夹后再跑。")
        raise SystemExit(2)
    return ws


# ---------------------------------------------------------------- 1 目录

def make_dirs() -> list[str]:
    made = []
    for d in DIRS:
        p = ROOT / d
        if not p.is_dir():
            p.mkdir(parents=True, exist_ok=True)
            made.append(d)
    return made


# ---------------------------------------------------------------- 2 模板

def release_assets() -> list[str]:
    problems = []
    for a in ASSETS:
        src, dst = SKILL / a["src"], ROOT / a["dst"]
        if not src.is_file():
            problems.append(f"包内缺资产：{a['src']}（包被裁剪，需重新解压原包）")
            continue
        if dst.is_file():
            if a.get("sha256") and sha256_file(dst).upper() != a["sha256"].upper():
                problems.append(f"模板被改动：{a['dst']}（不要替换标准模板；恢复方法：删掉该文件后重跑本命令）")
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        say(f"  释放模板：{a['dst']}")
    for a in INPUT_TEMPLATES:
        src, dst = SKILL / a["src"], ROOT / a["dst"]
        if src.is_file() and not dst.is_file():
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            say(f"  释放输入模板：{a['dst']}")
    return problems


# ---------------------------------------------------------------- 3 依赖

def install(py: str) -> bool:
    req = SKILL / "requirements.txt"
    if not req.is_file():
        return True
    for name, url, host in MIRRORS:
        say(f"  安装依赖（{name}）…")
        r = subprocess.run([py, "-m", "pip", "install", "-r", str(req), "-i", url,
                            "--trusted-host", host, "--disable-pip-version-check", "--no-input"],
                           cwd=str(ROOT), capture_output=True, text=True, timeout=900)
        if r.returncode == 0:
            return True
        say(f"  {name} 失败，换下一个源。")
    return False


def system_pythons() -> list[str]:
    """找**真正的**系统 Python 候选（排除便携解释器本身）。

    为什么需要：便携解释器是内嵌的精简版，**没有 venv 模块**。
    若用它 `python -m venv` 必然失败，报错会指向"创建 venv 失败"，
    掩盖真正的问题（便携环境缺哪个包）。建 venv 必须用系统 Python。
    """
    out: list[str] = []
    self_path = str(RUNTIME_PY.resolve()) if RUNTIME_PY.is_file() else ""
    seen = set()

    def add(p: str) -> None:
        try:
            rp = str(Path(p).resolve())
        except OSError:
            return
        if rp in seen or (self_path and rp == self_path):
            return
        seen.add(rp)
        out.append(p)

    # ① 当前解释器不是便携的 → 它就是系统 Python
    if self_path and str(Path(sys.executable).resolve()) != self_path:
        add(sys.executable)
    # ② Windows 上的 python.exe（PATH 里）
    if os.name == "nt":
        for name in ("python.exe", "python3.exe"):
            found = shutil.which(name)
            if found:
                add(found)
        # ③ 常见的安装位置
        for base in (os.environ.get("LOCALAPPDATA", ""),
                     os.environ.get("ProgramFiles", ""),
                     r"C:\Python312", r"C:\Python311", r"C:\Python310"):
            if not base:
                continue
            for minor in range(9, 14):
                cand = Path(base) / "Programs" / "Python" / f"Python3{minor}" / "python.exe"
                if cand.is_file():
                    add(str(cand))
    else:
        for name in ("python3", "python"):
            found = shutil.which(name)
            if found:
                add(found)
    return out


def report_manual_install(missing: list[str]) -> None:
    """缺依赖时给出明确的手工补法（离线环境可用）。

    便携环境的 site-packages 路径是固定的，拷 wheel 进去即可，不需要建 venv、
    不需要联网——这是包内便携环境的本意。
    """
    site = RUNTIME_PY.parent / "Lib" / "site-packages"
    if os.name != "nt":
        site = RUNTIME_PY.parent / "lib" / "python3" / "site-packages"
    say("")
    say("  缺这些包：{}".format("、".join(missing)))
    say("  离线补法（任选其一）：")
    say("  A) 拷 wheel 到便携环境（推荐，无需联网）：")
    say("     把对应的 .whl 放进 {} 后重跑本命令".format(norm(site)))
    say("  B) 用系统 Python：")
    say("     python -m pip install {}（装到系统 Python 即可，流水线会自动改用它）".format(
        " ".join(missing)))
    say("  C) 有网络时：直接重跑本命令，会自动安装。")


def ensure_env(no_install: bool) -> tuple[str | None, bool]:
    """返回 (解释器, 是否刚装过)。解释器为 None 表示环境不可用。

    选解释器的顺序（2026-11-01 调整）：
      1) **包内便携环境**：拷包即用，不联网、不建 venv（首选）。
      2) **系统 Python**：便携环境缺依赖时优先用它——
         便携解释器是内嵌的精简版，**没有 venv 模块**，
         拿它 `python -m venv` 必然失败（实测 exit 2，且报错指向"创建 venv 失败"，
         掩盖了真正的问题：便携环境缺哪个包）。
      3) **联网安装**：系统 Python 也没有依赖时，才尝试装到便携环境（需联网）。
      4) 都不行 → 报清楚缺哪个包、怎么补（拷 wheel 到 runtime/Lib/site-packages）。
    """
    # 1) 包内便携环境：拷包即用，不联网、不建 venv
    if RUNTIME_PY.is_file():
        miss = missing_mods(str(RUNTIME_PY), IMPORTS)
        if not miss:
            return str(RUNTIME_PY), False
        say(f"  便携环境缺依赖：{'、'.join(miss)}")
    # 2) **系统 Python** 已有依赖 → 直接用。
    #    便携解释器缺 venv 模块时这是唯一出路（实测它 `python -m venv` 必失败）。
    sys_miss = list(IMPORTS)
    for cand in system_pythons():
        cmiss = missing_mods(cand, IMPORTS)
        if not cmiss:
            say(f"  改用系统 Python（依赖齐全）：{cand}")
            return cand, False
        sys_miss = cmiss
    # 3) 已有 .cardo/venv-status 且依赖齐全
    py = venv_python()
    if py.is_file() and not missing_mods(str(py), IMPORTS):
        return str(py), False
    if no_install:
        say(f"  缺依赖：{'、'.join(sys_miss)}（已指定 --no-install，不安装）")
        report_manual_install(sys_miss)
        return None, False
    # 4) 联网安装：装到 .cardo/venv-status。
    #    建 venv **必须用系统 Python**——便携解释器没有 venv 模块。
    if not py.is_file():
        say("  建虚拟环境 .cardo/venv-status …")
        venv_dir().parent.mkdir(parents=True, exist_ok=True)
        base = next((c for c in system_pythons()
                     if missing_mods(c, ["venv"]) == []), None)
        if base is None:
            say("  找不到带 venv 模块的系统 Python，无法自动建虚拟环境。")
            report_manual_install(sys_miss)
            return None, False
        r = subprocess.run([base, "-m", "venv", str(venv_dir())],
                           cwd=str(ROOT), capture_output=True, text=True, timeout=600)
        if r.returncode != 0:
            say("  创建虚拟环境失败：")
            say((r.stderr or r.stdout or "")[-800:])
            report_manual_install(sys_miss)
            return None, False
    if not install(str(py)):
        say("  依赖安装失败。离线环境可手动拷 wheel 后执行：")
        say("  python -m pip install --no-index --find-links <目录> -r requirements.txt")
        report_manual_install(sys_miss)
        return None, False
    left = missing_mods(str(py), IMPORTS)
    if left:
        say(f"  安装后仍缺：{'、'.join(left)}")
        return None, False
    return str(py), True


# ---------------------------------------------------------------- 4 参数与输入

def params_path() -> Path:
    return ROOT / PARAMS


def load_params() -> dict:
    data = read_json(params_path(), {}) or {}
    data.setdefault("schema", SCHEMA)
    for k, v in DEFAULTS.items():
        data.setdefault(k, v)
    return data


def docx_headings(p: Path) -> tuple[bool, str]:
    try:
        import zipfile as _zf
        with _zf.ZipFile(p) as z:
            if "word/document.xml" not in z.namelist():
                return False, "不是有效 docx（缺 word/document.xml）"
    except Exception as e:
        return False, f"无法读取 docx：{e}"
    try:
        from docx import Document
    except Exception:
        return True, "docx 结构有效（未装 python-docx，标题样式由流水线复核）"
    try:
        doc = Document(str(p))
        n = sum(1 for para in doc.paragraphs
                if (para.style.name or "").lower().startswith("heading") and para.text.strip())
        if n == 0:
            return False, "docx 没有标题样式（需 H1/H2），无法按章节映射"
        return True, f"docx 结构有效，标题段落 {n} 个"
    except Exception as e:
        return False, f"读取 docx 失败：{e}"


def find_sources() -> list[Path]:
    """投递区里的成篇材料（docx）：可作主线。

    只认 docx —— 因为只有 docx 能提供 H1/H2 章节结构，其余格式（pdf/xlsx/dwg…）
    属散材料，由 03 自动汇编。
    """
    root = ROOT / "输入文件"
    if not root.is_dir():
        return []
    return sorted([p for p in root.rglob("*.docx")
                   if p.is_file() and not p.name.startswith("~$")])


# 源报告只认「现状分析类文本」：现状底稿 / 旧报告 / 可研 / 片区策划（带现状章节）。
# AHP 结论、OSGB 体检报告、指标说明是上游成果与数据说明，不是底稿，
# 绝不能拿来当主线骨架——一份都没有时走「无源报告、按模板汇编」，不硬拉一份凑数。
NARRATIVE_HINTS = ("现状分析", "现状调查", "现状评估", "现状底稿", "现状报告",
                   "可行性研究", "可研", "项目建议书", "片区策划", "城市设计",
                   "更新策划", "旧改")
NON_NARRATIVE_HINTS = ("指标说明", "体检报告", "模型体检", "指标口径", "口径说明",
                       "数据说明", "成果说明", "权重", "评分", "矩阵", "一致性",
                       "依据对照", "结论", "ahp")


def is_narrative_source(path: Path) -> bool:
    """按「文件名 + 相对目录」判定能否当主线骨架（只有现状分析类文本可以）。"""
    try:
        rel = norm(path.relative_to(ROOT / "输入文件"))
    except ValueError:
        rel = path.name
    text = f"{path.name} {rel}".lower()
    if any(h.lower() in text for h in NON_NARRATIVE_HINTS):
        return False
    return any(h.lower() in text for h in NARRATIVE_HINTS)


def find_any_material() -> list[Path]:
    """投递区里的任意材料（不限格式）：用于判断「有没有放材料」。

    只判「有没有」，不判「能不能当主线」。空目录 → 阻塞并提示用户放料；
    非空 → 放行，让 01/03 去分类与汇编。
    """
    root = ROOT / "输入文件"
    if not root.is_dir():
        return []
    out = []
    for p in root.rglob("*"):
        if not p.is_file() or p.name.startswith("~$"):
            continue
        if p.name.startswith("."):
            continue
        out.append(p)
    return sorted(out)


# ---------------------------------------------------------------- 自动识别项目

def run_detect(py: str, project: str = "") -> dict:
    """调 detect_project.py 从材料里识别项目名与方向。

    先跑一次形式参考件识别：投递区里可能有「别的项目的同类成品」，它的标题
    会盖过本项目。参考件只学形式，不参与项目名识别（--exclude-json）。
    识别失败（脚本报错、目录为空、材料里读不出项目名）一律返回空 dict ——
    识别是「锦上添花」，不该成为新的失败点。
    """
    script = SKILL / DETECT
    if not script.is_file():
        return {}
    fr_json = ROOT / "过程文件" / "00-总控" / "形式参考" / "形式参考.json"
    fr_script = SKILL / "Skills/01-建档与材料归集/scripts/detect_form_reference.py"
    if fr_script.is_file():
        try:
            fr_json.parent.mkdir(parents=True, exist_ok=True)
            subprocess.run([py, str(fr_script), "--project-dir", str(ROOT),
                            "--delivery", str(ROOT / "输入文件"),
                            "--project", project or "", "--out", str(fr_json.parent)],
                           cwd=str(ROOT), capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=1200)
        except Exception:
            pass
    cmd = [py, str(script), "--input", str(ROOT / "输入文件")]
    if fr_json.is_file():
        cmd += ["--exclude-json", str(fr_json)]
    try:
        proc = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=600)
    except Exception:
        return {}
    if proc.returncode != 0:
        return {}
    try:
        data = json.loads((proc.stdout or "").strip() or "{}")
    except ValueError:
        return {}
    return data if isinstance(data, dict) else {}


def auto_fill_params(params: dict, detected: dict) -> list[str]:
    """用识别结果补齐 project / direction。返回说明行（写进放料清单）。"""
    notes = []
    if not detected:
        return notes
    proj = detected.get("project") or {}
    direc = detected.get("direction") or {}

    # project：只在用户没给、且识别出来时填。用户显式指定的永远优先。
    if not params.get("project") and proj.get("value"):
        params["project"] = proj["value"]
        params["project_source"] = "auto-detect"
        notes.append("项目名（自动识别）：{}".format(proj["value"]))
        if proj.get("basis"):
            notes.append("　依据：{}".format(proj["basis"]))
        if proj.get("ambiguous"):
            notes.append("　注意：材料里有多个项目名，本次按证据最强的写；"
                         "其余名字先判归属，判不出会停下来问你，不静默并入也不静默排除。")
    # direction：同上
    if not params.get("direction") and direc.get("value"):
        params["direction"] = direc["value"]
        params["direction_source"] = "auto-detect"
        notes.append("项目方向（自动识别）：{}（依据：{}）".format(
            direc["value"], "、".join(direc.get("evidence") or []) or "材料正文"))
    return notes


def scope_gate(detected: dict, answers: dict | None = None,
               areas: dict | None = None) -> tuple[list[str], list[dict]]:
    """项目归属闸门：材料里多出来的项目名，是本项目区域还是独立项目。

    证据够就判（写进 项目识别.json 的 scopeCandidates，附依据）；
    不够就当场问，绝不按名字片段静默排除、也不静默并入。
    答复只对**本次运行**有效（--scope-answer 传入），记在
    过程文件/00-总控/项目范围答复.json 备查，不做跨轮默认值。
    """
    notes, asks = [], []
    if areas:
        detected["areas"] = areas
        # 区域口径也按「本轮」记录，不落 记忆与规则
        try:
            write_json(ROOT / "过程文件/00-总控/项目范围答复.json",
                       {"schema": "cardo/scope-answer/v1",
                        "answeredAt": datetime.now().astimezone().isoformat(timespec="seconds"),
                        "注意": "本轮口径，仅供本轮归属判定；不做跨轮默认值。",
                        "answers": {}, "decisions": {}, "areas": areas})
        except OSError:
            pass
    cands = detected.get("scopeCandidates") or []
    if not cands:
        return notes, asks
    answers = {str(k): str(v) for k, v in (answers or {}).items()}
    rec = {"schema": "cardo/scope-answer/v1",
           "answeredAt": datetime.now().astimezone().isoformat(timespec="seconds"),
           "注意": "本轮答复，仅供本轮归属判定；下次运行按当次材料重新判，判不出再问。",
           "answers": {}, "decisions": {}, "areas": areas or {}}
    for c in cands:
        name = str(c.get("name") or "")
        val = str(answers.get(name) or "").strip().lower()
        if val in ("area", "separate"):
            c["decision"], c["decidedBy"] = val, "user"
            rec["answers"][name] = val
            rec["decisions"][name] = ("本项目的一个区域" if val == "area"
                                      else "独立项目，不入正文")
            notes.append("项目归属（本轮按你的答复）：{} → {}".format(
                name, rec["decisions"][name]))
        elif c.get("decision") == "area":
            c["decidedBy"] = "evidence"
            notes.append("项目归属（凭包内证据判为本项目区域）：{}｜依据：{}".format(
                name, "；".join(c.get("signals") or []) or "—"))
        elif c.get("decision") == "separate":
            c["decidedBy"] = "evidence"
            notes.append("项目归属（凭包内证据判为独立项目，不入正文）：{}".format(name))
        else:
            asks.append(c)
    # 判为“本项目的一个区域”的候选，自动并入 areas 关键词表。
    # 只写 areas，不改 scopeCandidates 的 decision；下一轮仍按当次材料重判。
    auto_areas = {}
    for c in cands:
        if str(c.get("decision") or "").lower() != "area":
            continue
        core = str(c.get("core") or c.get("name") or "").strip()
        core = re.sub(r"(片区|项目|改造|更新|工程)+$", "", core)
        if not core:
            continue
        area_name = core if core.endswith(("区域", "片区")) else core + "区域"
        auto_areas.setdefault(area_name, [])
        if core not in auto_areas[area_name]:
            auto_areas[area_name].append(core)
    if auto_areas:
        merged = dict(detected.get("areas") or {})
        for k, keys in auto_areas.items():
            cur = merged.setdefault(k, [])
            for key in keys:
                if key not in cur:
                    cur.append(key)
        detected["areas"] = merged
        rec["areas"] = merged
        notes.append("区域归属（已并入台账筛选关键词）：{}".format(
            "；".join("{}={}".format(k, "、".join(v)) for k, v in sorted(auto_areas.items()))))
    if rec["answers"] or auto_areas:
        try:
            write_json(ROOT / "过程文件/00-总控/项目范围答复.json", rec)
        except OSError:
            pass
    return notes, asks


def parse_scope_answers(items) -> dict:
    """把 --scope-answer "项目全称=area" 解析成 {项目名: area|separate}。"""
    out = {}
    for raw in items or []:
        text = str(raw or "").strip()
        if "=" not in text:
            raise SystemExit(
                "错误：--scope-answer 的写法是 \"项目全称=area\" 或 \"项目全称=separate\"，"
                "收到的是：{}".format(text))
        name, _, val = text.rpartition("=")
        val = val.strip().lower()
        if val not in ("area", "separate"):
            raise SystemExit("错误：--scope-answer 的取值只能是 area 或 separate，收到：{}".format(val))
        out[name.strip()] = val
    return out


def parse_areas(items) -> dict:
    """把 --area "示例片区B=示例社区B" 解析成 {区域名: [关键词]}。"""
    out = {}
    for raw in items or []:
        text = str(raw or "").strip()
        if "=" not in text:
            raise SystemExit(
                "错误：--area 的写法是 \"区域名=关键词,关键词\"，收到的是：{}".format(text))
        name, _, keys = text.partition("=")
        keys = [k.strip() for k in keys.replace("，", ",").split(",") if k.strip()]
        if not name.strip() or not keys:
            raise SystemExit("错误：--area 需要区域名和至少一个关键词，收到：{}".format(text))
        out[name.strip()] = keys
    return out


def precheck(params: dict, args) -> tuple[list[str], list[str]]:
    """返回 (阻塞项, 提示)。必要时回填 params。"""
    blockers, notes = [], []

    if not params.get("project"):
        blockers.append("缺参数 project：用 --project \"项目全称\" 指定")
    if not params.get("direction"):
        blockers.append(f"缺参数 direction：用 --direction 指定（{' / '.join(DIRECTIONS)}）")
    elif params["direction"] not in DIRECTIONS:
        blockers.append(f"direction 只能是 {' / '.join(DIRECTIONS)}")

    tf = params.get("template_file")
    if tf and not (ROOT / tf).is_file():
        blockers.append(f"模板不在位：{tf}（删掉后重跑本命令可恢复标准模板）")

    src = params.get("source_report")
    any_material = find_any_material()
    # 参数里钉死的源报告也要过类型闸门：AHP 结论/体检报告/指标说明不是现状底稿，
    # 不能因为上一轮写进参数就继续当主线骨架；直接忽略并改走自动判定。
    if (src and src != ASSEMBLE_MARK and (ROOT / "输入文件" / src).is_file()
            and not is_narrative_source(ROOT / "输入文件" / src)):
        notes.append("已忽略参数指定的源报告 {}：它不是现状分析类文本"
                     "（AHP 结论/体检报告/指标说明不能当主线骨架），改按材料自动判定".format(src))
        params.pop("source_report", None)
        src = None
    if src == ASSEMBLE_MARK:
        # 散材料模式（约定标记）：源报告由 03 步按材料自动汇编生成，
        # 不是投递区里的某个真实文件。此前预检不认这个标记，会把它当成
        # 「没指定源报告」，转去找 docx 候选，多份候选反而把散材料模式堵死。
        if not any_material:
            blockers.append(
                "散材料模式需要材料：请把项目材料放进 {}".format(
                    norm(ROOT / "输入文件") + "（任意层级，可放文件夹或压缩包）")
                + "，放好后重跑本命令。")
        else:
            notes.append("散材料模式（source_report=__assemble__）："
                         "03 步将把 输入文件/ 里的 {} 份材料按模板章节汇编成主线".format(
                             len(any_material)))
    elif src and (ROOT / "输入文件" / src).is_file():
        ok, msg = docx_headings(ROOT / "输入文件" / src)
        if not ok:
            blockers.append(f"源报告不可用：{src}，{msg}")
    else:
        all_docx = find_sources()
        # 只有现状分析类文本能当主线骨架；AHP 结论 / 体检报告 / 指标说明一律排除。
        cands = [p for p in all_docx if is_narrative_source(p)]
        skipped = [p for p in all_docx if p not in cands]
        if skipped:
            notes.append("已排除 {} 份非现状叙述材料（不当主线骨架，交 01 按材料类型归类）：{}".format(
                len(skipped), "、".join(norm(p.relative_to(ROOT / "输入文件")) for p in skipped)))
        good = []
        for p in cands:
            ok, msg = docx_headings(p)
            if ok:
                good.append(p)
            else:
                notes.append(f"{norm(p.relative_to(ROOT))} 不合格：{msg}")
        if len(good) == 1:
            rel = norm(good[0].relative_to(ROOT / "输入文件"))
            params["source_report"] = rel
            notes.append(f"已登记源报告：{rel}")
        elif not good:
            # 没有成篇材料也放行：03 会把 输入文件/ 里的散材料按模板章节汇编成主线
            if any_material:
                params["source_report"] = ASSEMBLE_MARK
                if skipped:
                    notes.append("未识别到可作主线的现状叙述材料（已排除 {} 份非叙事材料）；"
                                 "改走散材料汇编：把 输入文件/ 里的 {} 份材料按模板章节汇编成主线。"
                                 "若确有成篇现状报告，请用 --source 指定。".format(
                                     len(skipped), len(any_material)))
                else:
                    notes.append("没找到成篇材料：将把 输入文件/ 里的 {} 份材料按模板章节汇编成主线".format(
                        len(any_material)))
            else:
                # 投递区没有任何材料 → 必须停下来等用户放料。
                # 不能带着空目录往下跑：那样 01 会报「投递区为空」，看起来像流水线坏了；
                # 也不能去别处找材料——材料该由用户放到投递区。
                blockers.append(
                    "输入文件/ 里还没有材料：请把项目材料放进 {}".format(
                        norm(ROOT / "输入文件") + "（任意层级，可放文件夹或压缩包）")
                    + "，放好后重跑本命令。至少一份 docx / pdf / xlsx / dwg / rar 都算。")
        else:
            # 多份合格 docx：自动挑一份，不再逼用户二选一。
            # 挑法：优先文件名带「现状分析」的（那才是本流水线的主线形态），
            # 其次取最新修改的（通常是最新的定稿）。挑中的记进参数，
            # 其余在放料清单里摊开 —— 用户想换用 --source 指定即可。
            named = [p for p in good if "现状分析" in p.name]
            pool = named or good
            chosen = max(pool, key=lambda p: p.stat().st_mtime)
            rel = norm(chosen.relative_to(ROOT / "输入文件"))
            params["source_report"] = rel
            notes.append(f"有 {len(good)} 份合格 docx，自动选了一份当主线：{rel}"
                         + ("（文件名含「现状分析」）" if named else "（取最新修改的）"))
            for p in good:
                if p != chosen:
                    notes.append("　未选用：{}（想用它当主线加 --source \"{}\"）".format(
                        norm(p.relative_to(ROOT / "输入文件")), p.name))

    others = [p for p in find_sources() if "现状分析" not in p.name]
    if others:
        notes.append(f"另发现 {len(others)} 份项目资料 docx，01 会自动归类")
    if list((ROOT / "输入文件").rglob("*.pdf")):
        notes.append("发现 PDF：有文本层的抽全文作事实证据（带页码，不进正文）；扫描件不解析")
    return blockers, notes


# ---------------------------------------------------------------- 4b 联网补料

RESEARCH_DIR = "过程文件/00-总控/网络补料"


def _ledger_lib():
    """惰性导入台账校验模块（唯一实现：09/scripts/_ledger.py）。技能包不完整时返回 None。"""
    scripts = SKILL / "Skills" / "09-三重自检与边界" / "scripts"
    if str(scripts) not in sys.path:
        sys.path.insert(0, str(scripts))
    try:
        import _ledger
        return _ledger
    except ImportError:
        return None


def check_research() -> tuple[list[str], list[str]]:
    """校验《检索台账.json》：每条必须有 URL、检索日期、原文摘录与落盘文件。

    规则与 06/09 共用 _ledger.py——台账是"证据目录"，落盘文件才是证据本体：
    缺 URL / 缺检索日期 / 缺摘录 / file 不存在或为空，该条即无效（等同不存在）。
    历史项目的旧式 网络补料.json 仅在无台账时兜底校验。
    """
    lib = _ledger_lib()
    if lib is None:
        return [], ["台账校验模块缺失（09/scripts/_ledger.py）——技能包不完整"]
    led = lib.load(str(ROOT))
    if led["path"]:
        blockers = [f"检索台账 {p}" for p in led["problems"]]
        notes = [f"{eid}：{e.get('publisher', '')} {e.get('title', '')}"
                 f"（{e.get('date', '')}）"
                 for eid, e in led["valid"].items()]
        return blockers, notes
    p = ROOT / RESEARCH_DIR / "网络补料.json"
    if not p.is_file():
        return [], []
    data = read_json(p, {}) or {}
    items = data.get("items") or []
    blockers, notes = [], []
    for i, it in enumerate(items, 1):
        key = it.get("key") or f"补料#{i}"
        if it.get("not_found"):
            notes.append(f"{key}：标记查无公开数据（按缺口登记，不估算）")
            continue
        miss = [f for f in ("value", "unit", "source_url", "publisher", "date") if not it.get(f)]
        if miss:
            blockers.append(f"网络补料「{key}」缺字段：{'、'.join(miss)}")
            continue
        if not re.match(r"https?://", str(it["source_url"])):
            blockers.append(f"网络补料「{key}」source_url 不是 http(s) 链接")
            continue
        notes.append(f"{key}={it['value']}{it['unit']}（{it['publisher']} {it['date']}）"
                     + ("已确认" if it.get("confirmed_by_user") else "未确认，仅作参考值"))
    return blockers, notes


def research_guide(params: dict, blockers: list[str]) -> int:
    """生成《检索任务清单》：程序只定契约、不发网络请求，检索由 agent 执行后回填台账。"""
    d = ROOT / RESEARCH_DIR
    d.mkdir(parents=True, exist_ok=True)
    script = SKILL / "Skills" / "05-事实抽取与事实库" / "scripts" / "build_research_tasks.py"
    if script.is_file():
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
                   PYTHONDONTWRITEBYTECODE="1")
        try:
            r = subprocess.run([sys.executable, str(script), "--workspace", str(ROOT)],
                               cwd=str(ROOT), env=env, text=True, encoding="utf-8",
                               errors="replace", stdout=subprocess.PIPE,
                               stderr=subprocess.STDOUT)
            for ln in (r.stdout or "").splitlines():
                if ln.strip():
                    say(f"[检索] {ln}")
        except OSError as exc:
            say(f"[检索][错误] 任务清单生成失败：{type(exc).__name__}")
    else:
        say("[检索][错误] 缺少 build_research_tasks.py（技能包不完整）")
    if blockers:
        say("[检索] 当前阻塞项：")
        for b in blockers:
            say(f"[检索]   - {b}")
    say(f"[检索] 任务清单：{RESEARCH_DIR}/检索任务清单.md")
    say(f"[检索] 回填：{RESEARCH_DIR}/检索台账.json"
        "（每条须有 URL、检索日期、原文摘录与 file 落盘文件）")
    say("[检索] 程序不联网：检索由 agent 按清单执行，结果落盘并回填台账，正文用 [S-xxx] 引用")
    rb, rn = check_research()
    for x in rn:
        say(f"[检索] {x}")
    for x in rb:
        say(f"[检索][错误] {x}")
    return 3


# ---------------------------------------------------------------- 5 运行

def run_pipeline(cmd: list[str], env: dict) -> tuple[int, str]:
    """边跑边打印，同时留全文供判定停顿/失败位置。"""
    lines = []
    p = subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
    for line in p.stdout:
        say(line.rstrip("\n"))
        lines.append(line)
    return p.wait(), "".join(lines)


def cmd_hint(*args: str) -> str:
    """给用户看的一条可直接粘贴的命令。

    优先用包内便携 python（runtime/python）—— 没装 Python 的机器也能跑；
    包里没有 runtime 时才退回系统 python。
    """
    if RUNTIME_PY.is_file():
        try:
            exe = str(RUNTIME_PY.relative_to(ROOT))
        except ValueError:
            exe = str(RUNTIME_PY)
        if os.name == "nt":
            exe = exe.replace("/", "\\")
    else:
        exe = "python"
    if RUNTIME_PY.is_file():
        parts = [exe, "{}\\开工.py".format(SKILL.name)]
    else:
        parts = ["python", "<技能目录>/scripts/cardo.py"]
    parts.extend(args)
    return " ".join(parts)


def write_board_status() -> None:
    """把工作区真实进度写进看板的 WORKSPACE_STATUS 块（顶栏据此显示真实状态）。

    看板是单文件 HTML，双击打开读不到工作区里的文件，所以状态只能由 开工.py 注入。
    只替换标记之间的那一段，其余字节原样保留；写不进去（包只读、看板被换掉等）
    就跳过 —— 看板会退回手动勾选计数，不影响主流程。
    """
    if not WORKSPACE_RESOLVED:
        return                      # 工作区没解析成功，别拿默认 ROOT 写错状态
    board = SKILL / "开工看板.html"
    if not board.is_file():
        return
    try:
        with open(board, "r", encoding="utf-8", newline="") as f:
            text = f.read()
    except OSError:
        return
    if STATUS_BEGIN not in text or STATUS_END not in text:
        return
    try:
        status = {
            "updatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
            "materialsReady": (ROOT / "过程文件/01-建档与材料归集/分类清单/分类清单.json").is_file(),
            "reportRendered": bool(list(ROOT.glob("成果文件/*现状分析报告*.docx"))),
            "deliverables": sum(1 for pat in DELIVERABLES for _ in ROOT.glob(pat)),
        }
    except Exception:
        return
    # </script> 出现在 JSON 里会提前闭合脚本块 —— 转义 "<" 即可
    payload = json.dumps(status, ensure_ascii=False, indent=2).replace("<", "\\u003c")
    eol = "\r\n" if "\r\n" in text else "\n"
    block = (STATUS_BEGIN + eol + "<script>window.WORKSPACE_STATUS = "
             + payload + ";</script>" + eol + STATUS_END)
    head, _, rest = text.partition(STATUS_BEGIN)
    _, _, tail = rest.partition(STATUS_END)
    try:
        with open(board, "w", encoding="utf-8", newline="") as f:
            f.write(head + block + tail)
    except OSError:
        pass


CRASH_LOG = "过程文件/00-总控/异常日志.log"

# 异常类型 → (缺什么, 建议怎么处理)。只覆盖用户真会遇到的几类，
# 其余一律按「程序内部错误」处理，保留日志交给维护人。
_CRASH_HINTS = {
    "ModuleNotFoundError": ("缺少 Python 依赖（包内运行时可能不完整）",
                            "确认 runtime/ 与 Skills/ 是整体复制的；或用系统 Python 按 requirements.txt 装依赖"),
    "ImportError": ("运行环境不完整（缺模块或版本不符）", "重新完整复制 runtime/ 后重跑"),
    "FileNotFoundError": ("需要的文件不存在", "按上面给出的路径补齐后重跑"),
    "PermissionError": ("文件被占用或没有写权限",
                        "关掉正在打开该文件的 Word / Excel / 资源管理器预览窗格后重跑"),
    "UnicodeDecodeError": ("文件不是 UTF-8 编码", "用记事本另存为 UTF-8 后重跑"),
    "KeyError": ("输入表缺少必需的列或字段", "对照 放料清单.md 补齐该列后重跑"),
    "JSONDecodeError": ("JSON 文件内容损坏", "删掉该文件让程序重新生成，或恢复备份"),
}


def _last_step_hint() -> str:
    """从运行状态/运行日志里找最后一个进入的步骤，回答「哪一步失败」。"""
    for name in ("运行状态.json", "运行日志.json"):
        path = ROOT / "过程文件/00-总控" / name
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if not isinstance(data, dict):
            continue
        if name == "运行状态.json":
            steps = data.get("steps") or {}
            for key in list(steps)[::-1]:
                st = steps.get(key) or {}
                if isinstance(st, dict) and st.get("status") in ("running", "failed"):
                    return "{} {}".format(key, st.get("name", "")).strip()
        else:
            entries = data.get("entries") or []
            for item in entries[::-1]:
                if isinstance(item, dict) and item.get("step"):
                    return "{} {}".format(item.get("step"), item.get("name", "")).strip()
    return "开工准备阶段"


def crash_report(exc: BaseException) -> int:
    """未预期异常的统一出口：人话四段 + 落日志；--debug 时保留原始堆栈。"""
    # 用 format_exception 而不是 format_exc：后者不在 except 块里调用时只会写出
    # "NoneType: None"，日志等于没记。这里始终带上 exc 自身的堆栈。
    tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    log = ROOT / CRASH_LOG
    try:
        log.parent.mkdir(parents=True, exist_ok=True)
        with log.open("a", encoding="utf-8") as f:
            f.write("\n===== {} =====\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            f.write("技能包：{}\n".format(SKILL.name))
            f.write("命令：{}\n".format(" ".join(sys.argv)))
            f.write(tb)
    except OSError:
        log = None
    if "--debug" in sys.argv or os.environ.get("CARDO_DEBUG"):
        raise exc
    name = type(exc).__name__
    what, action = _CRASH_HINTS.get(
        name, ("程序内部错误", "把异常日志发给维护人；急用可先 --from <上一步> 续跑"))
    say("")
    say("[异常] 程序没有正常跑完")
    say("  哪一步失败：{}".format(_last_step_hint()))
    say("  缺什么：{}（{}: {}）".format(what, name, exc))
    say("  建议怎么处理：{}".format(action))
    say("  日志在哪：{}".format(norm(log.relative_to(ROOT)) if log else "（日志写入失败）"))
    say("  要看完整堆栈：命令末尾加 --debug 重跑")
    return 1


def main(argv=None) -> int:
    """跑完（含各种提前返回）后刷新看板顶栏的真实状态。"""
    try:
        return _main(argv)
    except SystemExit:
        raise
    except KeyboardInterrupt:
        say("\n[中断] 已停止；已完成步骤的产物保留，续跑用 --from <步>。")
        return 130
    except BaseException as exc:  # noqa: BLE001
        return crash_report(exc)
    finally:
        try:
            write_board_status()
        except Exception:
            pass


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=f"{TITLE} · 一条命令开工")
    ap.add_argument("--debug", action="store_true", help="出错时打印完整堆栈（默认只给人话提示）")
    ap.add_argument("--project", default="", help="项目全称（写入 工作区参数.json）")
    ap.add_argument("--direction", default="", help=f"项目方向：{' / '.join(DIRECTIONS)}")
    ap.add_argument("--source", default="", help="源报告文件名（输入文件/ 下）")
    ap.add_argument("--from", dest="start_from", default="", help="从第几步开始（续跑用，如 01）")
    ap.add_argument("--to", dest="stop_at", default="", help="跑到第几步为止（如 09）")
    ap.add_argument("--research", action="store_true",
                    help="缺料时生成《检索任务清单》：查什么、为什么、关键词（程序不联网，由 agent 执行）")
    ap.add_argument("--ledger", action="store_true",
                    help="只校验《检索台账.json》（URL/检索日期/摘录/落盘文件），不跑流水线")
    ap.add_argument("--intake", action="store_true",
                    help="收料引导：一项一项问，给绝对路径与准备方法，确认后再校验开工")
    ap.add_argument("--check", action="store_true", help="只做结构与输入校验，不跑流水线")
    ap.add_argument("--no-install", action="store_true", help="缺依赖时不安装")
    ap.add_argument("--edition", choices=["auto", "final", "gap"], default="auto",
                    help="成果版本：auto=按材料齐套情况自动定；final=正式版（正文有缺项会报错）；"
                         "gap=缺口版。文件名统一为《XX现状分析报告》")
    ap.add_argument("--dry-run", action="store_true", help="只打印将要执行的命令")
    ap.add_argument("--workspace", default="", help="工作区根（缺省为技能包所在目录，即上一级）")
    ap.add_argument("--scope-answer", action="append", default=[],
                    metavar="项目全称=area|separate",
                    help="本轮的项目归属答复（只对本次运行有效）；判不出时才会用到")
    ap.add_argument("--area", action="append", default=[], metavar="区域名=关键词,关键词",
                    help="本轮的区域归属口径，例如 \"示例片区B=示例社区B\"；"
                         "用于正文按区域分列数据，不做跨轮记忆")
    ap.add_argument("passthrough", nargs=argparse.REMAINDER, help="其余参数原样传给流水线")
    args = ap.parse_args(argv)

    if not os.environ.get("CARDO_REEXEC"):
        say(f"== {TITLE} · 开工 ==")
    if sys.version_info < PY_MIN:
        say(f"[环境] 需要 Python 3.{PY_MIN[1]} 及以上，当前 {sys.version.split()[0]}。")
        return 2

    global ROOT, WORKSPACE_RESOLVED
    ROOT = resolve_workspace(args.workspace)
    WORKSPACE_RESOLVED = True
    if not os.environ.get("CARDO_REEXEC"):
        say(f"[0/4] 工作区：{ROOT}（技能包只读：{SKILL.name}）")

    if args.ledger:
        lib = _ledger_lib()
        has = bool(lib and lib.load(str(ROOT))["path"])
        blockers, notes = check_research()
        for x in notes:
            say(f"[检索] {x}")
        for x in blockers:
            say(f"[检索][错误] {x}")
        if not has:
            say("[检索] 本项目未建立检索台账（检索台账.json）——按无外部数据项目处理")
        elif not blockers:
            say("[检索] 台账校验通过")
        return 0 if not blockers else 1

    # 技能包完整性自检：放在依赖检查之前。
    # 技能包被复制不完整时（实测只拷了 开工.py + assets + runtime，
    # 或 runtime 只拷了一部分——site-packages 从 50 个包变成 12 个），
    # 若不先查，后面会一路跑到 09 才失败，报错还不指向根因。
    skill_problems = check_skill_package()
    if skill_problems:
        if not os.environ.get("CARDO_REEXEC"):
            report_skill_problems(skill_problems)
        return 2
    runtime_problems = check_runtime()
    if runtime_problems:
        if not os.environ.get("CARDO_REEXEC"):
            report_runtime_problems(runtime_problems)
        return 2

    if not os.environ.get("CARDO_REEXEC"):
        board = SKILL / "开工看板.html"
        if board.is_file():
            say("      开工看板：{}（双击打开：步骤、放料清单、可复制提示词）".format(norm(board)))
            say("      放料目录：{}  ← 材料放这里（空目录时流水线会停下来等你放料）".format(
                norm(ROOT / "输入文件")))
            # 带路径的看板链接：打开即显示确切放料位置
            try:
                from urllib.parse import quote
                _board_url = "file:///" + norm(board).replace("\\", "/").replace(":", "%3A") \
                    + "#" + quote(norm(ROOT / "输入文件"))
                say("      带路径打开看板：{}".format(_board_url))
            except Exception:
                pass

    # 1 依赖（必要时切到 .venv 解释器重跑本脚本）
    py, installed = ensure_env(args.no_install)
    if py is None:
        return 2
    if py != sys.executable:
        say(f"[1/4] 依赖：改用 {py}")
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
                   PYTHONDONTWRITEBYTECODE="1", CARDO_REEXEC="1")
        r = subprocess.run([py, str(Path(__file__).resolve())] + list(sys.argv[1:]),
                           cwd=str(ROOT), env=env)
        return r.returncode
    if py == str(RUNTIME_PY):
        say("[1/4] 依赖：就位（包内便携环境，无需联网）")
    else:
        say("[1/4] 依赖：就位" + ("（已装入 .venv）" if installed else ""))

    # 2 目录与模板
    made = make_dirs()
    problems = release_assets()
    if problems:
        for p in problems:
            say(f"  ✗ {p}")
        return 2
    say(f"[2/4] 目录：{'新建 ' + '、'.join(made) if made else '已就绪'}；标准模板在位且 sha 一致")

    # 3 参数与输入
    rows = intake_scan()
    ready = sum(1 for item, hits in rows if hits)
    list_path = write_intake_list(rows)
    say(f"  放料清单：{norm(list_path.relative_to(ROOT))}（已到位 {ready}/{len(rows)} 项）")
    if args.intake:
        code = run_intake(rows)
        rows = intake_scan()
        write_intake_list(rows)
        if code != 0:
            say(f"  按清单放料后重跑：{cmd_hint()}，或再跑 --intake 逐项确认")
            return code

    say("[3/4] 参数与输入")
    params = load_params()
    if args.project:
        params["project"] = args.project
        params["project_source"] = "user"
    if args.direction:
        params["direction"] = args.direction
        params["direction_source"] = "user"
    if args.source:
        rel = norm(Path(args.source).as_posix().lstrip("/"))
        if (ROOT / "输入文件" / rel).is_file():
            params["source_report"] = rel
        else:
            hit = [p for p in find_sources() if p.name == Path(args.source).name]
            if len(hit) == 1:
                params["source_report"] = norm(hit[0].relative_to(ROOT / "输入文件"))
            else:
                params["source_report"] = rel

    # 自动识别项目名与方向：材料里往往已经写明了，不必反问用户。
    # 识别结果照常写进参数（后面成果文件要用），并在放料清单里摊开依据与改法。
    detected = run_detect(py, params.get("project", ""))
    if detected:
        try:
            out = ROOT / "过程文件/00-总控/项目识别.json"
            write_json(out, detected)
        except Exception:
            pass
    auto_notes = auto_fill_params(params, detected)
    write_json(params_path(), params)

    blockers, notes = precheck(params, args)
    rb, rn = check_research()
    blockers += rb
    notes += rn
    notes += auto_notes
    write_json(params_path(), params)
    for n in notes:
        say(f"  · {n}")
    # 归属闸门：多出来的项目名判不了就当场问，不替用户选边，也不做跨轮记忆
    scope_notes, scope_asks = scope_gate(detected, parse_scope_answers(args.scope_answer),
                                        parse_areas(args.area))
    for n in scope_notes:
        say(f"  · {n}")
    if detected:
        try:
            write_json(ROOT / "过程文件/00-总控/项目识别.json", detected)
        except Exception:
            pass
    # 放料清单刷新一次：这次带上「本次替你定的值」与各项目名的归属状态
    list_path = write_intake_list(rows, params, detected)
    if scope_asks:
        say("")
        say("  ✗ 有一件事只有你知道，材料里判不出来，先停下来问：")
        for c in scope_asks:
            say("")
            say("    「{0}」和「{1}」是什么关系？".format(c.get("name"), params.get("project")))
            say("      证据（都在投递区里）：")
            for s in (c.get("signals") or ["只在文件名里出现，没有可用证据"]):
                say("        · " + s)
            say("      A. 算本项目的一个区域 —— 并入，正文按区域归属写，数据不串")
            say("      B. 另属一个项目 —— 排除，不进正文")
            say("      建议：{0}".format(
                "倾向 A：结构证据指向同一轮摸底，但包内没有一句「实施范围包括该区域」的明文，"
                "按规定不替您定。" if c.get("score") else "证据不足，无法给倾向。"))
        say("")
        say("      怎么答：用本次参数带上你的选择重跑一次，例如")
        say('        {} --scope-answer "{}={}"'.format(
            cmd_hint(), scope_asks[0].get("name"), "area"))
        say("      答复只对本次运行有效（记在 {} 备查）；"
            "下次运行仍会按当次材料重新判、判不出再问。".format(
                norm(ROOT / "过程文件/00-总控/项目范围答复.json")))
        return 3
    if args.research:
        research_guide(params, blockers)
    if blockers:
        for b in blockers:
            say(f"  ✗ {b}")
        say(f"  放料目录（绝对路径）：{norm(ROOT / '输入文件')}")
        say(f"  逐项确认：{cmd_hint('--intake')}")
        say(f"  补齐后重跑：{cmd_hint()}")
        if not args.research:
            say(f"  或缺料想联网补：{cmd_hint('--research')}"
                "（给检索清单与来源要求，结果须标来源并经用户确认）")
        return 3
    say(f"  项目：{params['project']}｜方向：{params['direction']}｜源报告：{params['source_report']}")

    # PDF 口径：有文本层的抽全文供事实引用（不进正文）；扫描件不解析。
    # 输入里只有 PDF 或 PDF 占比高时必须当场说清，别让用户以为正文靠它就能成章。
    pdfs = [p for p in (ROOT / "输入文件").rglob("*")
            if p.is_file() and p.suffix.lower() == ".pdf"]
    text_mats = [p for p in (ROOT / "输入文件").rglob("*")
                 if p.is_file() and p.suffix.lower() in (".docx", ".xlsx", ".pptx", ".txt", ".md")]
    pdf_ratio = len(pdfs) / float(len(pdfs) + len(text_mats)) if pdfs else 0.0
    _, pdf_text_layer, pdf_scanned = pdf_layer_stats(pdfs)
    pdf_hint = ""
    if pdfs and (not text_mats or pdf_ratio >= 0.5):
        pdf_hint = ("检测到 {} 份 PDF（占材料 {:.0%}）：其中 {} 份有文本层，将抽取全文供事实引用"
                    "（不作文本结构来源、不进正文）；{} 份扫描件/无文本层，本版不解析。"
                    "现状正文仍需 Word 或文字版，否则正文可能不完整"
                    .format(len(pdfs), pdf_ratio, pdf_text_layer, pdf_scanned))
        say("  [!] " + pdf_hint)
    # 每次开工都覆盖写一遍：这次没有 PDF 问题就把上次的提示清掉，避免旧话一直在运行报告里
    write_input_hints(len(pdfs), pdf_ratio, pdf_hint, pdf_text_layer, pdf_scanned)

    if args.check:
        say("  校验通过（--check 不跑流水线）。")
        return 0

    # 4 流水线
    say("[4/4] 流水线")
    cmd = [py, str(SKILL / PIPELINE), "run", "--workspace", str(ROOT)]
    if args.start_from:
        cmd += ["--from", args.start_from]
    if args.stop_at:
        cmd += ["--to", args.stop_at]
    if args.edition and args.edition != "auto":
        cmd += ["--edition", args.edition]
    cmd += list(args.passthrough)
    say("  " + " ".join(cmd))
    if args.dry_run:
        return 0

    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
               PYTHONDONTWRITEBYTECODE="1", CARDO_SKILLS_ROOT=str(SKILL / "Skills"))
    code, text = run_pipeline(cmd, env)

    if code == 0:
        say("[完成] 全链通过。成果：")
        for pat in DELIVERABLES:
            for p in sorted(ROOT.glob(pat)):
                say(f"  - {norm(p.relative_to(ROOT))}")
        return 0

    failed = ""
    try:
        entries = (read_json(ROOT / "过程文件/00-总控/运行日志.json", {}) or {}).get("entries", [])
        failed = next((e.get("step") for e in reversed(entries) if e.get("exit") not in (0, None)), "") or ""
    except Exception:
        failed = ""
    stop = re.search(r"\[STOP\]\s*\[(\d{2})\]", text)
    if not failed and stop:
        failed = stop.group(1)

    if code == 3 or stop:
        say(f"[待确认] 停在需要人工介入的位置（步骤 {failed or '?'}）。")
        g = GUIDES.get(failed)
        if g:
            say(f"  做什么：{g}")
        say(f"  办完后续跑：{cmd_hint('--from ' + (failed or '01'))}")
        return 3

    say(f"[失败] 流水线退出码 {code}（步骤 {failed or '?'}）。")
    report = ROOT / "过程文件/00-总控/运行报告.md"
    if report.is_file():
        say("  运行报告末段：")
        for line in report.read_text(encoding="utf-8", errors="replace").splitlines()[-20:]:
            say("  " + line)
    say(f"  修好后续跑：{cmd_hint('--from ' + (failed or '01'))}（不要跳过失败步骤）")
    return code if code in (1, 2) else 1


if __name__ == "__main__":
    sys.exit(main())
