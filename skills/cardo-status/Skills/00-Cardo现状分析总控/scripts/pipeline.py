#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""CardoStatus 现状分析总控：九步流水线（run / status / scaffold）。

用法（工作区任意路径，无绝对路径硬编码）：
    python <技能目录>/Skills/00-Cardo现状分析总控/scripts/pipeline.py run
    python <技能目录>/Skills/00-Cardo现状分析总控/scripts/pipeline.py status
    python <技能目录>/Skills/00-Cardo现状分析总控/scripts/pipeline.py scaffold --dest <新工作区> \
        --project <项目名> --direction 老旧街区|老旧小区|老旧厂区 \
        --template <模板.docx> --source <原始报告.docx> [--materials <目录>]

退出码：0 全通过；1 业务不合格；2 参数或输入缺失；3 需人工确认/脚本自身异常；
        4 设计好的拒收（如无标注源、无法建立追溯链）。
"""
from __future__ import annotations

import argparse
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

WS = Path(__file__).resolve().parents[3]
# 技能链根：优先环境变量（工作区可在技能包之外），缺省为本包 Skills/
SKILLS = Path(os.environ.get("CARDO_SKILLS_ROOT")
              or Path(__file__).resolve().parents[2]).resolve()
P = WS / "过程文件"
S = SKILLS
OUT = WS / "成果文件"
D00 = P / "00-总控"
D01 = P / "01-建档与材料归集"
D02 = P / "02-齐套检查与缺料拍板"
D03 = P / "03-模板与源报告转储"
D04 = P / "04-章节映射"
D05 = P / "05-事实抽取与事实库"
D06 = P / "06-分章成文"
D07 = P / "07-表3.1口径计算"
D08 = P / "08-组装与模板渲染"
D09 = P / "09-三重自检与边界"

BLOCKING_CHECKS = {"CHK-HEAD-1", "CHK-HEAD-2", "CHK-HEAD-3", "CHK-LEN-1", "CHK-LEN-2",
                   "CHK-GAP-3", "CHK-PROC-1", "CHK-BAN-1", "CHK-BAN-2"}


def gates():
    """00 总控门禁配置（缺文件时用内置默认值，保证可移植/可裁剪）。"""
    return load_json(S / "00-Cardo现状分析总控" / "规则" / "门禁.json",
                     {"blockingChecks": sorted(BLOCKING_CHECKS),
                      "gapModeRelaxed": ["CHK-LEN-1", "CHK-LEN-2"],
                      "gapModeThreshold": 0.3, "traceCoverageMin": 1.0})


def _project_narrative_source():
    """项目侧正本：<工作区>/项目输入/项目叙述.json。

    只认工作区根以内的约定位置，不向上遍历父目录——包只碰工作区自己的东西，
    缺就当成"没有正本"走材料侧推导，不到外面找。
    """
    p = WS / "项目输入" / "项目叙述.json"
    return p if p.is_file() else None


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_json(path, default=None):
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return default if default is not None else {}


def dump_json(path, obj):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = str(path) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def rel(p):
    return os.path.relpath(p, WS).replace("\\", "/")


def prod_dir():
    """成品目录：一成果一子目录（只有《现状分析报告》进成品版）。"""
    d = OUT / "成品版" / "现状分析报告"
    d.mkdir(parents=True, exist_ok=True)
    return d


def internal_dir(name):
    """内部件目录：缺口清单、补齐确认单、证据包，逐件一子目录。"""
    d = OUT / "内部件" / name
    d.mkdir(parents=True, exist_ok=True)
    return d


def bind_workspace(path):
    """切换工作区（供 run/status --workspace 使用）；所有目录常量随之重绑。"""
    global WS, P, S, OUT, D00, D01, D02, D03, D04, D05, D06, D07, D08, D09
    WS = Path(path).resolve()
    P, S, OUT = WS / "过程文件", SKILLS, WS / "成果文件"
    D00, D01, D02, D03, D04 = (P / "00-总控", P / "01-建档与材料归集", P / "02-齐套检查与缺料拍板",
                               P / "03-模板与源报告转储", P / "04-章节映射")
    D05, D06, D07 = P / "05-事实抽取与事实库", P / "06-分章成文", P / "07-表3.1口径计算"
    D08, D09 = P / "08-组装与模板渲染", P / "09-三重自检与边界"


class ParamsError(Exception):
    pass


def normalize_zip(path):
    """固定 docx/xlsx 内部 ZIP 条目的时间戳，使同样内容的成稿 sha256 可复现。"""
    import zipfile
    tmp = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            zi = zipfile.ZipInfo(info.filename, date_time=(1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            dst.writestr(zi, data)
    tmp.replace(path)


def params():
    p = load_json(WS / "工作区参数.json", {})
    need = ["project", "direction", "template_file", "source_report"]
    miss = [k for k in need if not p.get(k)]
    if miss:
        raise ParamsError(f"[参数缺失] 工作区参数.json 缺少：{', '.join(miss)}")
    p.setdefault("delivery", "输入文件")
    return p


def tmpl(p=None):
    """标准模板：只在工作区根以内解析（提资数据/模板/…）。

    不查工作区的上级目录：模板要么由开工.py 释放到工作区，要么就是缺件。
    缺件时把应有的路径指名报出来并停下，不到包外去找。
    """
    p = p or params()
    path = WS / p["template_file"]
    if path.is_file():
        return path
    name = Path(p["template_file"]).name
    candidate = WS / "提资数据" / "模板" / name
    if candidate.is_file():
        return candidate
    raise SystemExit(
        "错误：找不到标准模板。工作区内应在 {0}（或参数里写的 {1}），实际没有。\n"
        "      补法：重跑 开工.py（会从包内 assets/ 释放模板），不要从工作区外面拿模板。"
        .format(candidate, path))


ASSEMBLE_MARK = "__assemble__"     # 「散材料自动汇编」约定标记（非真实文件）
ASSEMBLED_NAME = "素材汇编.docx"     # 03 汇编产物（工作区相对路径）


def src(p=None):
    """定位当前生效的源报告。

    * source_report == "__assemble__" → 散材料模式，源报告是 03 汇编产物；
    * 否则按「输入文件/<值>」→「工作区/<值>」→「输入文件里带『现状分析』的 docx」逐级找。
    找不到时返回 None（调用方须显式处理，不当文件用）。
    """
    p = p or params()
    value = p.get("source_report") or ""
    if value == ASSEMBLE_MARK:
        cand = WS / "过程文件" / "03-模板与源报告转储" / ASSEMBLED_NAME
        return cand if cand.is_file() else None
    if not value:
        return None
    path = WS / "输入文件" / value
    if path.is_file():
        return path
    alt = WS / value            # 允许相对工作区的路径（如 03 汇编出的 过程文件/…/素材汇编.docx）
    if alt.is_file():
        return alt
    # 用户可能把源报告放进文件夹或 zip（01 会先解包到 输入文件/解包-*）
    hits = sorted((WS / "输入文件").rglob("*现状分析*.docx")) if (WS / "输入文件").is_dir() else []
    return hits[0] if hits else None


def script(skill, name):
    return S / skill / "scripts" / name


def call(args, allow=(0,), capture=False):
    """执行子进程；返回 (code, output)。allow 之外的退出码由调用方处理。"""
    cmd = [sys.executable, *map(str, args)]
    print("  $ " + " ".join(str(x) for x in args[1:]))
    proc = subprocess.run(cmd, cwd=str(WS), capture_output=capture, text=True,
                          encoding="utf-8", errors="replace")
    code = proc.returncode
    out = (proc.stdout or "") + (proc.stderr or "")
    if out.strip():
        lines = out.strip().splitlines()
        for ln in lines[-24:]:
            print("    " + ln)
    return code, out


STEP_T0 = None
RUN_T0 = None
MATERIAL_CHOICE = "先出稿"
# 04 步的中等置信映射是否停下来等人工复核（默认否：自动采纳，不打断）
STRICT_REVIEW = False


def load_state():
    return load_json(D00 / "运行状态.json", {"schema": "cardo/run-state/v1", "steps": {}})


def save_step(state, key, name, code, notes="", outputs=None):
    duration = round(time.time() - STEP_T0, 1) if STEP_T0 else None
    entry = {
        "step": key, "name": name, "exit": code,
        "status": {0: "PASS", 3: "NEED_USER", 4: "REJECT"}.get(code, "FAIL"),
        "durationSeconds": duration, "notes": notes, "outputs": outputs or [],
        "at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    state.setdefault("steps", {})[key] = entry
    logs = load_json(D00 / "运行日志.json", {"schema": "cardo/run-log/v1", "entries": []})
    logs.setdefault("entries", []).append(entry)
    dump_json(D00 / "运行日志.json", logs)
    dump_json(D00 / "运行状态.json", state)
    write_run_report(state)


def write_run_report(state):
    steps = state.get("steps", {})
    total = round(time.time() - RUN_T0, 1) if RUN_T0 else None
    lines = ["# 运行报告（现状分析流水线）", "",
             f"- 工作区：{WS.name}（当前工作区）",
             f"- 项目：{params().get('project')} / {params().get('direction')}",
             f"- 版本：{state.get('variant', '未完成')}",
             f"- 开始：{state.get('startedAt', '—')}",
             f"- 结束：{state.get('finishedAt', '—')}",
             f"- 本次耗时：{total if total is not None else '—'} 秒", "",
             "| 步骤 | 状态 | 耗时(s) | 摘要 |", "| --- | --- | --- | --- |"]
    for key, _ in STEPS:
        st = steps.get(key)
        if not st:
            continue
        lines.append(f"| {key} {st['name']} | {st['status']} | "
                     f"{st.get('durationSeconds') or '—'} | {str(st.get('notes', '')).replace('|', '／')} |")
    lines.append("")
    # 开工阶段发现的口径问题（如 PDF 只有扫描件）留在这里，报告重写也不会丢
    hints = load_json(D00 / "输入提示.json", None)
    if isinstance(hints, dict) and hints.get("messages"):
        lines += ["## 输入提示", ""]
        for msg in hints["messages"]:
            lines.append("- " + str(msg))
        lines.append("")
    (D00 / "运行报告.md").write_text("\n".join(lines), encoding="utf-8")


class Gate(Exception):
    def __init__(self, msg, code=1):
        super().__init__(msg)
        self.code = code


# 《校验未通过说明》的定位：问题清单里已有「文件:行」，缺的按锚点编号回查分章稿。
_ANCHOR_TOKEN = re.compile(r"([A-F]-[A-Za-z0-9]+-P\d+)")


def _read_text_safe(path):
    try:
        return Path(path).read_text(encoding="utf-8-sig")
    except OSError:
        return ""


def _evidence_failures(d09):
    """从 09 的依据校验报告里取 (未通过的检查项, 问题清单)。"""
    files = sorted(Path(d09).glob("依据校验_*.md"))
    if not files:
        return [], []
    fails, problems, in_problem = [], [], False
    for ln in _read_text_safe(files[-1]).splitlines():
        s = ln.strip()
        if s.startswith("## 问题清单"):
            in_problem = True
            continue
        if in_problem:
            if s.startswith("- "):
                problems.append(s[2:].strip())
                continue
            if s.startswith("## "):
                in_problem = False
            continue
        if s.startswith("|") and "| 不通过 |" in s:
            cells = [c.strip() for c in s.strip("|").split("|")]
            if len(cells) >= 3:
                fails.append((cells[0], cells[2]))
    return fails, problems


def _locate_problem(problem, root):
    """给没有行号的问题补「文件:行」：按锚点编号回查分章稿。"""
    if not problem or re.search(r"\.md:\d+", problem):
        return problem
    m = _ANCHOR_TOKEN.search(problem)
    if not m or not root:
        return problem
    for f in sorted(Path(root).glob("*.md")):
        for i, ln in enumerate(_read_text_safe(f).splitlines(), 1):
            if m.group(1) in ln:
                return f"{f.name}:{i} {problem}"
    return problem


def _write_verify_note(title, checks, problems, code, extra="", locator_root=None):
    """在成果目录写《校验未通过说明》：逐条列未过项、卡在哪个文件哪一行。"""
    OUT.mkdir(parents=True, exist_ok=True)
    lines = [f"# 校验未通过说明（{title}）", "",
             f"- 生成时间：{datetime.now().astimezone().isoformat(timespec='seconds')}",
             f"- 结论：整链非 0 退出（exit={code}）",
             "- 退出码口径：1 业务不合格；4 设计好的拒收；3 脚本自身异常", ""]
    if checks:
        lines += ["## 一、未通过的检查项", "", "| 检查项 | 说明 |", "|---|---|"]
        lines += [f"| {n} | {d} |" for n, d in checks]
        lines.append("")
    if problems:
        lines += ["## 二、逐条定位（文件:行）", ""]
        for p in problems:
            lines.append(f"- {_locate_problem(str(p), locator_root)}")
        lines.append("")
    lines += ["## 三、处置", "",
              "- 修正后重跑本命令；校验收敛前不得交付或对外引用本版报告。"]
    if extra:
        lines += [f"- {extra}"]
    (OUT / "校验未通过说明.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"  已写《校验未通过说明》→ {OUT / '校验未通过说明.md'}")


# ------------------------------------------------------------------ 各步

def _unpack_zip(archive: Path, target: Path) -> str:
    import zipfile
    with zipfile.ZipFile(archive) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename.replace("\\", "/")
            if name.startswith("__MACOSX/") or "/__MACOSX/" in name:
                continue
            dest = (target / name).resolve()
            if target.resolve() not in dest.parents:
                continue
            dest.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(dest, "wb") as out:
                shutil.copyfileobj(src, out)
    return f"已解包 {archive.name} → {target.name}"


def _find_unrar() -> str:
    """定位 rar 解压工具：优先 UnRAR.exe（WinRAR 自带），其次 7z.exe，均可解 rar。"""
    cands = [
        r"C:\Program Files\WinRAR\UnRAR.exe",
        r"C:\Program Files\WinRAR\WinRAR.exe",
        r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
        r"C:\Program Files (x86)\WinRAR\WinRAR.exe",
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
        r"C:\ProgramData\chocolatey\tools\7z.exe",
    ]
    for c in cands:
        if os.path.isfile(c):
            return c
    for name in ("unrar", "7z", "7za", "rar"):
        found = shutil.which(name)
        if found:
            return found
    return ""


def _unpack_rar(archive: Path, target: Path) -> str:
    """解 .rar：调用本机 unrar/7z 命令行（无需 rarfile 库）。

    注意：UnRAR 的 -o 输出目录参数与含中文的路径组合时会「报成功但不落盘」，
    因此这里切到目标目录后用相对路径解压（已实测可靠）。
    解不开时抛异常，由调用方记成提示（不静默当作普通文件）。
    """
    exe = _find_unrar()
    if not exe:
        raise RuntimeError("本机未找到 rar 解压工具（UnRAR.exe / 7z.exe），"
                           "请先解包后再投，或把解包后的目录直接放进投递区")
    import subprocess
    name = os.path.basename(exe).lower()
    target.mkdir(parents=True, exist_ok=True)
    if name.startswith("7z"):
        cmd = [exe, "x", "-y", f"-o{target}", str(archive)]
        run_cwd = None
    else:
        # UnRAR：切到目标目录后解压（压缩包用绝对路径，输出落到当前目录=目标目录）
        cmd = [exe, "x", "-y", str(archive)]
        run_cwd = str(target)
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace",
                          cwd=run_cwd)
    out = (proc.stdout or "") + (proc.stderr or "")
    if proc.returncode != 0:
        raise RuntimeError(out.strip()[:300] or f"退出码 {proc.returncode}")
    # UnRAR 可能报成功却没落盘（中文路径场景）：以实际产物为准做校验
    produced = any(target.rglob("*"))
    if not produced:
        raise RuntimeError(f"解压命令已执行但未产出文件：{out.strip()[:200]}")
    return f"已解包 {archive.name} → {target.name}（{os.path.basename(exe)}）"


def unpack_delivery_zips(delivery):
    """把投递区里的压缩包解到 投递区/解包-<名字>/，支持用户直接投压缩包。

    支持 zip 与 rar；rar 通过本机 unrar/7z 命令行解压（不依赖第三方库）。
    解不开时如实记一条提示（不静默当作普通文件），由人工确认环节处理。
    """
    notes = []
    if not delivery.is_dir():
        return notes
    for pattern, unpacker in (("*.zip", _unpack_zip), ("*.rar", _unpack_rar)):
        for archive in sorted(delivery.rglob(pattern)):
            if archive.parent.name.startswith("解包-"):
                continue        # 跳过上次解出来的内容，避免重复解包
            target = archive.parent / ("解包-" + archive.stem)
            if target.exists():
                shutil.rmtree(target)
            target.mkdir(parents=True, exist_ok=True)
            try:
                notes.append(unpacker(archive, target))
            except Exception as exc:
                notes.append(f"解包失败 {archive.name}：{type(exc).__name__}: {exc}"
                             f"（请解包后再投，或确认包内内容已另存）")
    return notes


def step01(state):
    p = params()
    print("[01] 建档与材料归集")
    manifest = D01 / "分类清单" / "分类清单.json"
    # 零写入闸门：分类/落位前后原件哈希必须一致
    delivery = WS / p["delivery"]
    zip_notes = unpack_delivery_zips(delivery)
    for note in zip_notes:
        print("  " + note)
    def delivery_fingerprint(root):
        """投递区全量指纹：递归所有文件。

        原来只扫投递区顶层，而材料通常是成套文件夹投进来的，顶层没有文件 →
        指纹恒为空，"零写入闸门"等于没开，报告里还会写「原件 0 份」跟 09 的 43 份打架。
        这里递归统计；跳过 Office 临时件（~$）与 01 自己解压出来的 解包-*/。
        """
        out = {}
        for f in sorted(root.rglob("*")):
            if not f.is_file() or f.name.startswith("~$"):
                continue
            parts = f.relative_to(root).parts
            if any(part.startswith("解包-") for part in parts):
                continue
            out[f.relative_to(root).as_posix()] = sha256(f)
        return out

    before = delivery_fingerprint(delivery)
    # 形式参考件：投递区里「明显不是本项目、但同类报告成品」的 docx。
    # 只学骨架/句式，正文零引用；此处识别并登记，供 01 分类、03 汇编跳过，
    # 09 再拿它的专名与数值做成品串项扫描。
    fr_json = D00 / "形式参考" / "形式参考.json"
    code, _ = call([script("01-建档与材料归集", "detect_form_reference.py"),
                    "--project-dir", WS, "--delivery", p["delivery"],
                    "--project", p["project"], "--out", D00 / "形式参考"], allow=(0,))
    if code != 0:
        raise Gate("形式参考件识别失败（程序/环境错误）")
    code, _ = call([script("01-建档与材料归集", "classify_materials.py"),
                    "--project-dir", WS, "--delivery", p["delivery"],
                    "--out", rel(D01 / "分类清单"),
                    *(["--exclude-json", fr_json] if fr_json.is_file() else [])])
    if code != 0:
        raise Gate("材料分类失败")
    source_path = src(p)
    code, _ = call([script("01-建档与材料归集", "confirm_and_merge.py"),
                    "--project-dir", WS, "--manifest", manifest,
                    "--confirm", D01 / "分类确认.yaml",
                    "--rules", S / "01-建档与材料归集/规则/分类规则.yaml",
                    "--template", tmpl(p),
                    *(["--source", str(source_path)]
                      if source_path is not None and Path(source_path).is_file() else []),
                    "--project", p["project"], "--direction", p["direction"],
                    "--delivery", p["delivery"],
                    "--index-out", D00 / "项目索引.md",
                    "--fingerprint", D01 / "原件指纹.json"])
    if code == 3:
        raise Gate(f"分类存在待确认文件：请填写 {rel(D01 / '分类确认.yaml')} 后重跑 run", 3)
    if code != 0:
        raise Gate("建档/索引失败")
    after = delivery_fingerprint(delivery)
    if before != after:
        raise Gate("原始材料在 01 期间被改动/增删（零写入闸门）")
    save_step(state, "01", "建档与材料归集", 0,
              notes=f"原件 {len(before)} 份前后哈希一致",
              outputs=[rel(manifest), rel(D00 / "项目索引.md"), rel(D01 / "原件指纹.json")])


def step02(state):
    p = params()
    print("[02] 齐套检查与缺料拍板")
    code, _ = call([script("02-齐套检查与缺料拍板", "check_materials.py"),
                    "--project-dir", WS, "--direction", p["direction"],
                    "--project-name", p["project"],
                    "--classification", D01 / "分类清单" / "分类清单.json",
                    "--out", D02 / "齐套检查"], allow=(0, 1))
    gap = code == 1
    if code not in (0, 1):
        raise Gate(f"齐套检查执行失败（exit {code}）：参数/输入或程序错误，不是缺料判定；详见上方输出")
    gap_json = D02 / "齐套检查" / "材料齐套表.json"
    if not gap_json.exists():
        raise Gate("齐套表未生成")
    code, _ = call([script("02-齐套检查与缺料拍板", "make_material_request.py"),
                    "--project-dir", WS,
                    "--gap-json", gap_json,
                    "--out", internal_dir("输入材料补齐确认单") / "输入材料补齐确认单.docx",
                    "--md", D02 / "补齐确认单.md",
                    "--classification", D01 / "分类清单" / "分类清单.json",
                    "--project-id", D00 / "项目识别.json",
                    "--report-name", f"{p['project']}现状分析报告"])
    if code != 0:
        raise Gate("补齐确认单生成失败")
    # 缺料时显式拍板：可选「先补料」（exit 3 停下）或「先出稿」（干净版出稿，缺口另出清单）
    if gap:
        print(f"[02] 存在缺料：可选「先补料」（exit 3 停下，补齐后重跑）"
              f"或「先出稿」（干净版出稿，缺口进 数据缺口清单.xlsx）；"
              f"当前选择：{MATERIAL_CHOICE}")
    if gap and MATERIAL_CHOICE == "先补料":
        dump_json(D02 / "缺料拍板.json", {
            "schema": "cardo/material-decision/v1", "hasGap": True,
            "choice": "先补料",
            "reason": "按用户选择先补料：见 成果文件/输入材料补齐确认单.docx，补齐后重跑同一命令。",
            "at": datetime.now().astimezone().isoformat(timespec="seconds")})
        raise Gate("存在缺料且选择先补料：请按 输入材料补齐确认单 补齐后重跑；"
                   "如要先出稿，运行 run --material-choice 先出稿", 3)
    decision = {"choice": "先出稿（干净版）" if gap else "材料齐套",
                "reason": ("缺料不影响出稿：正文按推断给初步判定，缺口另出 数据缺口清单.xlsx，"
                           "后续补料再重跑。" if gap else "必备材料已齐套。"),
                "at": datetime.now().astimezone().isoformat(timespec="seconds")}
    dump_json(D02 / "缺料拍板.json", {"schema": "cardo/material-decision/v1",
                                     "hasGap": gap, **decision})
    save_step(state, "02", "齐套检查与缺料拍板", 0,
              notes="存在缺料，已拍板先出稿" if gap else "材料齐套",
              outputs=[rel(gap_json), rel(internal_dir("输入材料补齐确认单") / "输入材料补齐确认单.docx"),
                       rel(D02 / "缺料拍板.json")])


def _other_project_names() -> list[str]:
    """投递区里**其它**项目的全称（读 开工.py 写的 项目识别.json）。

    识别结果里有全部候选项目名与「本次选哪个」；这里取其余的，
    供 03 步把它们排除在本项目之外。没有识别结果时返回空（不排除任何材料）。
    """
    data = load_json(D00 / "项目识别.json", {})
    if not data:
        return []
    chosen = ((data.get("project") or {}).get("value") or "").strip()
    names = []
    for item in data.get("candidates") or []:
        name = str(item.get("name") or "").strip()
        if name and name != chosen and name not in names:
            names.append(name)
    # 只取真正有材料支撑的候选（hits ≥1），免得把纯噪声塞进排除词
    return [n for n, item in
            zip(names, [c for c in (data.get("candidates") or [])
                        if str(c.get("name") or "").strip() in names])
            if int(item.get("hits") or 0) >= 1]


def _area_names() -> list[str]:
    """本轮确认纳入本项目实施范围的区域（开工.py 范围闸门写进 项目识别.json）。

    区域（如示例片区B）是本项目的一块实施范围而不是别的项目：它的台账与
    沟通稿要一并进汇编，否则 06 读不到区域台账，报告只能写成一个片区。
    """
    data = load_json(D00 / "项目识别.json", {})
    names: list[str] = []
    for area, keys in (data.get("areas") or {}).items():
        for name in [str(area)] + [str(k) for k in (keys or [])]:
            name = name.strip()
            if name and name not in names:
                names.append(name)
    return names


def _in_area(name: str, areas: list[str]) -> bool:
    """候选项目名是否就是本轮确认的区域（去「项目/区域/工程」尾缀后同一专名）。"""
    def core(x):
        return re.sub(r"(项目|区域|工程)$", "", str(x or "").strip())
    c = core(name)
    return bool(c) and any(c == core(a) for a in areas)


def step03(state):
    print("[03] 模板与源报告转储")
    p = params()
    code, _ = call([script("03-模板与源报告转储", "dump_docx_outline.py"), tmpl(p),
                    "--out", D03 / "模板转储", "--tables"])
    if code != 0:
        raise Gate("模板转储失败")
    source = src(p)
    coverage = D03 / "素材汇编.docx.覆盖.json"
    if p.get("source_report") == ASSEMBLE_MARK or source is None:
        # 结构由模板定：把散材料按模板章节汇编成主线，不要求用户自己攒 Word
        assembled = D03 / ASSEMBLED_NAME
        # 投递区常装着多个项目的材料（课题提资包、批次移交…）：
        # 把**其它项目**的名字一并告诉汇编器，让它把这些项目的材料排除在外。
        # 本轮确认的区域也算本项目边界：写进 --project，并从「其它项目」里摘掉。
        areas = _area_names()
        other_projects = ",".join(
            n for n in _other_project_names() if not _in_area(n, areas))
        fr_json = D00 / "形式参考" / "形式参考.json"
        code, _ = call([script("03-模板与源报告转储", "assemble_materials.py"),
                        "--input", WS / "输入文件", "--template", tmpl(p), "--out", assembled,
                        # 传项目名：投递区常装着多个项目的材料，按项目边界过滤，
                        # 避免把别的项目的数据写进本项目现状报告。
                        "--project", "/".join([p.get("project", "")] + areas),
                        *(["--other-projects", other_projects] if other_projects else []),
                        # 形式参考件只学骨架，不进汇编正文
                        *(["--exclude-json", fr_json] if fr_json.is_file() else [])])
        if code != 0:
            raise Gate("材料汇编失败（把散材料按模板章节归位）")
        # source_report 保持 "__assemble__" 标记不变：它是「本项目走散材料模式」的约定，
        # 每次跑 03 都应按当前材料重新汇编；把路径写回参数会让后续跑误当成成篇源报告。
        source = assembled
        # 散材料模式标志：源报告由散材料汇编而成，不是成篇现状分析。
        # 06 据此改走「写作任务包 + agent 回填」而非「摘句搬运」。
        state["scatterMode"] = True
        # 覆盖度清单：记录哪些章节真有料，04 据此判真实缺口
        dump_json(D00 / "汇编覆盖度.json", load_json(coverage, {}))
    code, _ = call([script("03-模板与源报告转储", "dump_docx_outline.py"), source,
                    "--out", D03 / "源报告转储", "--tables"])
    if code != 0:
        raise Gate("源报告转储失败")
    tpl_outline = next(Path(D03 / "模板转储").glob("*.outline.json"))
    src_outline = next(Path(D03 / "源报告转储").glob("*.outline.json"))
    code, _ = call([script("03-模板与源报告转储", "build_skeleton.py"),
                    "--template", tmpl(p), "--template-outline", tpl_outline,
                    "--source-outline", src_outline,
                    "--contract", S / "03-模板与源报告转储/规则/章节契约.yaml",
                    "--out", D03])
    if code != 0:
        raise Gate("章节骨架/模板指纹校验失败（可能是旧模板）")
    save_step(state, "03", "模板与源报告转储", 0,
              outputs=[rel(tpl_outline), rel(src_outline), rel(D03 / "章节骨架.json")])


def step04(state):
    p = params()
    print("[04] 章节映射")
    m = D04 / "章节映射.yaml"
    source = src(p)
    if source is None:
        raise Gate("找不到源报告：03 步骤应已生成 过程文件/03-模板与源报告转储/素材汇编.docx")
    code, _ = call([script("04-章节映射", "make_mapping_skeleton.py"),
                    "--template", tmpl(p), "--source", source, "--out", D04 / "_骨架.yaml"])
    if code != 0:
        raise Gate("映射骨架生成失败")
    src_outline = next(Path(D03 / "源报告转储").glob("*.outline.json"))
    coverage = D03 / "素材汇编.docx.覆盖.json"
    code, _ = call([script("04-章节映射", "fill_mapping.py"),
                    "--skeleton", D04 / "_骨架.yaml", "--source-outline", src_outline,
                    "--out", m, "--confirm", D04 / "人工确认.yaml",
                    "--direction", p["direction"],
                    *(["--coverage", coverage] if coverage.is_file() else []),
                    *(["--strict-review"] if STRICT_REVIEW else [])])
    if code == 3:
        raise Gate("存在中等置信映射（提案），请在 过程文件/04-章节映射/人工确认.yaml 逐条复核："
                   "保留则填 sources，判为缺项则写 sources: []", 3)
    if code != 0:
        raise Gate("映射填充失败")
    code, out = call([script("04-章节映射", "make_mapping_skeleton.py"), "--check", m,
                      "--template", tmpl(p), "--source", source], capture=True)
    if code != 0:
        # 把自检的具体问题带上，并给出最可能的处置（人工确认失效 / 源标题对不上）
        detail = [ln.strip() for ln in out.splitlines()
                  if ln.strip() and (ln.strip()[0].isdigit() or "FAIL" in ln)]
        hint = ""
        if any("在 source 大纲中找不到" in ln for ln in detail):
            hint = ("；多半是 人工确认.yaml 里还留着上一轮源报告的映射。"
                    "请核对该文件（无效条目会在 [04][提示] 里点名），"
                    "删掉失效条目或按当前源报告重填")
        raise Gate("章节映射自检未通过" + (("：" + "；".join(detail[:6])) if detail else "") + hint)
    stats = load_json(D04 / "映射统计.json", {})
    total = max(1, sum(stats.get(k, 0) for k in ("exact", "alias", "gap")))
    ratio = stats.get("gap", 0) / total
    # 缺口版判定：
    #   ① 真实材料缺口（已剔除结构节点）≥ 1 个 → 启用；
    #   ② 或缺口占比 ≥ 阈值 → 启用。
    # 口径说明：早先用「缺口节数 / 全部节点」做阈值判断，修正装配器把 H1 的材料
    # 归入 H2 子节后，节点级缺口数被高估、真实缺口又被低估，两个口径都会误判。
    # 现在以「真实材料缺口节点」为准——只要有一节确实没有材料，成稿就带缺口，
    # 就应走缺口版（逐项留痕 + 放宽字数门禁），而不是按完整版严检后卡死。
    material_gaps = stats.get("materialGapCount", stats.get("gap", 0))
    threshold = float(gates().get("gapModeThreshold", 0.3))
    # v2 结论优先：缺料不再自动升级为「缺口版」。
    # 默认一律出干净版（缺口只进 数据缺口清单.xlsx，正文照样给初步判定）；
    # 只有显式 --edition gap 才出缺项留痕版。
    state["gapMode"] = (EDITION == "gap")
    state["materialGapCount"] = material_gaps
    state["edition"] = "gap" if state["gapMode"] else "clean"
    save_step(state, "04", "章节映射", 0,
              notes=f"high={stats.get('exact',0)} medium={stats.get('alias',0)} "
                    f"gap={stats.get('gap',0)}（真实材料缺口 {material_gaps} 个）；"
                    f"缺口占比 {ratio:.0%}"
                    + ("（启用缺口版）" if state["gapMode"] else ""),
              outputs=[rel(m)])


def step05(state):
    print("[05] 事实抽取与事实库")
    code, _ = call([script("05-事实抽取与事实库", "extract_facts.py"),
                    "--project-dir", WS,
                    "--dump-dir", D03 / "源报告转储",
                    "--mapping", D04 / "章节映射.yaml",
                    "--out", D05 / "facts"])
    if code != 0:
        raise Gate("事实抽取失败")
    facts_path = D05 / "facts" / "事实库.json"
    facts = load_json(facts_path)
    n = len(facts.get("facts", []))
    if n == 0:
        raise Gate("事实库为空（源报告无可抽取事实）")
    # 可移植：事实来源路径一律落成工作区相对 posix 路径，不写盘符
    for f in facts.get("facts", []):
        s = f.get("source") or {}
        fp = str(s.get("file") or "")
        if fp:
            s["file"] = rel(fp) if os.path.isabs(fp) else fp.replace("\\", "/")
    dump_json(facts_path, facts)
    # 包内依据接入：OSGB 现状指标 + AHP 结论（只认原值，不重算），生成 包内依据.json/.md 与同源核对
    code, _ = call([script("05-事实抽取与事实库", "ingest_baseline.py"),
                    "--project-dir", WS,
                    "--out", D05 / "包内依据"])
    if code != 0:
        raise Gate("包内依据接入失败")
    # 标准限值/测绘参数并入包内依据：必要性论证要对照的规范限值（25%、4 米、1/12、
    # 30%—40%、0.25 米）不是本项目实测数，必须有一个明确出处，否则会被
    # 「数值有源」判为无源。资产生成、幂等追加，不覆盖 05 的原始产出。
    limits = S.parent / "assets" / "依据库" / "标准限值.md"
    md_ev0 = D05 / "包内依据" / "包内依据.md"
    if limits.is_file() and md_ev0.is_file():
        body = md_ev0.read_text(encoding="utf-8")
        text = limits.read_text(encoding="utf-8").strip()
        if "## 标准限值与测绘参数" not in body:
            with open(md_ev0, "a", encoding="utf-8") as f:
                f.write("\n\n" + text + "\n")
    # 编制依据运行时联网核验：断网不阻断出稿，但必须留痕并在依据表标“未联网复核”
    basis_lib = S.parent / "assets" / "依据库" / "依据库.json"
    if basis_lib.is_file():
        code, _ = call([script("06-分章成文", "net_verify_citations.py"),
                        "--lib", basis_lib, "--out-dir", D05 / "包内依据",
                        "--timeout", "8", "--workers", "8"])
        if code != 0:
            raise Gate("编制依据联网核验脚本执行失败（程序/环境错误，不是断网）")
        os.environ["CARDO_BASIS_LIB"] = str(basis_lib)
        os.environ["CARDO_BASIS_VERIFY"] = str(D05 / "包内依据" / "依据核验记录.json")
    else:
        print("  [依据] 包内未找到 assets/依据库/依据库.json，按缺库口径退回“待补文号”")
    baseline = load_json(D05 / "包内依据" / "包内依据.json", {})
    n_metric = len((baseline.get("osgb") or {}).get("metrics") or [])
    # 包内依据是正文指标的唯一原值出处：并入来源转储，供 CHK-NUM-1 回溯（不并入则
    # 正文引用 OSGB 实测值会被判「无源数字」）。
    md_ev = D05 / "包内依据" / "包内依据.md"
    blocks_path = D03 / "源报告转储" / "素材汇编.blocks.txt"
    if md_ev.is_file() and blocks_path.is_file():
        old = blocks_path.read_text(encoding="utf-8")
        marker = "【来源：包内依据.md（OSGB 现状指标 / AHP 结论）】"
        if marker not in old:
            body = md_ev.read_text(encoding="utf-8")
            blocks_path.write_text(old.rstrip() + "\n\n" + marker + "\n" + body + "\n",
                                   encoding="utf-8")
    save_step(state, "05", "事实抽取与事实库", 0,
              notes=f"事实 {n} 条；包内依据接入：OSGB 指标 {n_metric} 项"
                    + ("、AHP 结论已接入" if (baseline.get("ahp") or {}) else ""),
              outputs=[rel(facts_path), rel(D05 / "包内依据" / "包内依据.json"),
                       rel(D05 / "包内依据" / "同源核对.json")])


def _has_filled_drafts(filled_dir) -> bool:
    """判断写作任务目录里是否已有 agent/大模型回填的章节稿。

    判定条件：存在 .md 文件，且已写入实际正文（模板提示注释已清除）。
    """
    d = Path(filled_dir)
    if not d.is_dir():
        return False
    for f in sorted(d.glob("*.md")):
        try:
            text = f.read_text(encoding="utf-8-sig")
        except OSError:
            continue
        if not text.strip():
            continue
        if "<!-- 请由大模型成文后写入本文件" in text:
            continue
        if "# " in text:
            return True
    return False


def _refresh_edition(state):
    """按 --edition 重新定版。

    为什么不能直接沿用 state：`--from 06` 这类断点续跑会跳过 04 步，
    state 里留着上一次运行的 gapMode=True，会导致干净版被当成缺口版出稿
    （实测：续跑时 06 仍带 --gap-mode）。v2 起缺料不再自动升版，
    因此每次进入 06/08 都按 EDITION 重算一次。
    """
    state["gapMode"] = (EDITION == "gap")
    state["edition"] = "gap" if state["gapMode"] else "clean"
    state["variant"] = "缺口版" if state["gapMode"] else "干净版"


def step06(state):
    _refresh_edition(state)
    print("[06] 分章成文")
    p = params()
    task_dir = D06 / "写作任务"
    filled_dir = task_dir / "分章初稿_待回填"

    # ── 两段式：先出「写作任务包」，等 agent/大模型回填，再消费回填稿 ──
    # 脚本保持零 LLM 调用、离线可跑；成文质量由大模型按 skill 规范保证。
    need_write = bool(state.get("scatterMode")) and not _has_filled_drafts(filled_dir)
    # 降级稿标记：供 09 把「篇幅/依据块/契约」类检查降为提示级（降级稿天然达不到成文稿指标），
    # 但 06 的标题/禁句/缺项标记等结构性检查照旧拦截。
    state["degraded"] = need_write
    if need_write:
        code, _ = call([script("06-分章成文", "build_write_tasks.py"),
                        "--project-dir", WS, "--out", task_dir])
        if code != 0:
            raise Gate("写作任务包生成失败")
        state["writeTaskPending"] = True
        dump_json(D00 / "运行状态.json", state)
        # 不阻断：任务包供 agent/大模型**可选**改写；未回填时本步按确定性散材料成文降级。
        # （原实现 raise Gate 会让散材料项目每次运行都停在 06，永远跑不到 09 与三件套。）
        print("[06] 写作任务包已生成；未收到回填稿，按确定性成文降级")

    args = [script("06-分章成文", "gen_chapters.py"),
            "--template-outline", next(Path(D03 / "模板转储").glob("*.outline.json")),
            "--source-dump", D03 / "源报告转储",
            "--mapping", D04 / "章节映射.yaml",
            "--materials", D01 / "分类清单" / "分类清单.json",
            "--facts", D05 / "facts" / "事实库.json",
            "--out", D06 / "分章初稿"]
    baseline_json = D05 / "包内依据" / "包内依据.json"
    if baseline_json.exists():
        args += ["--baseline", baseline_json,
                 "--baseline-md", rel(D05 / "包内依据" / "包内依据.md"),
                 "--project", p["project"]]
        # 项目叙述：正本在项目根 项目输入/项目叙述.json；没有正本时从材料侧
        # （项目识别 + 台账转储 + 包内依据）推导一份，推不出的登记「叙述缺口」，
        # 不因缺叙述而 Gate 整条流水线。
        src = _project_narrative_source()
        narrative = WS / "记忆与规则" / "项目叙述.json"
        if src:
            code, _ = call([script("06-分章成文", "build_project_narrative.py"),
                            "--src", src, "--out", narrative,
                            "--project", p["project"],
                            "--dump", D03 / "源报告转储",
                            "--weights", baseline_json])
            if code != 0:
                raise Gate("项目叙述配置校验失败（见上方缺字段/不匹配项）")
        else:
            print("[06] 无项目叙述正本，按材料侧推导（推不出的登记缺口，不阻断）")
            code, _ = call([script("06-分章成文", "build_project_narrative.py"),
                            "--derive", "--out", narrative,
                            "--project", p["project"],
                            "--dump", D03 / "源报告转储",
                            "--baseline", baseline_json])
            if code != 0:
                print("[06] 项目叙述推导脚本异常（exit $code），继续尝试按现有配置成文")
        if narrative.exists():
            args += ["--narrative", narrative]
    if state.get("gapMode"):
        args.append("--gap-mode")
    # 散材料模式：源报告是「按模板章节自动汇编」的散材料（设计成果/清单/图纸），
    # 不是成篇现状分析。此时 06 不做摘句搬运，改为按事实库成文。
    if state.get("scatterMode"):
        args.append("--scatter-mode")
    if filled_dir.is_dir():
        args += ["--filled-dir", filled_dir]
    code, _ = call(args)
    if code != 0:
        raise Gate("分章成文失败")

    # 回填稿**格式预检**：在单章自检之前把「缺项内联标记写成『见缺项-D01』」
    # 这类问题挑出来。09 步的契约检查也会判「附件四有行但正文未引用」，
    # 但那已经到最后一步，报错指向附件而不是根因。
    # 预检只针对「回填后的成稿」：无人回填时 filled_dir 里放的是模板本身（自带禁句与占位来源行），
    # 降级稿由程序生成、生成即干净，不该拿模板去撞格式闸门。
    if not need_write and filled_dir.is_dir() and any(filled_dir.glob("*.md")):
        code, _ = call([script("06-分章成文", "precheck_drafts.py"),
                        "--draft-dir", filled_dir,
                        "--out", D06 / "自检" / "回填稿预检.json"])
        # exit 1 才是「格式不合格」；其余非 0 是脚本自身出错（崩溃/环境），
        # 必须按执行失败如实上报，不能翻译成业务判定。
        if code == 1:
            raise Gate("回填稿格式预检未通过：见 过程文件/06-分章成文/自检/回填稿预检.json"
                       "（缺项内联标记须写成【缺项-XNN：见本节缺项表】，方括号内、冒号在括号内）")
        if code != 0:
            raise Gate(f"回填稿格式预检脚本执行失败（exit {code}）：这是程序/环境错误，"
                       f"不是格式判定；详见上方输出与 回填稿预检.json")

    results, blocking = {}, []
    # 缺口版：章节本身即因无源可引而留空，长度指标不构成结构性失败；
    # 标题/过程段/缺项/禁句等结构性检查仍然拦截。
    g = gates()
    blocking_checks = set(g.get("blockingChecks") or BLOCKING_CHECKS)
    if state.get("gapMode") or need_write:
        # 降级稿（无人回填、程序确定性成文）篇幅天然短于成文稿：篇幅类指标降为提示，
        # 结构性检查（标题/禁句/缺项标记）照旧拦截。
        blocking_checks -= set(g.get("gapModeRelaxed") or ["CHK-LEN-1", "CHK-LEN-2"])
    check_dir = D06 / "自检"
    if check_dir.is_file():  # 兼容旧版把 --out 当文件写的产物
        check_dir.unlink()
    check_dir.mkdir(parents=True, exist_ok=True)
    for L in "ABCDEF":
        f = next(Path(D06 / "分章初稿").glob(f"{L}_*.md"), None)
        if not f:
            blocking.append(f"{L} 章缺失")
            continue
        jf = D06 / "自检" / f"{L}.json"
        jf.parent.mkdir(parents=True, exist_ok=True)
        # 六章共用 06 的统一自检引擎，--unit 选本章判据（本章契约见各章 规则/章节契约.yaml）
        code, _ = call([script("06-分章成文", "check_chapter.py"), "--unit", L,
                        "--project-dir", WS, "--draft", f,
                        "--direction", p["direction"], "--out", jf], allow=(0, 1))
        # 0/1 才是自检结论；其余退出码是脚本自身出错，不能当成「本章通过」
        if code not in (0, 1):
            raise Gate(f"单章自检执行失败（{L}，exit {code}）：程序/环境错误，"
                       f"不是本章不合格；详见上方输出")
        if not jf.is_file():
            raise Gate(f"单章自检未产出结果（{L}）：{rel(jf)} 不存在，不能按通过处理")
        res = load_json(jf, {})
        fails = [c for c in res.get("checks", [])
                 if c.get("level") == "FAIL" and c.get("id") in blocking_checks]
        warns = [c for c in res.get("checks", []) if c.get("level") in ("FAIL", "WARN")]
        results[L] = {"exit": code, "fails": len(fails), "warns": len(warns)}
        if fails:
            blocking.append(f"{L}: " + "；".join(f"{c['id']} {c.get('message','')}" for c in fails))
    if blocking:
        raise Gate("分章自检存在结构性失败：" + " | ".join(blocking))
    note = ", ".join(f"{k}(FAIL {v['fails']}/WARN {v['warns']})" for k, v in results.items())
    if state.get("gapMode"):
        note += "；缺口版：CHK-LEN-1/2 降为提示级"
    save_step(state, "06", "分章成文", 0,
              notes="六章结构通过；其余 FAIL/WARN 按提示级留痕：" + note,
              outputs=[rel(D06 / "分章初稿")])


def step07(state):
    print("[07] 表 3.1 口径计算")
    p = params()
    call([script("07-表3.1口径计算", "build_calc_input.py"),
          "--facts", D05 / "facts" / "事实库.json", "--out", D07 / "计算输入.json"])
    # 清掉上一次的指标计算产物，避免同日重复运行累积 -v2/-v3 之类版本名
    for stale in list(D07.glob("指标计算-*")):
        stale.unlink()
    code, _ = call([script("07-表3.1口径计算", "calc_metrics.py"),
                    "--project-dir", WS, "--input", D07 / "计算输入.json",
                    "--out", D07, "--format", "text"], allow=(0, 1, 3))
    if code == 3:
        raise Gate("指标计算命中硬门禁（除零/单位不一致/占比超差），已停止")
    calc = sorted(Path(D07).glob("指标计算-*.json"), key=os.path.getmtime)
    if not calc:
        raise Gate("指标计算无输出")
    trace = load_json(calc[-1])
    call([script("07-表3.1口径计算", "build_table31.py"), "--trace", calc[-1],
          "--out", D07 / "表3.1.md", "--title", "表 3.1 现状用地分类面积与占比"])
    save_step(state, "07", "表3.1口径计算", 0,
              notes=("全缺项留空，未补数" if code == 1 else "计算通过"),
              outputs=[rel(D07 / "计算输入.json"), rel(calc[-1]), rel(D07 / "表3.1.md")])


def step08(state):
    _refresh_edition(state)
    print("[08] 组装与模板渲染")
    p = params()
    md = D08 / "正文.md"
    gap_md = D08 / "数据缺口汇总.md"
    # 出图：只从 06 导出的聚合值画（用地构成、高度分位、台账层数），PNG 供 docx 与 PDF 嵌入
    fig_data = D06 / "分章初稿" / "图表数据.json"
    if fig_data.exists():
        code_fig, _ = call([script("08-组装与模板渲染", "make_figures.py"),
                            "--data", fig_data, "--out-dir", D08 / "图表"])
        if code_fig != 0:
            raise Gate("图表生成失败")
    else:
        print("  [图表] 未找到 图表数据.json，跳过出图")
    # 现状图：正射影像＋地类＋建筑轮廓＋模型覆盖范围（B4）
    # 缺 OSGB 影像/边界即硬失败，不生成示意图顶替
    site_map = D08 / "图表" / "fig_site_map.png"
    code_map, _ = call([script("08-组装与模板渲染", "make_site_map.py"),
                        "--workspace", WS,
                        "--baseline", D05 / "包内依据" / "包内依据.json",
                        "--out", site_map])
    if code_map != 0 or not site_map.is_file():
        raise Gate("现状图生成失败：缺 OSGB 正射影像/地类/建筑轮廓，"
                   "或出图脚本执行失败；按“不静默降级”口径停止，不用示意图冒充")
    cmd = [script("08-组装与模板渲染", "assemble_chapters.py"),
           "--extract-dir", D06 / "分章初稿", "--out", md,
           "--gap-json", D06 / "分章初稿" / "_缺口登记.json",
           "--mapping-out", D08 / "来源映射.md"]
    if state.get("gapMode"):
        # 缺口版：正文只留「缺口索引」一页（指向 数据缺口清单.xlsx），不造完整附件四
        cmd += ["--gap-heading", "附件四 数据缺口索引", "--gap-index-only",
                "--append-to-chapter", "F"]
    else:
        # 干净版：缺口清单单独落文件，正文剥离内联缺项标记
        cmd += ["--gap-out", str(gap_md), "--strip-inline-gaps"]
    code, _ = call(cmd)
    if code != 0:
        raise Gate("章节组装失败")
    text = md.read_text(encoding="utf-8")
    inline_gaps = set(re.findall(r"【(缺项-[A-F]\d{1,3})：", text))
    if state.get("gapMode"):
        lines = text.splitlines()
        idx = next((i for i, l in enumerate(lines) if l.startswith("# ")), 0)
        note = ("缺口版说明：本版按项目现有材料逐条留痕，缺口不影响既有结论的定性成立；"
                "缺口明细见同目录《数据缺口清单.xlsx》，补齐资料后重跑即可生成干净版。")
        lines.insert(idx + 1, "")
        lines.insert(idx + 2, note)
        text = "\n".join(lines) + "\n"
        # 缺口版至少保留一页「缺口索引」（指向 xlsx）：正文没有内联缺项标记时，
        # 08 不会生成附件四，这里补一页索引，保证缺口版与干净版有可见差别。
        if "## 附件四 数据缺口索引" not in text:
            text = text.rstrip("\n") + "\n\n## 附件四 数据缺口索引\n\n" + \
                ("本版为缺口留痕版：正文各节均按现有材料给出初步判定（推断依据与置信度随文标注）；"
                 "本次登记的数据缺口（材料缺口 / 章节口径缺口 / 事实缺失）见与本报告同目录的"
                 "《数据缺口清单.xlsx》，按「编号 / 缺什么 / 向谁要 / 补齐后强化哪一章」逐条列示。"
                 "补齐后重跑即可收敛为干净版。") + \
                "\n\n（数据来源：项目资料、结合项目资料进行整理）\n"
        md.write_text(text, encoding="utf-8")
    else:
        # 干净版：正文不得出现缺口提示；缺失的附件只记录到过程文件，不进正文
        _att_missing = [t for t in ("附件一 现有基础资料清单", "附件二 图件与影像资料",
                                    "附件三 专项台账")
                        if f"## {t}" not in text]
        if _att_missing:
            print(f"  [附件] 未成篇：{'、'.join(_att_missing)}（已记入过程文件，正文不提示）")
            dump_json(D08 / "附件齐备情况.json",
                      {"missing": _att_missing, "note": "干净版不把缺图/缺台账写进正文"})

    # 成果文件名**不带「（缺口版）」**（用户要的是《现状分析报告》）；
    # 缺口版作为**元数据 + 正文开头的缺口说明**体现，正文开头已插入缺口说明段。
    state["edition"] = "gap" if state.get("gapMode") else "clean"
    title = f"{p['project']}现状分析报告"
    docx = prod_dir() / f"{title}.docx"
    _now = datetime.now()
    cover_note = f"编制时间：{_now.year} 年 {_now.month} 月"
    code, out = call([script("08-组装与模板渲染", "build_report_docx.py"),
                      "--template", tmpl(p), "--content", md, "--out", docx,
                      "--title", title, "--cover-note", cover_note, "--force"], capture=True)
    if code != 0 or "[verify] ok" not in out:
        raise Gate("模板渲染/回读校验失败")
    verify_lines = [ln.strip() for ln in out.splitlines() if ln.strip().startswith("[verify]")]
    headings = next((ln for ln in verify_lines if "headings:" in ln), "")
    tables = next((ln for ln in verify_lines if "tables:" in ln and "shapes" not in ln), "")
    readback = {"schema": "cardo/docx-readback/v1",
                "verified": True, "title": title,
                "headings": headings.split("headings:", 1)[1].strip() if headings else "",
                "tables": tables.split("tables:", 1)[1].strip() if tables else "",
                "lines": verify_lines}
    dump_json(D08 / "回读报告.json", readback)
    normalize_zip(docx)  # 同一正文 → 同一 sha256，便于基线 diff

    # 三件套：标注源 → 干净版 / 标注版 / 依据对照表（同源生成，唯一实现在 09）
    # 校验失败不在 08 拦（09 步统一拦），这里只负责导出并把问题留痕。
    code, out = call([script("08-组装与模板渲染", "export_evidence_pack.py"),
                      "--project-dir", WS, "--out-dir", internal_dir("证据包")],
                     allow=(0, 1, 2), capture=True)
    evidence_note = ""
    if code == 2:
        evidence_note = "；未发现标注源，三件套未导出（旧格式章节）"
        print("  提示：未发现含〔锚点〕的标注源，三件套未导出（旧格式章节不纳入）")
    elif code == 1:
        evidence_note = "；三件套已导出但依据校验不通过（09 拦）"
        print("  提示：三件套已导出，依据校验不通过 —— 见 过程文件/09-三重自检与边界/依据校验_*.md")
    else:
        evidence_note = "；三件套已导出"

    # 正文写着「源文件与 sha256 登记在内部件《同源核对.xlsx》」，该件必须真在内部件里；
    # 成果目录只收 Office 文件，所以把 json 转成同名 xlsx 放证据包，不塞 .json。
    src_check = D05 / "包内依据" / "同源核对.json"
    if src_check.is_file():
        from openpyxl import Workbook
        chk = json.loads(src_check.read_text(encoding="utf-8"))
        wb = Workbook()
        ws = wb.active
        ws.title = "同源核对"
        ws.append(["类别", "源文件", "sha256", "schema"])
        for s in chk.get("sources") or []:
            ws.append([s.get("kind", ""), s.get("file", ""), s.get("sha256", ""), s.get("schema", "")])
        ws.append([])
        ws.append(["核对项", "结果", "", ""])
        ws.append(["OSGB 指标条数", chk.get("osgbMetricCount"), "", ""])
        ws.append(["AHP 权重合计", chk.get("ahpWeightsSum"), "合计为 1" if chk.get("ahpWeightsSumOk") else "合计不为 1", ""])
        ws.append(["AHP 排序单调", "是" if chk.get("ahpRankingDesc") else "否", "", ""])
        pack = internal_dir("证据包")
        pack.mkdir(parents=True, exist_ok=True)
        wb.save(pack / "同源核对.xlsx")

    # 成品 PDF：docx → A4 HTML → 无头 Edge 打印 → 目录页码回填 → 页脚烘入
    pdf = docx.with_suffix(".pdf")
    toc_json = D08 / "目录.json"
    code, _ = call([script("08-组装与模板渲染", "pdf_render.py"),
                    "--docx", docx, "--out", pdf, "--title", title,
                    "--toc-json", str(toc_json)])
    if code != 0 or not pdf.exists():
        raise Gate("成品 PDF 生成失败")

    # 公文版式：sectPr / 仿宋 16pt / 29pt 固定行距 / 首行缩进 / 页脚双域 / 目录域+缓存
    laid = docx.with_name(docx.stem + ".layout.docx")
    code, log = call([script("08-组装与模板渲染", "gongwen_layout.py"),
                      str(docx), str(laid), "--toc", str(toc_json)], capture=True)
    if code != 0 or not laid.exists():
        raise Gate(f"公文版式套用失败：{log.strip()[-200:]}")
    os.replace(laid, docx)
    normalize_zip(docx)

    # 成品闸门：缺口词零命中、政策文号有据、OSGB/AHP 同源
    code, gate_out = call([script("09-三重自检与边界", "final_gates.py"),
                           "--project-dir", WS, "--docx", docx, "--pdf", pdf,
                           "--baseline", D05 / "包内依据" / "包内依据.json",
                           "--basis-lib", S.parent / "assets" / "依据库" / "依据库.json",
                           "--out", D09 / "成品闸门.json"], capture=True)
    if code != 0:
        raise Gate(f"成品闸门未通过：{gate_out.strip().splitlines()[-1] if gate_out.strip() else '见 过程文件/09-三重自检与边界/成品闸门.json'}")

    save_step(state, "08", "组装与模板渲染", 0,
              notes="[verify] ok；标题/章节/表格/样式/东亚字体全部通过；成品 PDF 已生成；成品闸门通过"
                    + evidence_note,
              outputs=[rel(md), rel(docx), rel(pdf), rel(D08 / "回读报告.json"),
                       rel(D09 / "成品闸门.json")])


def step09(state):
    print("[09] 三重自检与边界")
    p = params()
    degraded = bool(state.get("degraded"))
    md = D08 / "正文.md"
    docx = next((OUT / "成品版").rglob("*现状分析报告*.docx"))
    code, _ = call([script("09-三重自检与边界", "build_trace.py"),
                    "--project-dir", WS, "--artifact", md,
                    "--facts", D05 / "facts" / "事实库.json",
                    "--out", D09 / "回溯" / "回溯表.json", "--baseline-version", "1"])
    if code != 0:
        raise Gate("回溯表为空（正文未引用事实）")
    code, _ = call([script("09-三重自检与边界", "check_trace.py"),
                    "--project-dir", WS, "--artifact", D09 / "回溯" / "回溯表.json",
                    "--facts", D05 / "facts" / "事实库.json",
                    "--out", D09 / "回溯检查"])
    if code != 0:
        raise Gate("回溯校验未通过")

    gap_src = (D08 / "正文.md") if state.get("gapMode") else (D08 / "数据缺口汇总.md")
    if not gap_src.exists():
        gap_src = D06 / "分章初稿"
    call([script("02-齐套检查与缺料拍板", "export_gap_xlsx.py"),
          "--materials", D02 / "齐套检查" / "材料齐套表.json",
          # 缺口版：成稿正文里的附件四；干净版：08 单独落盘的 数据缺口汇总.md
          "--chapters", gap_src,
          "--facts", D05 / "facts" / "事实库.json",
          # 01 低置信度/不可读、06 叙述推导缺口一并进缺口清单
          "--data-gap", D01 / "数据缺口登记.json",
          "--data-gap", WS / "记忆与规则" / "叙述缺口.json",
          # 投递区路径：只用来取末级目录名（如 输入文件），落位目录写项目内相对路径
          "--delivery-root", str(WS / (p.get("delivery") or "输入文件")),
          "--out", internal_dir("数据缺口清单") / "数据缺口清单.xlsx"])

    code, out = call([script("09-三重自检与边界", "check_contracts.py"),
                      "--project-dir", WS, "--direction", p["direction"],
                      "--template", tmpl(p), "--report", md, "--docx", docx,
                      "--contracts", S / "09-三重自检与边界/规则",
                      "--source-dump", D03 / "源报告转储",
                      # 缺口版只放宽篇幅类要求；内容与追溯类检查不因缺料放行
                      *(["--gap-mode"] if (state.get("gapMode") or degraded) else ["--edition", "clean"]),
                      "--out", D09 / "契约检查"], allow=(0, 1, 3), capture=True)
    contract_fails = [ln.split("FAIL ", 1)[1].strip() for ln in out.splitlines()
                      if "[contracts] FAIL " in ln]
    if code != 0:
        # 自检说「不通过」就必须让整链失败：业务不合格 1、设计拒收 4、脚本异常 3。
        ccode = 4 if code == 4 else (3 if code == 3 else 1)
        _write_verify_note(
            "契约检查", [(f.split("：", 1)[0], f.split("：", 1)[1] if "：" in f else f)
                        for f in contract_fails],
            [f for f in contract_fails], ccode,
            extra="详见 过程文件/09-三重自检与边界/契约检查/",
            locator_root=D06 / "分章初稿")
        raise Gate(f"契约检查存在硬失败（fail={len(contract_fails)}），"
                   f"已写《校验未通过说明》；详见 过程文件/09-三重自检与边界/契约检查",
                   ccode)
    contract_line = next((l for l in out.splitlines() if "[contracts] exit=" in l), "")

    # 依据三件套：只校验不写产物（导出在 08），口径与 08 完全同源。
    # 第 3 轮起：数值有源 / 正文洁净 / 依据有效 / 对照表回指 / 链条闭合 一律硬失败，
    # 降级稿只放宽篇幅，不放宽这几类。exit 4=无标注源（设计拒收）、1=业务不合格、3=脚本异常。
    code, ev_out = call([script("09-三重自检与边界", "check_evidence.py"),
                         "--project-dir", WS], allow=(0, 1, 3, 4), capture=True)
    if code != 0:
        ev_fails, ev_problems = _evidence_failures(D09)
        if not ev_fails and not ev_problems:
            ev_problems = [ln.strip() for ln in ev_out.splitlines()
                           if ln.strip().startswith("FAIL ")]
        ccode = 4 if code == 4 else (3 if code == 3 else 1)
        _write_verify_note("依据三件套", ev_fails, ev_problems, ccode,
                           extra="详见 过程文件/09-三重自检与边界/依据校验_*.md",
                           locator_root=D06 / "分章初稿")
        if ccode == 4:
            raise Gate("依据三件套拒收：06 分章初稿没有〔锚点〕依据块，"
                       "追溯链无法建立；已写《校验未通过说明》", 4)
        if ccode == 3:
            raise Gate("依据校验脚本自身异常，已写《校验未通过说明》", 3)
        raise Gate("依据三件套校验未通过：见 过程文件/09-三重自检与边界/依据校验_*.md"
                   "（已写《校验未通过说明》）", 1)

    # 形式参考件串项守卫：拿参考件的专名与数值扫成品正文，命中即失败
    fr_json = D00 / "形式参考" / "形式参考.json"
    code, _ = call([script("09-三重自检与边界", "check_form_reference.py"),
                    "--project-dir", WS, "--artifact", md,
                    "--register", fr_json,
                    "--out", D09 / "形式参考检查"], allow=(0, 1, 2))
    if code == 1:
        raise Gate("形式参考件串入：成品正文命中参考件的专名或数值，"
                   "详见 过程文件/09-三重自检与边界/形式参考检查/")
    if code == 2:
        raise Gate("形式参考件串项检查输入缺失（正文/登记件）")

    code, _ = call([script("09-三重自检与边界", "boundary_check.py"),
                    "--project-dir", WS,
                    "--fingerprint", D01 / "原件指纹.json",
                    "--template", tmpl(p),
                    "--contract", S / "09-三重自检与边界/规则/章节契约.yaml",
                    "--inputs", WS / p["delivery"],
                    "--outputs", OUT,
                    "--deliverables", docx,
                    internal_dir("输入材料补齐确认单") / "输入材料补齐确认单.docx",
                    internal_dir("数据缺口清单") / "数据缺口清单.xlsx",
                    "--report", md,
                    *(["--gap-mode"] if state.get("gapMode") else []),
                    "--out", D09 / "边界检查.json"])
    if code != 0:
        raise Gate("边界检查未通过（原件被改/绝对路径/交付物缺失）")

    trace = load_json(D09 / "回溯" / "回溯表.json", {})
    facts = load_json(D05 / "facts" / "事实库.json", {})
    stats = load_json(D04 / "映射统计.json", {})

    # 《输入材料补齐确认单》重新生成一次：它在 02 步生成（那时章节缺项还没定稿），
    # 而报告在 09 才定稿——两份文件的缺口口径可能脱节。
    # 这里按**最终**的缺口清单重生成，保证「报告里标的缺项」与「单子里列的缺口」一致。
    _regen = call([script("02-齐套检查与缺料拍板", "make_material_request.py"),
                   "--project-dir", WS,
                   "--gap-json", D02 / "齐套检查" / "材料齐套表.json",
                   "--classification", D01 / "分类清单" / "分类清单.json",
                   "--project-id", D00 / "项目识别.json",
                   "--out", internal_dir("输入材料补齐确认单") / "输入材料补齐确认单.docx",
                   "--md", D02 / "补齐确认单.md",
                   "--report-name", f"{p['project']}现状分析报告",
                   "--report-ready"])
    if _regen[0] != 0:
        say("  提示：补齐确认单重新生成未成功（exit {}），沿用 02 步的版本".format(_regen[0]))

    source_path = src(p)
    chapters = {f.name.split("_")[0]: len(f.read_text(encoding="utf-8"))
                for f in Path(D06 / "分章初稿").glob("*.md") if not f.name.startswith("_")}
    contract = {"fail": 0, "hint": None, "manual": None}
    m = __import__("re").search(r"fail=(\d+)\s+hint=(\d+)\s+skip=(\d+)\s+manual=(\d+)", contract_line)
    if m:
        contract = {"fail": int(m.group(1)), "hint": int(m.group(2)), "manual": int(m.group(4))}
    dump_json(D09 / "契约检查" / "契约检查.json",
              {"schema": "cardo/contract-check/v1", "exit": code, **contract,
               "fails": contract_fails, "summary": contract_line.strip()})
    trace_checks = sorted((D09 / "回溯检查").glob("回溯检查-*.json"))
    trace_check = load_json(trace_checks[-1], {}) if trace_checks else {}
    if not trace_check or trace_check.get("summary", {}).get("exit") != 0:
        raise Gate("回溯检查未通过")
    art0 = (trace_check.get("artifacts") or [{}])[0]
    coverage = art0.get("coverage")
    min_cov = float(gates().get("traceCoverageMin", 1.0))
    if coverage is None or float(coverage) < min_cov - 1e-4:
        raise Gate(f"回溯覆盖率不足 {min_cov * 100:.0f}%（实际 {coverage}）")
    metrics = {
        "schema": "cardo/run-metrics/v1",
        "generatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "durationSeconds": round(time.time() - RUN_T0, 1) if RUN_T0 else None,
        "templateSha256": sha256(tmpl(p)),
        "sourceSha256": (sha256(source_path)
                          if source_path is not None and Path(source_path).is_file() else None),
        "nodeCount": load_json(D03 / "章节骨架.json", {}).get("nodeCount"),
        "mapping": {"high": stats.get("exact", 0), "medium": stats.get("alias", 0),
                    "gap": stats.get("gap", 0)},
        "facts": len(facts.get("facts", [])),
        "chapters": chapters,
        "contract": contract,
        "contractFails": contract_fails,
        "trace": {"items": len(trace.get("items", [])), "coverage": coverage,
                  "exit": trace_check.get("summary", {}).get("exit")},
        "reportSha256": sha256(docx),
        "gapMode": bool(state.get("gapMode")),
    }
    dump_json(D09 / "运行指标.json", metrics)
    call([script("09-三重自检与边界", "baseline_check.py"),
          "--baseline", WS / "记忆与规则/验证基线-CardoStatus.json", "--metrics", D09 / "运行指标.json"])

    write_verify_report(metrics)
    save_step(state, "09", "三重自检与边界", 0,
              notes=contract_line.strip(), outputs=[rel(D09 / "验证报告.md")])


def write_verify_report(mt):
    warn_note = ("章内自检的 FAIL/WARN 逐条留痕于 06 分章自检报告；其中 B 章证据数、"
                 "D/E 章 S 编号检查面向回填成稿结构，散材料降级稿不适用（本版按章节标题引用、"
                 "全文无 S-NN 编号）。硬门禁以 09 契约检查 fail=0 与依据三件套全通过为准。")
    gap_note = ("缺口版：契约未达标项均为原始资料缺失导致，已逐项登记于 数据缺口清单.xlsx 与"
                " 契约检查 报告；不作推断、不补数字。补齐资料后重跑本命令，契约检查须回到 fail=0。")
    src_fp = mt.get("sourceSha256")
    src_fp_text = (f"`{src_fp[:16]}…`（原件未改）" if src_fp
                   else "无（散材料自动汇编，源报告由 03 步骤生成）")
    lines = [
        "# 验证报告（现状分析流水线）", "",
        f"- 生成时间：{mt['generatedAt']}",
        f"- 模板指纹：`{mt['templateSha256'][:16]}…`（与章节契约一致）",
        f"- 源报告指纹：{src_fp_text}",
        f"- 模板节点：{mt['nodeCount']}（H1=6 / H2=17）",
        f"- 章节映射：high={mt['mapping']['high']} / medium={mt['mapping']['medium']} / "
        f"gap={mt['mapping']['gap']}",
        f"- 事实库：{mt['facts']} 条",
        f"- 回溯表：{mt['trace']['items']} 项"
        f"（覆盖率 {mt['trace']['coverage'] * 100:.2f}%，exit={mt['trace']['exit']}）",
        f"- 契约检查：fail={mt['contract']['fail']} / hint={mt['contract']['hint']} / "
        f"manual={mt['contract']['manual']}",
        f"- 成稿指纹：`{mt['reportSha256'][:16]}…`",
        f"- 缺口版：{'是' if mt['gapMode'] else '否'}", "",
        "## 分章字数", "", "| 章 | 字数 |", "| --- | --- |",
    ]
    for k in "ABCDEF":
        if k in mt["chapters"]:
            lines.append(f"| {k} | {mt['chapters'][k]} |")
    if mt.get("contractFails"):
        lines += ["", "## 契约未达标项（缺口版）", ""]
        lines += [f"- {x}" for x in mt["contractFails"]]
    # 自检不通过就是未通过：缺口版只放宽篇幅，不放宽契约结论。
    passed = mt["contract"]["fail"] == 0
    if mt["gapMode"]:
        verdict = (f"缺口版：契约 fail={mt['contract']['fail']}，"
                   f"{'通过' if passed else '未通过'}；补齐资料后重跑须回到 fail=0。")
    else:
        verdict = (f"干净版：契约 fail={mt['contract']['fail']}，"
                   f"{'通过' if passed else '未通过'}。")
    lines += ["", "## 判定", "", f"- 09 独立验证：{verdict}",
              f"- 边界：原件只读、模板未替换、技能无绝对路径、交付物齐套（见 边界检查.json）。",
              f"- 说明：{warn_note}"]
    if mt["gapMode"]:
        lines.append(f"- 说明：{gap_note}")
    lines.append("")
    (D09 / "验证报告.md").write_text("\n".join(lines), encoding="utf-8")


STEPS = [("01", step01), ("02", step02), ("03", step03), ("04", step04), ("05", step05),
         ("06", step06), ("07", step07), ("08", step08), ("09", step09)]
STEP_NAMES = {"01": "建档与材料归集", "02": "齐套检查与缺料拍板", "03": "模板与源报告转储",
              "04": "章节映射", "05": "事实抽取与事实库", "06": "分章成文",
              "07": "表3.1口径计算", "08": "组装与模板渲染", "09": "三重自检与边界"}


def clean_legacy_outputs():
    """证据包与成品只留本轮版本：清旧平面文件与带时间戳的历史证据包。"""
    for pat in ("*现状分析报告*.docx", "*现状分析报告*.pdf",
                "输入材料补齐确认单.docx", "数据缺口清单.xlsx"):
        for f in OUT.glob(pat):
            if f.is_file():
                f.unlink()
    for d in OUT.glob("证据包*"):
        if d.is_dir():
            shutil.rmtree(d)
    ev = OUT / "内部件" / "证据包"
    if ev.is_dir():
        shutil.rmtree(ev)
    # 上一轮失败时写的说明文件不属于交付物；本轮通过后不能残留
    stale = OUT / "校验未通过说明.md"
    if stale.is_file():
        stale.unlink()


def cmd_run(args):
    global RUN_T0, STEP_T0, MATERIAL_CHOICE, STRICT_REVIEW, EDITION
    MATERIAL_CHOICE = getattr(args, "material_choice", "先出稿")
    STRICT_REVIEW = bool(getattr(args, "strict_review", False))
    EDITION = getattr(args, "edition", "auto")
    params()
    for d in (D00, D01, D02, D03, D04, D05, D06, D07, D08, D09, OUT):
        d.mkdir(parents=True, exist_ok=True)
    clean_legacy_outputs()
    state = load_state()
    state["startedAt"] = datetime.now().astimezone().isoformat(timespec="seconds")
    state.pop("finishedAt", None)
    RUN_T0 = time.time()
    start = 0
    if args.start_from:
        start = next(i for i, (k, _) in enumerate(STEPS) if k == args.start_from)
    stop = len(STEPS)
    if args.stop_at:
        stop = next(i + 1 for i, (k, _) in enumerate(STEPS) if k == args.stop_at)
    if stop <= start:
        raise Gate("--to 必须晚于 --from", 2)
    print(f"CardoStatus run：{params()['project']}（{params()['direction']}）")
    for key, fn in STEPS[start:stop]:
        STEP_T0 = time.time()
        try:
            fn(state)
            print(f"  → {key} PASS ({time.time() - STEP_T0:.1f}s)\n")
        except Gate as e:
            save_step(state, key, STEP_NAMES.get(key, key), e.code, notes=str(e))
            print(f"  → {key} {'需人工确认' if e.code == 3 else 'FAIL'}：{e}\n")
            dump_json(D00 / "运行状态.json", state)
            return e.code
    state["finishedAt"] = datetime.now().astimezone().isoformat(timespec="seconds")
    state["durationSeconds"] = round(time.time() - RUN_T0, 1)
    state["variant"] = "缺口版" if state.get("gapMode") else "干净版"
    dump_json(D00 / "运行状态.json", state)
    write_run_report(state)
    if start == 0 and stop == len(STEPS):
        print(f"全部九步通过：exit 0（{state['variant']}，{state['durationSeconds']}s）")
    else:
        print(f"已执行 {STEPS[start][0]}–{STEPS[stop - 1][0]}：exit 0"
              f"（{state['variant']}，{state['durationSeconds']}s）")
    print(f"交付物目录：{rel(OUT)}")
    return 0


def cmd_status(args):
    state = load_state()
    print(f"工作区：{WS}")
    print(f"项目：{params().get('project')} / {params().get('direction')}")
    print(f"状态：{state.get('variant', '未开始')}")
    for key, _ in STEPS:
        st = state.get("steps", {}).get(key)
        if not st:
            print(f"  {key} 未运行")
        else:
            print(f"  {key} {st['status']:<9} {st.get('notes','')[:80]}")
    for name, sub in (("现状分析报告", "成品版/现状分析报告"),
                      ("输入材料补齐确认单", "内部件/输入材料补齐确认单"),
                      ("数据缺口清单", "内部件/数据缺口清单")):
        files = list((OUT / sub).glob(f"*{name}*"))
        print(f"  {name}: {rel(files[0]) if files else '（未生成）'}")
    return 0


def cmd_scaffold(args):
    dest = Path(args.dest).resolve()
    if dest.exists() and any(dest.iterdir()):
        raise SystemExit(f"目标目录非空：{dest}")
    dest.mkdir(parents=True, exist_ok=True)
    shutil.copytree(S, dest / "Skills")
    # 记忆与规则：只带规则说明，不带本项目的验证基线（基线须由新工作区首次运行自行生成）
    (dest / "记忆与规则").mkdir(parents=True, exist_ok=True)
    for f in (WS / "记忆与规则").glob("*"):
        if f.is_file() and not f.name.startswith("验证基线"):
            shutil.copy2(f, dest / "记忆与规则" / f.name)
    if (WS / "README.md").exists():
        shutil.copy2(WS / "README.md", dest / "README.md")
    for d in ["过程文件/00-总控", "过程文件/01-建档与材料归集", "过程文件/02-齐套检查与缺料拍板",
              "过程文件/03-模板与源报告转储", "过程文件/04-章节映射", "过程文件/05-事实抽取与事实库",
              "过程文件/06-分章成文", "过程文件/07-表3.1口径计算", "过程文件/08-组装与模板渲染",
              "过程文件/09-三重自检与边界", "输入文件", "成果文件"]:
        (dest / d).mkdir(parents=True, exist_ok=True)
    tpl_src, src_src = Path(args.template).resolve(), Path(args.source).resolve()
    # 模板属于新工作区自己的 提资数据/ 标准资产：不往工作区外面放东西
    templates = dest / "提资数据" / "模板"
    templates.mkdir(parents=True, exist_ok=True)
    target_tpl = templates / tpl_src.name
    if not target_tpl.exists():
        shutil.copy2(tpl_src, target_tpl)
    shutil.copy2(src_src, dest / "输入文件" / src_src.name)
    if args.materials:
        md = Path(args.materials).resolve()
        for f in md.rglob("*"):
            if f.is_file():
                shutil.copy2(f, dest / "输入文件" / f.name)
    p = {"schema": "cardo/workspace-params/v1", "project": args.project,
         "direction": args.direction, "delivery": "输入文件",
         "template_file": os.path.relpath(target_tpl, dest).replace("\\", "/"),
         "source_report": src_src.name,
         "gap_mode": args.gap_mode,
         "notes": "scaffold 生成；模板在根目录 提资数据/（相对路径），source_report 为人工指定的只读原件。"}
    dump_json(dest / "工作区参数.json", p)
    print(f"已创建新工作区：{dest}")
    print(f"  下一步：cd 到该目录后运行 python <技能目录>/Skills/00-Cardo现状分析总控/scripts/pipeline.py run")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="CardoStatus 现状分析九步流水线")
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run", help="依次执行 01–09")
    r.add_argument("--workspace", default="", help="工作区根（缺省为本脚本所在工作区）")
    r.add_argument("--from", dest="start_from", default="")
    r.add_argument("--to", dest="stop_at", default="")
    r.add_argument("--material-choice", choices=["先出稿", "先补料"], default="先出稿",
                   help="02 缺料拍板：先出稿=缺口版继续；先补料=exit 3 停下等补齐")
    r.add_argument("--strict-review", action="store_true",
                   help="04 步中等置信映射改为停下来等人工复核（默认自动采纳，不打断）")
    r.add_argument("--edition", choices=["clean", "gap", "auto", "final"], default="clean",
                   help="成果版本：clean=干净版（缺省，正文无缺口提示、不附附件四）；"
                        "gap=缺项留痕版（可选产物）；auto/final 为兼容别名，等同 clean。"
                        "（文件名统一为《XX现状分析报告》，版本体现在元数据里）")
    s = sub.add_parser("status", help="查看运行状态")
    s.add_argument("--workspace", default="", help="工作区根（缺省为本脚本所在工作区）")
    sc = sub.add_parser("scaffold", help="复制技能与目录，建立新工作区")
    sc.add_argument("--dest", required=True)
    sc.add_argument("--project", required=True)
    sc.add_argument("--direction", default="老旧小区")
    sc.add_argument("--template", required=True)
    sc.add_argument("--source", required=True)
    sc.add_argument("--materials", default="")
    sc.add_argument("--gap-mode", default="auto", choices=["auto", "on", "off"])
    args = ap.parse_args(argv)
    try:
        if getattr(args, "workspace", ""):
            ws = Path(args.workspace).resolve()
            if not SKILLS.is_dir():
                print(f"[参数错误] 找不到技能链（CARDO_SKILLS_ROOT）：{SKILLS}")
                return 2
            bind_workspace(ws)
        if args.cmd == "run":
            return cmd_run(args)
        if args.cmd == "status":
            return cmd_status(args)
        return cmd_scaffold(args)
    except ParamsError as e:
        print(str(e))
        return 2
    except Gate as e:
        print(f"[门禁] {e}")
        return e.code
    except StopIteration:
        print("[参数错误] --from/--to 必须是 01–09")
        return 2


if __name__ == "__main__":
    sys.exit(main())
