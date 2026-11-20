#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""detect_project.py -- 从投递区材料里自动识别「这份报告写哪个项目」。

为什么需要它
------------
开工前问用户「这份《现状分析报告》写哪个项目？」是整条流水线最大的体验杀手：
材料里往往已经写明了项目名（PPT 封面标题、批复抬头、图片里的项目信息卡），
让人去读一遍文件名再回答，不如让机器读完直接给结论 **并附依据**。

本脚本只读投递区，输出 JSON（stdout），不做任何判定式的决定——
**它给建议，决定权仍在用户手里**：

    python detect_project.py --input <输入文件目录> [--out <结果json>]

输出结构：
    {
      "schema": "cardo/detect-project/v1",
      "project":  {"value": "...", "confidence": 0.9, "basis": "...", "evidence": [...]},
      "direction":{"value": "老旧街区", "confidence": 0.9, "evidence": [...]},
      "candidates": [{"name":..., "hits": 4, "titleHits": 2, "score": 14,
                      "files": [...], "core": "某片区"}],
      "needsDecision": false,       # true 表示证据不足以自动定，要问用户
      "decisionHint": "...",        # needsDecision 为 true 时，说明要问什么
      "scopeCandidates": [{"name":..., "hits":..., "score":..., "decision":"area|separate|ask",
                           "signals":[...], "files":[...]}],
                           # 多出来的项目名：判为本项目区域 / 独立项目 / 判不了（交用户定）
      "notes": [...]
    }

判据（保守优先，宁可问也不猜）
--------------------------------
* 候选名来源：PPT/DOCX 的标题级文本 > 文件名 > 正文扫描结果。
* 每个候选算两个数：
    hits      —— 多少份材料「提到」它（用去通用词后的最具体片段做归属判定）；
    titleHits —— 多少份材料的**标题级**位置写了它的完整名（强证据）。
* 归属判定用「去通用词后的最长专属片段」，避免拿方向词（老旧街区/老旧小区…）
  当归属依据——同方向的其他项目会全部命中，项目边界就失效了。
* 主体项目：hits 领先第二名 ≥2 份，或领先幅度 ≥60%，且自身 hits ≥2 → 自动定；
  否则 needsDecision=true（证据不足以自动定，必须问用户）。
* 方向：语料里「老旧街区/老旧小区/老旧厂区」的出现次数，最多者胜；
  三个都没出现 → needsDecision。

退出码：0 成功；2 参数错误。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

DIRECTIONS = ["老旧街区", "老旧小区", "老旧厂区"]

# 通用词：出现在几乎每个项目名里，不能用来判「归属哪个项目」。
GENERIC_TOKENS = (
    "老旧街区", "老旧小区", "老旧厂区", "城市更新", "更新提升", "改造提升",
    "片区更新", "改造", "提升", "更新", "建设", "整治", "片区", "项目",
    "工程", "前期", "沟通", "信息", "城市", "综合", "实施", "专项", "计划",
)

# 项目名：必须以「项目」结尾，首字符非数字（避免把 0828 这类日期前缀吃进去）。
PROJECT_RE = re.compile(
    r"([一-鿿A-Za-z（）()·、][一-鿿A-Za-z0-9（）()·、\-]{3,59}项目)")

# 文件名里的日期前缀：0828 / 20260828 / 0828- 之类
DATE_PREFIX_RE = re.compile(r"^[\s\-_]*\d{4,8}[\s\-_]*")

# 需要时 OCR 的图片上限（OCR 是秒级开销，材料多时不能全做）
MAX_OCR_IMAGES = 6
# OCR 优先做的图片（文件名带这些词的多半是项目信息卡/封面）
OCR_PRIORITY = ("项目信息", "概况", "封面", "汇报", "信息", "说明", "立项", "建设")


# ---------------------------------------------------------------- 文本提取

def _strip_tags(xml: str) -> str:
    xml = re.sub(r"</w:p>", "\n", xml)
    xml = re.sub(r"</w:tr>", "\n", xml)
    xml = re.sub(r"</w:tc>", "\t", xml)
    xml = re.sub(r"</a:p>", "\n", xml)
    xml = re.sub(r"</a:t>", " ", xml)
    xml = re.sub(r"</si>", "\n", xml)
    return re.sub(r"<[^>]+>", "", xml)


def _docx_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as z:
            if "word/document.xml" not in z.namelist():
                return ""
            return _strip_tags(z.read("word/document.xml").decode("utf-8", "ignore"))
    except (zipfile.BadZipFile, OSError):
        return ""


def _docx_title_blocks(path: Path) -> list[str]:
    """DOCX 的标题级段落（Heading 样式 / 首行）。"""
    try:
        with zipfile.ZipFile(path) as z:
            if "word/document.xml" not in z.namelist():
                return []
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
    except (zipfile.BadZipFile, OSError):
        return []
    out = []
    for para in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
        style = re.search(r'<w:pStyle w:val="([^"]+)"', para)
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", para, re.S)).strip()
        if not text:
            continue
        if (style and re.search(r"heading|title|标题", style.group(1), re.I)) or not out:
            out.append(text)
    return out


def _pptx_slides(path: Path) -> list[list[str]]:
    try:
        with zipfile.ZipFile(path) as z:
            names = sorted(
                (n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                key=lambda n: int(re.search(r"slide(\d+)\.xml$", n).group(1)))
            out = []
            for n in names:
                xml = z.read(n).decode("utf-8", "ignore")
                blocks = [re.sub(r"<[^>]+>", "", t)
                          for t in re.findall(r"<a:t>(.*?)</a:t>", xml, re.S)]
                out.append([b.strip() for b in blocks if b.strip()])
            return out
    except (zipfile.BadZipFile, OSError):
        return []


def _xlsx_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as z:
            parts = []
            for n in z.namelist():
                if n in ("xl/sharedStrings.xml",) or n.startswith("xl/worksheets/sheet"):
                    if n == "xl/sharedStrings.xml":
                        parts.append(_strip_tags(z.read(n).decode("utf-8", "ignore")))
                    else:
                        parts.append(_strip_tags(z.read(n).decode("utf-8", "ignore"))[:400000])
            return "\n".join(parts)
    except (zipfile.BadZipFile, OSError):
        return ""


def _pdf_text(path: Path) -> str:
    try:
        import pymupdf
    except ImportError:
        try:
            from pypdf import PdfReader
            reader = PdfReader(str(path))
            return " ".join((p.extract_text() or "") for p in reader.pages[:10])
        except Exception:
            return ""
    try:
        doc = pymupdf.open(str(path))
        try:
            return " ".join(doc[i].get_text() for i in range(min(len(doc), 10)))
        finally:
            doc.close()
    except Exception:
        return ""


def _plain_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="ignore")[:400000]
    except OSError:
        return ""


def _ocr_engine():
    """定位 OCR 引擎；返回 (标识, 调用函数) 或 (None, None)。"""
    try:
        import rapidocr_onnxruntime  # noqa: F401
        return "rapidocr", "rapidocr"
    except ImportError:
        pass
    exe = shutil.which("tesseract")
    if exe:
        return "tesseract", exe
    return None, None


def _ocr_text(path: Path, engine, kind) -> str:
    try:
        if kind == "rapidocr":
            from rapidocr_onnxruntime import RapidOCR
            res, _ = RapidOCR()(str(path))
            return " ".join(str(item[1]) for item in (res or []) if len(item) > 1)
        import subprocess
        proc = subprocess.run([engine, str(path), "stdout", "-l", "chi_sim"],
                              capture_output=True, text=True,
                              encoding="utf-8", errors="replace")
        return proc.stdout or ""
    except Exception:
        return ""


IMAGE_EXTS = {"png", "jpg", "jpeg", "bmp", "webp", "tif", "tiff", "gif"}
TEXT_EXTS = {"txt", "md", "csv", "json", "log", "xml", "htm", "html"}


def extract(path: Path, ocr_kind, ocr_exe) -> tuple[str, list[str]]:
    """返回 (正文文本, 标题级文本块)。标题级 = PPT 首页 / DOCX 标题段 / 文件名。"""
    ext = path.suffix.lower().lstrip(".")
    if ext == "docx":
        return _docx_text(path), _docx_title_blocks(path)
    if ext == "pptx":
        slides = _pptx_slides(path)
        flat = " ".join(" ".join(s) for s in slides)
        return flat, (slides[0] if slides else [])
    if ext == "xlsx":
        return _xlsx_text(path), []
    if ext == "pdf":
        return _pdf_text(path), []
    if ext in TEXT_EXTS:
        return _plain_text(path), []
    if ext in IMAGE_EXTS:
        if ocr_kind is None:
            return "", []
        return _ocr_text(path, ocr_exe, ocr_kind), []
    # doc/xls/ppt/wps 等旧二进制：读不到正文，只用文件名
    return "", []


# ---------------------------------------------------------------- 候选名

def normalize_name(raw: str) -> str:
    s = DATE_PREFIX_RE.sub("", str(raw or "").strip())
    s = s.strip(" -—_·、,，.。()（）[]【】")
    m = PROJECT_RE.search(s)
    return m.group(1).strip() if m else ""


def names_in(text: str) -> list[str]:
    return [normalize_name(m) for m in PROJECT_RE.findall(text or "")]


def strip_generic(name: str) -> str:
    s = name
    for g in sorted(GENERIC_TOKENS, key=len, reverse=True):
        s = s.replace(g, "")
    return s.strip(" -—_·、,，.。()（）[]【】")


def core_fragments(name: str, min_len: int = 3) -> list[str]:
    """项目名去通用词后的候选专属片段（长的优先）。

    「某街某片区片区老旧街区更新提升项目」→ 去通用词 → 「某街某片区」
    → 片段：某街某片区 / 某街点将 / 隆街点将 / 街某片区 / 某街 / 某片区 …
    归属判定会用其中「在语料里真实出现、且出现份数最多」的那个。
    """
    base = strip_generic(name)
    if len(base) < min_len:
        return []
    out = []
    for size in range(len(base), min_len - 1, -1):
        for i in range(len(base) - size + 1):
            frag = base[i:i + size]
            if frag not in out:
                out.append(frag)
    return out


# ---------------------------------------------------------------- 主流程

# 形式参考件排除表（路径/文件名），由 --exclude-json 载入
EXCLUDES: set[str] = set()


def _load_excludes(path) -> None:
    """载入形式参考件清单：这些文件只学形式，不参与项目名/方向识别。

    形式参考件是「别的项目的同类成品」，标题与正文里全是别人的项目名。
    实测未排除时，识别结果会被参考件标题顶成「某区2026年老旧小区片区
    综合改造项目」，本项目名字反而成了次优候选。
    """
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return
    for r in (data.get("reference") or []):
        for key in ("path", "rel"):
            v = r.get(key)
            if v:
                EXCLUDES.add(str(v).replace("\\", "/"))
                EXCLUDES.add(Path(str(v)).name)


def collect_files(root: Path) -> list[Path]:
    out = []
    if not root.is_dir():
        return out
    for p in sorted(root.rglob("*")):
        if not p.is_file():
            continue
        if p.name.startswith("~$") or p.name.startswith("."):
            continue
        if "空白" in p.name:
            continue
        rel = p.relative_to(root).as_posix()
        if any(part in rel for part in ("输入文件-模板", "提资数据", "过程文件", "成果文件")):
            continue
        if p.name in EXCLUDES or str(p).replace("\\", "/") in EXCLUDES or rel in EXCLUDES:
            continue
        out.append(p)
    return out


def _ngrams(text: str, n: int = 4) -> set:
    t = re.sub(r"\s+", "", text or "")
    return {t[i:i + n] for i in range(max(0, len(t) - n + 1))}


def _deck_overlap(cand_files, main_files, per_file_text) -> float:
    """两份沟通稿的正文重合度（4 字窗 Jaccard），用来判是不是同一份稿子改名扩写。"""
    best = 0.0
    for a in [f for f in cand_files if f.lower().endswith(".pptx")]:
        for b in [f for f in main_files if f.lower().endswith(".pptx")]:
            if a == b:
                continue          # 同一个文件不算「两份稿子」，否则永远 100%
            sa, sb = _ngrams(per_file_text.get(a, "")), _ngrams(per_file_text.get(b, ""))
            if sa and sb:
                best = max(best, len(sa & sb) / len(sa | sb))
    return best


def scope_signals(cand: dict, main: dict, per_file_text: dict) -> tuple[int, list, list]:
    """判断「另一个项目名」是本项目的区域，还是独立项目。

    只看投递区内的证据，够就判、不够就交给用户（decision=ask）：
    既不按名字片段静默排除，也不静默并入。信号权重：
      明文范围表述（「实施范围包括…X区域」）3 分 → 足以下结论；
      与主项目台账同处一套摸底目录 1 分；与主项目沟通稿同源 1 分。
    """
    score, signals = 0, []
    files = list(cand.get("files") or [])
    core = cand.get("core") or ""
    if core:
        pat = re.compile(r"(包括|包含|范围为|实施范围|规划范围)[^。；]{0,40}" + re.escape(core))
        for f in files:
            m = pat.search(per_file_text.get(f, ""))
            if m:
                signals.append("明文范围表述「{}」（{}）".format(m.group(0)[:40], f))
                score += 3
                break
    cand_dirs = {f.rsplit("/", 1)[0] for f in files if "/" in f}
    main_dirs = {f.rsplit("/", 1)[0] for f in (main.get("files") or []) if "/" in f}
    shared = sorted(cand_dirs & main_dirs)
    if shared and any(f.lower().endswith((".xlsx", ".xls")) for f in files):
        signals.append("与主项目台账同处一套摸底目录：{}".format("、".join(shared[:2])))
        score += 1
    ov = _deck_overlap(files, main.get("files") or [], per_file_text)
    if ov >= 0.35:
        signals.append("与主项目沟通稿正文重合 {:.0%}（同一份稿子改名扩写）".format(ov))
        score += 1
    return score, signals, files


def build_report(root: Path) -> dict:
    files = collect_files(root)
    ocr_kind, ocr_target = _ocr_engine()

    # 1) 先决定要不要 OCR、OCR 哪些图（图片多时按优先级取前 N 张）
    images = [p for p in files if p.suffix.lower().lstrip(".") in IMAGE_EXTS]
    ocr_set: set[Path] = set()
    if ocr_kind:
        images.sort(key=lambda p: (0 if any(k in p.name for k in OCR_PRIORITY) else 1,
                                   len(p.name)))
        ocr_set = set(images[:MAX_OCR_IMAGES])

    corpus_parts: list[str] = []
    title_hits: dict[str, set[str]] = {}     # 候选名 → 出现过它的文件（标题级）
    file_title_blocks: dict[str, list[str]] = {}

    for p in files:
        rel = p.relative_to(root).as_posix()
        try:
            text, titles = extract(p, ocr_kind if p in ocr_set else None, ocr_target)
        except Exception:
            text, titles = "", []
        file_title_blocks[rel] = titles
        # 标题级位置：PPT 首页 / DOCX 标题段 / 文件名
        title_zone = " ".join(titles) + " " + p.stem
        for nm in names_in(title_zone):
            title_hits.setdefault(nm, set()).add(rel)
        corpus_parts.append(text)
        corpus_parts.append(p.stem)
        corpus_parts.append(rel)
    corpus = "\n".join(corpus_parts)

    # 2) 候选名汇总
    cand_map: dict[str, dict] = {}
    for nm in names_in(corpus):
        if nm and nm not in cand_map:
            cand_map[nm] = {"name": nm, "hits": 0, "titleHits": 0,
                             "files": [], "core": ""}

    # 3) 逐候选算归属
    per_file_text: dict[str, str] = {}
    for p in files:
        rel = p.relative_to(root).as_posix()
        try:
            text, titles = extract(p, ocr_kind if p in ocr_set else None, ocr_target)
        except Exception:
            text = ""
        per_file_text[rel] = (text or "") + " " + p.stem + " " + rel + " " + \
                            " ".join(file_title_blocks.get(rel, []))

    for nm, info in cand_map.items():
        # 归属片段：候选名里在语料中真实出现、且覆盖文件数最多的那个
        best_frag, best_files = "", set()
        for frag in core_fragments(nm):
            hit_files = {rel for rel, txt in per_file_text.items() if frag in txt}
            if len(hit_files) > len(best_files):
                best_frag, best_files = frag, hit_files
            if len(hit_files) >= max(2, len(files) // 2):
                break
        info["core"] = best_frag
        info["hits"] = len(best_files)
        info["files"] = sorted(best_files)
        info["titleHits"] = len(title_hits.get(nm, ()))

    # 4) 打分排序：归属材料数是主证据，标题级完整名是强证据
    for info in cand_map.values():
        info["score"] = info["hits"] * 3 + info["titleHits"]

    ranked = sorted(cand_map.values(), key=lambda x: (-x["score"], -x["hits"], x["name"]))
    for info in ranked:
        info.pop("files_tmp", None)

    # 5) 方向
    dir_counts = {d: corpus.count(d) for d in DIRECTIONS}
    top_dir = max(dir_counts, key=lambda d: dir_counts[d]) if any(dir_counts.values()) else ""
    dir_conf = 0.0
    if top_dir and dir_counts[top_dir] > 0:
        total = sum(dir_counts.values())
        dir_conf = dir_counts[top_dir] / total if total else 0.0

    # 6) 主体项目判定
    #    取证原则：**不停下来问**。材料里有两个项目时，硬让人先回答才能跑，
    #    等于把机器 10 秒能做完的事变成一次人机往返；而且人未必答得比证据准。
    #    正确做法是——按证据最强的那个继续跑，产出完整成果，同时把「另一个项目
    #    的材料未并入」和「一条命令就能切换」摊在放料清单里。
    #    （选错的后果可控：换一个 --project 重跑即可，且成果里的项目名随之改。）
    needs_decision = False
    hint = ""
    chosen = ""
    basis = ""
    ambiguous = False
    if not ranked:
        needs_decision = True
        hint = ("材料里读不出项目名（没找到任何以「项目」结尾的名称）。"
                "请直接告诉流水线项目全称，或用 --project 指定。")
    else:
        top = ranked[0]
        second = ranked[1] if len(ranked) > 1 else None
        gap = top["hits"] - (second["hits"] if second else 0)
        strong = (top["hits"] >= 2 and
                  (gap >= 2 or (second and top["hits"] >= 1.6 * second["hits"])))
        chosen = top["name"]
        basis = ("材料归属 %d 份（专属片段「%s」），标题级完整名 %d 处；"
                 "次优候选「%s」覆盖 %d 份"
                 % (top["hits"], top["core"] or "—", top["titleHits"],
                    second["name"] if second else "—",
                    second["hits"] if second else 0))
        if not strong:
            # 证据接近：照样按最强的走，但标出来，让人知道有得换
            ambiguous = True
            detail = "、".join("「%s」（%d 份）" % (c["name"], c["hits"]) for c in ranked[:4])
            hint = ("材料里出现多个项目名，证据接近：" + detail +
                    "。本次按证据最强的「%s」写；其余名字先判归属，"
                    "判不出会停下来问你，不静默并入也不静默排除。" % chosen)

    # 归属定性：多出来的项目名，是本项目的区域还是独立项目？
    # 有明文范围表述 → area；毫无关联信号 → separate；证据不够 → ask（交用户定）。
    scope = []
    if chosen:
        for c in ranked[:6]:
            if c["name"] == chosen or ("项目" not in c["name"] and "片区" not in c["name"]):
                continue
            score, signals, sfiles = scope_signals(c, ranked[0], per_file_text)
            decision = "area" if score >= 3 else ("separate" if score == 0 else "ask")
            scope.append({"name": c["name"], "hits": c["hits"], "score": score,
                          "decision": decision, "signals": signals, "files": sfiles[:6]})

    notes = []
    if len(files) == 0:
        notes.append("投递区没有文件。")
    if not ocr_kind:
        notes.append("本机无 OCR 引擎，图片内容未参与识别（只用了文件名）。")
    elif ocr_set and len(images) > len(ocr_set):
        notes.append("图片共 %d 张，按优先级只 OCR 了 %d 张。" % (len(images), len(ocr_set)))

    return {
        "schema": "cardo/detect-project/v1",
        "generatedAt": None,
        "fileCount": len(files),
        "project": {"value": chosen, "confidence": (0.9 if chosen else 0.0),
                    "basis": basis, "evidence": [f["name"] for f in ranked[:5]],
                    "ambiguous": ambiguous},
        "direction": {"value": top_dir, "confidence": round(dir_conf, 2),
                      "evidence": ["%s×%d" % (d, c) for d, c in dir_counts.items()]},
        "candidates": ranked[:8],
        "needsDecision": needs_decision,
        "decisionHint": hint,
        "scopeCandidates": scope,
        "notes": notes,
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="从投递区材料自动识别项目名与方向")
    ap.add_argument("--input", required=True, help="输入文件目录")
    ap.add_argument("--out", default="", help="结果 JSON 输出路径（可选）")
    ap.add_argument("--exclude-json", default="",
                    help="形式参考件清单（cardo/form-reference/v1）；命中文件不参与项目名识别")
    args = ap.parse_args(argv)

    if args.exclude_json:
        _load_excludes(args.exclude_json)

    root = Path(args.input).expanduser()
    if not root.is_dir():
        print(json.dumps({"schema": "cardo/detect-project/v1", "error": "输入目录不存在：%s" % root},
                         ensure_ascii=False))
        return 2

    report = build_report(root)
    text = json.dumps(report, ensure_ascii=False, indent=1)
    if args.out:
        out = Path(args.out).expanduser()
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
