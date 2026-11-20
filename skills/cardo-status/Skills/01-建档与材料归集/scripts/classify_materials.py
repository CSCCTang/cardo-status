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
PDF_MAX_PAGES = 10
CONTENT_MAX_HITS = 3
WEIGHT_FILENAME = 3.0
WEIGHT_TAIL = 4.0
WEIGHT_PATH = 8.0
WEIGHT_EXT = 2.0
WEIGHT_CONTENT = 2.0

# 无文本层的格式：CAD 图纸/字体/形状、压缩包、工程模型等。
# 这些文件读不出正文，只能靠「扩展名 + 文件名」判定；
# 若仍按「置信度不足 → 待确认」处理，图纸类项目会永久卡在人工闸门。
TEXTLESS_EXTS = {"dwg", "dxf", "shx", "fas", "dgn", "rvt", "mdb", "7z", "rar", "zip", "log"}

# 旧版二进制 Office：读不出正文（只能按文件名判），但归档去向并不歧义。
# 每个 .doc/.xls/.ppt 都变成一次人工闸门是没有收益的打扰——它们不影响能否写报告，
# 只影响归档目录。故按「文件名线索 → 扩展名归类」自动落位。
LEGACY_OFFICE_EXTS = {"doc", "xls", "ppt", "wps"}

# 可按纯文本读取的格式（补充 .log/.rtf/.xml 等常见日志与文本件）
PLAIN_TEXT_EXTS = {"txt", "md", "csv", "json", "log", "xml", "htm", "html", "ini", "yaml", "yml"}

# 纯图片格式：读不到正文，且天然同属多个类别（图纸/影像/批复/策划），
# 靠文件名与路径特征归位，不进人工闸门。
IMAGE_EXTS = {"jpg", "jpeg", "png", "gif", "bmp", "webp", "heic", "tif", "tiff"}


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


def _exclude_set(args) -> set[str]:
    """形式参考件排除集：登记 JSON 里的 excluded 相对路径 + 绝对路径。"""
    p = getattr(args, "exclude_json", None)
    if not p:
        return set()
    try:
        d = json.loads(Path(p).read_text(encoding="utf-8"))
    except Exception:
        return set()
    out = set()
    for x in (d.get("excluded") or []):
        out.add(str(x).replace("\\", "/"))
    for ref in (d.get("reference") or []):
        for k in ("rel", "path", "copy"):
            if ref.get(k):
                out.add(str(ref[k]).replace("\\", "/"))
    return out


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
    parser.add_argument("--exclude-json", help="形式参考件登记 JSON；其中 excluded 列出的文件不进分类/正文")
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


def iter_delivery_files(delivery: Path, ignore_globs: list[str], exclude: set[str] | None = None):
    exclude = exclude or set()
    for dirpath, dirnames, filenames in os.walk(delivery):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        for name in sorted(filenames):
            if any(fnmatch.fnmatch(name, pattern) for pattern in ignore_globs):
                continue
            path = Path(dirpath) / name
            if not path.is_file():
                continue
            rel = rel_posix(path, delivery)
            if rel in exclude or path.as_posix() in exclude:
                continue
            yield path, rel


def dedupe_by_content(files: list[tuple[Path, str]]) -> tuple[list[tuple[Path, str]], list[dict]]:
    """按「文件名 + 大小 + 内容哈希」去重。

    典型场景：用户投了压缩包，01 又把包解开成「解包-<名字>/」，
    而包内文件在投递区里本来就有同名目录 —— 于是同一份材料被登记两次，
    人工确认量凭空翻倍。此处按内容去重，保留先出现的那份（路径更短者优先）。
    """
    seen: dict[tuple, str] = {}
    keep: list[tuple[Path, str]] = []
    dupes: list[dict] = []
    for path, rel in sorted(files, key=lambda x: (len(x[1]), x[1])):
        try:
            size = path.stat().st_size
            digest = sha256_file(path)
        except OSError:
            keep.append((path, rel))
            continue
        key = (path.name.lower(), size, digest)
        if key in seen:
            dupes.append({"file": rel, "duplicateOf": seen[key]})
            continue
        seen[key] = rel
        keep.append((path, rel))
    return keep, dupes


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
        if ext in PLAIN_TEXT_EXTS:
            return path.read_text(encoding="utf-8", errors="ignore")[:MAX_SAMPLE_BYTES], f"{ext} 文本"
        if ext == "pdf":
            sample, note = read_pdf_sample(path)
            return sample, note
        if ext in ("doc", "xls", "ppt", "wps", "dwg", "dxf"):
            return "", f"{ext} 旧版二进制格式，未读正文"
        return "", f"{ext or '无扩展名'} 未读取正文"
    except (zipfile.BadZipFile, OSError, RuntimeError) as exc:
        return "", f"读取失败（{type(exc).__name__}），仅按文件名与路径判定"


def _ocr_engine() -> str:
    """探测可用的 OCR 引擎。

    优先 RapidOCR（纯 onnx，模型随包，离线可用）；其次 tesseract 命令行。
    返回引擎标识；不可用时返回空串。OCR 只用于「有文本层的 PDF 之外」的扫描件，
    属于可选增强：没有引擎时扫描件被如实标记为待 OCR，而不是当作空文件。
    """
    try:
        import rapidocr_onnxruntime  # noqa: F401
        return "rapidocr"
    except ImportError:
        pass
    exe = shutil.which("tesseract")
    if exe:
        return "tesseract:" + exe
    for env in ("CARDO_OCR_MODEL_DIR",):
        d = os.environ.get(env)
        if d and os.path.isdir(d) and any(f.lower().endswith(".onnx") for f in os.listdir(d)):
            return "onnx:" + d
    return ""


def read_pdf_sample(path: Path) -> tuple[str, str]:
    """PDF 正文取样，返回 (可读文本, 说明)。

    引擎优先级：PyMuPDF（提取质量最好，支持页面渲染）→ pypdf（纯 Python 兜底）。
    两者都读不到文本时判定为扫描件，并区分两类：
      * A4 文书型扫描件（请示、批件、盖章件）：有 OCR 引擎则自动 OCR，
        否则标注「需 OCR」——这类文件通常含正文，值得 OCR；
      * CAD 图幅型（立面图/平面图/大样图）：即使 OCR 也只能得到图面标注，
        对现状分析无实质增量，一律归「图纸（无正文）」，不做 OCR。
    """
    text, note = "", ""
    try:
        import pymupdf
        doc = pymupdf.open(str(path))
        try:
            chunks = [doc[i].get_text() for i in range(min(len(doc), PDF_MAX_PAGES))]
            text = " ".join(chunks)
            page_size = (doc[0].rect.width, doc[0].rect.height) if len(doc) else (0, 0)
        finally:
            doc.close()
        note = "pdf 正文（PyMuPDF）"
    except ImportError:
        try:
            from pypdf import PdfReader
            reader = PdfReader(str(path))
            chunks = [(p.extract_text() or "") for p in reader.pages[:PDF_MAX_PAGES]]
            text = " ".join(chunks)
            note = "pdf 正文（pypdf）"
            page_size = (612, 792)
        except ImportError:
            return "", "pdf 未解析（本机无 PDF 解析库），仅按文件名与路径判定"
        except Exception:
            return "", "pdf 解析失败，仅按文件名与路径判定"
    except Exception:
        return "", "pdf 解析失败，仅按文件名与路径判定"

    if text.strip():
        return text[:MAX_SAMPLE_BYTES], note

    # 无文本层 → 扫描件
    if _is_cad_sheet(page_size):
        return "", "pdf 图纸（CAD图幅，无文本层；图面标注不作为正文引用）"

    engine = _ocr_engine()
    if engine:
        ocr_text = _ocr_pdf(path, engine)
        if ocr_text.strip():
            return ocr_text[:MAX_SAMPLE_BYTES], f"pdf 扫描件（OCR：{engine.split(':')[0]}）"
        return "", "pdf 扫描件（OCR 未识别出文字）"
    return "", "pdf 扫描件无文本层（需 OCR：本机无 OCR 引擎），仅按文件名与路径判定"


def _is_cad_sheet(page_size) -> bool:
    """判断是否为 CAD 图幅 PDF（图幅尺寸显著大于 A4：1 个标准单位 = 1pt）。"""
    try:
        w, h = float(page_size[0]), float(page_size[1])
    except (TypeError, ValueError, IndexError):
        return False
    # A4 = 595×842pt；A3 = 842×1191pt。图幅页边长超过 A3 视为 CAD 图纸。
    return max(w, h) > 1200


PDF_TEXT_DIR = "02_分类材料/11_PDF文本"
PDF_TEXT_SCHEMA = "cardo/pdf-text/v1"
PDF_TEXT_MAX_BYTES = 8_000_000


def extract_pdf_pages(path: Path) -> tuple[list[str], str]:
    """逐页抽取 PDF 文本层，返回 (每页文本, 状态)。

    状态：text=有文本层；scanned=扫描件；cad=CAD 图幅；no_lib=本机无解析库；error=解析失败。
    只抽文本层，**不做 OCR**：扫描件如实报 scanned，改由开工提示让用户提供文字版。
    """
    pages: list[str] = []
    page_size = (0.0, 0.0)
    try:
        import pymupdf
        doc = pymupdf.open(str(path))
        try:
            if len(doc):
                page_size = (doc[0].rect.width, doc[0].rect.height)
            pages = [doc[i].get_text() for i in range(len(doc))]
        finally:
            doc.close()
    except ImportError:
        try:
            from pypdf import PdfReader
            reader = PdfReader(str(path))
            pages = [(p.extract_text() or "") for p in reader.pages]
            page_size = (612, 792)
        except ImportError:
            return [], "no_lib"
        except Exception:
            return [], "error"
    except Exception:
        return [], "error"

    if any(p.strip() for p in pages):
        return pages, "text"
    if _is_cad_sheet(page_size):
        return [], "cad"
    return [], "scanned"


def write_pdf_text(project_dir: Path, results: list[dict]) -> tuple[list[dict], dict]:
    """把有文本层的 PDF 逐页写成 `<PDF文本目录>/<stem>.blocks.txt`（页首 `[PAGE n]`）。

    只写派生文本，不改 `分类清单.json`（避免下游材料清单结构变化）；扫描件/CAD 一律跳过。
    """
    out_dir = project_dir / PDF_TEXT_DIR
    entries: list[dict] = []
    stats = {"pdfs": 0, "withTextLayer": 0, "scanned": 0, "cad": 0, "noLib": 0, "failed": 0}
    for item in results:
        if str(item.get("ext") or "").lower() != "pdf":
            continue
        stats["pdfs"] += 1
        src = Path(str(item.get("abs")))
        pages, status = extract_pdf_pages(src)
        if status == "text":
            stats["withTextLayer"] += 1
        elif status in stats:
            stats[status] += 1
        if not pages:
            continue
        chunks: list[str] = []
        used = 0
        for idx, page in enumerate(pages, 1):
            text = page.strip()
            if not text:
                continue
            block = f"[PAGE {idx}]\n{text}\n"
            block_bytes = len(block.encode("utf-8"))
            if chunks and used + block_bytes > PDF_TEXT_MAX_BYTES:
                break
            chunks.append(block)
            used += block_bytes
        if not chunks:
            continue
        out_dir.mkdir(parents=True, exist_ok=True)
        name = Path(str(item.get("rel") or src.name)).stem + ".blocks.txt"
        target = out_dir / name
        if target.exists():
            target = out_dir / f"{Path(name).stem}__{sha256_file(src)[:8]}.blocks.txt"
        write_text(target, "".join(chunks))
        entries.append({
            "file": item.get("rel"),
            "textFile": rel_posix(target, project_dir),
            "categoryId": "pdf_text",
            "categoryLabel": "非叙事材料 / PDF 文本",
            "usage": "fact_only",
            "pages": len(chunks),
            "chars": sum(len(c) for c in chunks),
            "truncated": used >= PDF_TEXT_MAX_BYTES,
        })

    if stats["pdfs"]:
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            payload = {
                "schema": PDF_TEXT_SCHEMA,
                "generatedAt": now_str(),
                "dir": PDF_TEXT_DIR,
                "note": ("有文本层的 PDF 逐页抽取，只供事实抽取与引用；"
                         "不作章节结构来源、不生成整章正文；扫描件不解析。"),
                "stats": stats,
                "files": entries,
            }
            write_text(out_dir / "PDF文本清单.json",
                       json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
        except OSError as exc:
            print(f"[提示] PDF 文本清单写入失败（不影响分类）：{type(exc).__name__}: {exc}")
    return entries, stats


def _ocr_pdf(path: Path, engine: str) -> str:
    """对扫描件执行 OCR。rapidocr 走 onnx 引擎，tesseract 走命令行。"""
    try:
        if engine == "rapidocr":
            return _ocr_pdf_rapidocr(path)
        if engine.startswith("tesseract:"):
            exe = engine.split(":", 1)[1]
            import subprocess
            import tempfile
            with tempfile.TemporaryDirectory() as tmp:
                base = os.path.join(tmp, "page")
                rendered = 0
                try:
                    import pymupdf
                    doc = pymupdf.open(str(path))
                    try:
                        for i in range(min(len(doc), PDF_MAX_PAGES)):
                            doc[i].get_pixmap(dpi=200).save(f"{base}{i}.png")
                            rendered = i + 1
                    finally:
                        doc.close()
                except Exception:
                    return ""
                out = []
                for i in range(rendered):
                    proc = subprocess.run([exe, f"{base}{i}.png", "stdout", "-l", "chi_sim"],
                                          capture_output=True, text=True,
                                          encoding="utf-8", errors="replace")
                    out.append(proc.stdout or "")
                return " ".join(out)
        # onnx 引擎：交给 PyMuPDF 的 Tesseract 兼容层
        import pymupdf
        doc = pymupdf.open(str(path))
        try:
            chunks = []
            for i in range(min(len(doc), PDF_MAX_PAGES)):
                try:
                    tp = doc[i].get_textpage_ocr(language="chi_sim", dpi=200)
                    chunks.append(tp.extractText())
                except Exception:
                    break
            return " ".join(chunks)
        finally:
            doc.close()
    except Exception:
        return ""


_OCR_INSTANCE = None


def _ocr_pdf_rapidocr(path: Path) -> str:
    """RapidOCR（onnx）识别扫描件：渲染 → 逐页识别 → 按行拼接。"""
    global _OCR_INSTANCE
    try:
        import numpy as np
        import pymupdf
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        return ""
    try:
        doc = pymupdf.open(str(path))
    except Exception:
        return ""
    try:
        if _OCR_INSTANCE is None:
            _OCR_INSTANCE = RapidOCR()      # 模型只加载一次，后续文件复用
        ocr = _OCR_INSTANCE
        chunks = []
        for i in range(min(len(doc), PDF_MAX_PAGES)):
            try:
                pix = doc[i].get_pixmap(dpi=200)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                    pix.height, pix.width, pix.n)
                result, _ = ocr(img)
            except Exception:
                break
            if not result:
                continue
            lines = [str(item[1]) for item in result if len(item) > 1 and str(item[1]).strip()]
            chunks.append(" ".join(lines))
        return "\n".join(chunks)
    except Exception:
        return ""
    finally:
        try:
            doc.close()
        except Exception:
            pass


# ---------------------------------------------------------------- 判定


def category_map(categories: list[dict]) -> dict[str, dict]:
    return {str(cat.get("id")): cat for cat in categories if cat.get("id")}


# 真缺口口径（进数据缺口清单）：文件完全读不出来——空文件 / 损坏 / 加密。
# 低置信度不是缺口：按文件名、目录、扩展名、正文能判到哪一类就入哪一类，
# 登记低置信度即可，不打断流水线（真实用户不会替工具人肉分十几个类）。
OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def unreadable_reason(path: Path, size: int, content_note: str) -> str:
    """返回真缺口原因；可读（含无文本层但有文件名/扩展名线索）返回空串。"""
    if size <= 0:
        return "空文件（0 字节）"
    note = content_note or ""
    if "读取失败" not in note and "解析失败" not in note:
        return ""
    try:
        with open(path, "rb") as fh:
            head = fh.read(8)
    except OSError as exc:
        return f"无法打开（{type(exc).__name__}）"
    if head == OLE2_MAGIC:
        return "加密容器（OLE2），正文不可读"
    return "文件损坏或加密，正文不可读"


# ---------------------------------------------------------------- 上游成果整包
# AHP 成果 = 指标权重与评分依据；OSGB 成果 = 量化指标与体检依据。
# 按整包识别（目录形态 + 文件名 + 内部结构），认领后只取规定该取的件；
# 其余中间件与模型资产登记「已见、不入正文」，不撒进普通材料分类。
UPSTREAM_PKG_HINTS = ("成果文件", "成品版", "内部件")
UPSTREAM_CLAIM_NAMES = ("权重表", "评分表", "现状指标", "扩展指标",
                        "指标说明", "体检报告", "权重与评分结论", "依据对照表")


def upstream_kind(rel: str) -> str:
    """按路径判定上游包类型：AHP / OSGB。"""
    low = rel.lower()
    if "ahp" in low or any(k in rel for k in ("权重", "评分", "矩阵", "一致性")):
        return "AHP"
    return "OSGB"


def upstream_item(name: str) -> str:
    """认领件名；不属于规定取件（中间件/模型资产）返回空串。"""
    if "现状指标" in name and "依据对照" in name:
        return "现状指标依据对照表"
    if "现状指标" in name:
        return "现状指标"
    if "扩展指标" in name:
        return "扩展指标"
    if "指标说明" in name:
        return "指标说明"
    if "体检报告" in name or "模型体检" in name:
        return "体检报告"
    if "权重与评分结论" in name:
        return "AHP结论"
    if "权重表" in name:
        return "权重表"
    if "评分表" in name:
        return "评分表"
    if "依据对照表" in name:
        return "依据对照表"
    return ""


def in_upstream_package(rel: str) -> bool:
    parts = rel.replace("\\", "/").split("/")
    return any(h in part for part in parts for h in UPSTREAM_PKG_HINTS)


def classify_upstream(path: Path, rel: str, categories: list[dict]) -> dict | None:
    """整包识别上游成果；非上游包返回 None（交给普通分类）。"""
    if not in_upstream_package(rel):
        return None
    cat = next((c for c in categories if str(c.get("id")) == "upstream"), None)
    if cat is None:
        return None
    kind = upstream_kind(rel)
    item = upstream_item(path.name)
    in_body = bool(item)
    size = path.stat().st_size
    dest = f"02_分类材料/12_上游成果/{kind}/{path.name}"
    return {
        "rel": rel,
        "abs": str(path),
        "size": size,
        "size_h": format_size(size),
        "ext": path.suffix.lower().lstrip("."),
        "content_note": "",
        "score": 100.0,
        "confidence": 1.0,
        "gap": 1.0,
        "category_id": str(cat.get("id")),
        "category_label": str(cat.get("label") or "上游成果（AHP/OSGB）"),
        "pending": False,
        "pending_reason": "",
        "pending_candidates": [],
        "low_confidence": False,
        "low_confidence_reason": "",
        "unreadable": False,
        "unreadable_reason": "",
        "evidence": [f"上游成果整包识别（{kind}）",
                     (f"认领件：{item}（进正文依据）" if in_body else "中间件/模型资产：已见、不入正文")],
        "dest_rel": dest,
        "upstream": {"kind": kind, "item": item or "中间件", "inBody": in_body},
    }


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
        "low_confidence": False,
        "low_confidence_reason": "",
        "unreadable": False,
        "unreadable_reason": "",
        "evidence": [],
        "dest_rel": "",
    }
    unreadable = unreadable_reason(path, size, content_note)
    result["unreadable"] = bool(unreadable)
    result["unreadable_reason"] = unreadable
    fallback_cat = next((c for c in categories if str(c.get("id")) == "other"), categories[-1])

    def fallback_top() -> dict:
        return {"id": str(fallback_cat.get("id")),
                "label": str(fallback_cat.get("label") or ""),
                "dir": str(fallback_cat.get("dir") or ""), "score": 0.0,
                "evidence": [],
                "preserve_subdir": bool(fallback_cat.get("preserve_subdir"))}

    if not scored:
        fb = fallback_top()
        result.update(
            category_id=fb["id"],
            category_label=fb["label"],
            evidence=["无匹配：扩展名/文件名/路径/正文均未命中规则，按兜底类别入档"],
            pending_candidates=[{"label": str(c.get("label")), "score": 0.0} for c in categories],
        )
        result["_top"] = fb
        if unreadable:
            result["low_confidence_reason"] = f"文件不可读：{unreadable}；按兜底类别入档并登记数据缺口"
        else:
            result["low_confidence"] = True
            result["low_confidence_reason"] = ("没有任何类别命中（规则表可能缺这一类，或文件名太泛），"
                                               "按兜底类别入档")
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
    # 无文本层格式（CAD 图纸/压缩包等）：扩展名本身就是强证据。
    # 这类文件读不到正文，若仍走「置信度不足 → 待确认」，图纸类项目会永远卡在人工闸门，
    # 因此当扩展名唯一命中一个类别时直接按该类别落位，并注明判定依据。
    textless_by_ext = ext in TEXTLESS_EXTS
    ext_only_hit = (textless_by_ext and len(scored) == 1
                    and scored[0]["score"] <= WEIGHT_EXT
                    and not any(e.startswith("文件名命中") or e.startswith("路径命中")
                                or e.startswith("正文命中") for e in scored[0]["evidence"]))
    # 扩展名唯一命中的兜底类（如 .log → 其他）：同样直接落位。
    # 兜底类的语义是「按扩展名归入其他」，若还要人确认，图纸类项目会一直卡住。
    ext_only_other = (ext_only_hit and top["id"] == "other")
    # 旧二进制 Office（doc/xls/ppt/wps）：读不到正文，只能按文件名+扩展名判。
    # 若按「置信度不足」处理，每个 .doc 材料都会变成一次人工闸门——
    # 而这类文件的去处其实并无歧义（按扩展名归入对应类别即可）。
    legacy_office = ext in LEGACY_OFFICE_EXTS
    # 图片类（jpg/png/…）同属多个类别（图纸/影像/批复/策划），必然并列低分。
    # 照片/截图按「影像资料」归位；CAD 目录下的图片按「图纸」归位（路径已说明来源）。
    if ext in IMAGE_EXTS and not ext_only_hit:
        path_txt = rel.lower()
        cad_dir = any(k in path_txt for k in ("cad", "图纸", "图框", "dwg", "dxf"))
        prefer = "mapping" if cad_dir else "media"
        hit = next((s for s in scored if s["id"] == prefer), None)
        if hit:
            result.update(
                score=hit["score"] + 1.0,
                confidence=0.9,
                gap=1.0,
                category_id=hit["id"],
                category_label=hit["label"],
                pending_candidates=[{"label": s["label"], "score": s["score"]} for s in scored[:3]],
                evidence=list(hit["evidence"]) + [
                    ("图片文件且位于 CAD/图纸目录，按「测绘/图纸」归位" if cad_dir
                     else "图片文件按「影像资料」归位（文件名/路径无其他特征）")],
            )
            result["pending"] = False
            result["pending_reason"] = ""
            result["_top"] = hit
            return result
    # CAD 图幅型 PDF（无文本层，图面标注不作正文）：按「测绘/图纸」直接落位，
    # 与 .dwg/.dxf 同理——它们是图件，不是待判定的文书。
    cad_sheet = "pdf 图纸（CAD图幅" in (content_note or "")
    if ext_only_hit or cad_sheet:
        if cad_sheet and not ext_only_hit:
            draw = [s for s in scored if s["id"] == "mapping"]
            if len(draw) != 1:
                draw = []
            if not draw:
                fb = fallback_top()
                result.update(
                    category_id=fb["id"],
                    category_label=fb["label"],
                    low_confidence=True,
                    low_confidence_reason="CAD 图幅 PDF，且无唯一命中的图纸类别；按兜底类别入档",
                    pending_candidates=[{"label": s["label"], "score": s["score"]} for s in scored[:3]],
                    evidence=list(top["evidence"]) + ["CAD 图幅 PDF（无文本层）"],
                )
                result["_top"] = fb
                return result
            top = draw[0]
            result.update(
                score=top["score"],
                confidence=0.9,
                gap=1.0,
                category_id=top["id"],
                category_label=top["label"],
                evidence=list(top["evidence"]) + ["CAD 图幅 PDF（无文本层，图面标注不作正文引用）"],
                pending_candidates=[{"label": s["label"], "score": s["score"]} for s in scored[:3]],
            )
        else:
            result.update(
                evidence=list(top["evidence"]) + [
                    f"无文本层格式（.{ext}），按扩展名唯一命中类别「{top['label']}」直接落位"],
                note="按扩展名落位（无正文可读）",
            )
        result["pending"] = False
        result["pending_reason"] = ""
        result["_top"] = top
        return result

    # 旧二进制 Office：文件名有明确线索（批复/清单/纪要…）就按该线索落位；
    # 文件名无线索时，按扩展名归入「其他」——归类不精确，但**不该为此打断用户**。
    # 这些文件本来也读不出正文，分类只决定归档去向，不影响能否写报告。
    if legacy_office:
        name_hit = [s for s in scored
                    if any(e.startswith("文件名命中") for e in s["evidence"])]
        if name_hit:
            best = max(name_hit, key=lambda s: s["score"])
            result.update(
                score=best["score"], confidence=0.75, gap=1.0,
                category_id=best["id"], category_label=best["label"],
                evidence=list(best["evidence"]) + [
                    f"旧版 {ext} 格式读不出正文，按文件名线索归位"],
                pending_candidates=[{"label": s["label"], "score": s["score"]} for s in scored[:3]],
            )
        else:
            others = next((s for s in scored if s["id"] == "other"), None)
            if others is None:
                others = scored[0]
            result.update(
                score=others["score"], confidence=0.5, gap=0.0,
                category_id=others["id"], category_label=others["label"],
                evidence=list(others["evidence"]) + [
                    f"旧版 {ext} 格式读不出正文且文件名无线索，按「{others['label']}」归位"
                    "（仅决定归档去向，不影响成稿）"],
                pending_candidates=[{"label": s["label"], "score": s["score"]} for s in scored[:3]],
            )
        result["pending"] = False
        result["pending_reason"] = ""
        result["_top"] = result["_top"] if result.get("_top") else next(
            s for s in scored if s["id"] == result["category_id"])
        return result

    if top_score < min_score:
        reasons.append(f"最高分 {top_score:g} < 阈值 {min_score:g}")
    if confidence < conf_threshold:
        reasons.append(f"置信度 {confidence:.2f} < 阈值 {conf_threshold:.2f}")
    if second and gap < gap_threshold:
        reasons.append(f"与次优类别「{second['label']}」差距 {gap:.2f} < 阈值 {gap_threshold:.2f}")
    if top["id"] == "other":
        reasons.append("只命中兜底类别「其他」")
    if reasons:
        # 判不准不等于要人肉确认：照常入档，登记低置信度，让流水线继续。
        # 只有「文件完全读不出来」才登记为真缺口（见 unreadable_reason）。
        result["low_confidence"] = True
        result["low_confidence_reason"] = "；".join(reasons)
        if unreadable:
            result["low_confidence_reason"] += f"；文件不可读：{unreadable}（登记数据缺口）"
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
                # 重复执行：目标已在位且内容一致，无需再复制，也不算缺口。
                item["already_in_place"] = True
                continue
            # 目标已存在且内容不同：换一个不冲突的名字落位，不因命名冲突停流水线。
            stem, suffix = os.path.splitext(name)
            for n in range(2, 1000):
                cand = f"{dest_dir}/{stem}__{n}{suffix}"
                if not (project_dir / cand).exists():
                    item["dest_rel"] = cand
                    item["dest_conflict_note"] = f"目标同名但内容不同，改名落位：{cand}"
                    break

    for item in results:
        item.pop("_top", None)


# ---------------------------------------------------------------- 产物


def build_summary(results: list[dict], categories: list[dict]) -> dict:
    by_category: dict[str, int] = {}
    pending = [r for r in results if r["pending"]]
    low = [r for r in results if r.get("low_confidence")]
    gaps = [r for r in results if r.get("unreadable")]
    for item in results:
        if not item["pending"] and item["category_label"]:
            by_category[item["category_label"]] = by_category.get(item["category_label"], 0) + 1
    return {
        "total": len(results),
        "auto": len(results) - len(pending),
        "pending": len(pending),
        "low_confidence": len(low),
        "unreadable": len(gaps),
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
    lines.append(f"| 其中低置信度入档 | {summary.get('low_confidence', 0)} |")
    lines.append(f"| 文件不可读（真缺口） | {summary.get('unreadable', 0)} |")
    by_cat = "；".join(f"{label} {count}" for label, count in summary["by_category"].items()) or "—"
    lines.append(f"| 按类别 | {by_cat} |")
    lines.append("")
    lines.append("## 归档计划")
    lines.append("")
    header = "| # | 文件（相对投递区） | 大小 | 初判类别 | 判定依据 | 置信度 | 建议去向 | 低置信度？ |"
    lines.append(header)
    lines.append("|---|---|---|---|---|---|---|---|")
    for idx, item in enumerate(results, start=1):
        dest = item["dest_rel"] or "（未定）"
        pending = "是" if item.get("low_confidence") else ""
        conf = f"{item['confidence']:.2f}" if item["category_label"] else "—"
        lines.append(
            f"| {idx} | {esc(item['rel'])} | {item['size_h']} | {esc(item['category_label'] or '待定')} | "
            f"{esc('；'.join(item['evidence']))} | {conf} | {esc(dest)} | {pending} |"
        )
    lines.append("")
    lines.append("## 低置信度入档（不阻断流水线，登记备查）")
    lines.append("")
    lines.append("> 判不准的文件照样入档：按现有线索取最接近的类别，不让人工确认拦住整条流水线。")
    lines.append("> 下列条目同时登记进数据缺口清单（口径存疑，补齐材料后可重跑校正）。")
    lines.append("")
    low_items = [r for r in results if r.get("low_confidence")]
    if low_items:
        lines.append("| 编号 | 文件 | 入档类别 | 置信度 | 说明 |")
        lines.append("|---|---|---|---|---|")
        for idx, item in enumerate(low_items, start=1):
            lines.append(
                f"| L{idx:02d} | {esc(item['rel'])} | {esc(item['category_label'])} | "
                f"{item['confidence']:.2f} | {esc(item['low_confidence_reason'])} |")
    else:
        lines.append("无。全部文件均以高置信度落位。")
    lines.append("")
    lines.append("## 数据缺口（文件不可读）")
    lines.append("")
    gap_items = [r for r in results if r.get("unreadable")]
    if gap_items:
        lines.append("| 编号 | 文件 | 缺口原因 | 处置 |")
        lines.append("|---|---|---|---|")
        for idx, item in enumerate(gap_items, start=1):
            lines.append(
                f"| G{idx:02d} | {esc(item['rel'])} | {esc(item['unreadable_reason'])} | "
                f"已按兜底类别入档；请补可读版本后重跑 |")
    else:
        lines.append("无。所有文件均可读取（含无文本层但可按文件名/扩展名判定者）。")
    lines.append("")
    lines.append("## 上游成果认领（AHP / OSGB）")
    lines.append("")
    up_items = [r for r in results if r.get("upstream")]
    if up_items:
        lines.append("> AHP = 指标权重与评分依据；OSGB = 量化指标与体检依据。")
        lines.append("> 只取规定该取的件；中间件与模型资产登记「已见、不入正文」。")
        lines.append("")
        lines.append("| 包 | 件 | 角色 | 是否入正文 |")
        lines.append("|---|---|---|---|")
        for item in up_items:
            u = item["upstream"]
            lines.append(f"| {u['kind']} | {esc(item['rel'])} | {esc(u['item'])} | "
                         f"{'是' if u.get('inBody') else '否（已见、不入正文）'} |")
    else:
        lines.append("未发现上游成果整包。")
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
        # 低置信度入档：已落位、已复制，不阻断流水线；登记备查并进数据缺口清单。
        "low_confidence": [
            {
                "no": f"L{idx:02d}",
                "file": item["rel"],
                "category_id": item["category_id"],
                "category_label": item["category_label"],
                "confidence": item["confidence"],
                "reason": item["low_confidence_reason"],
                "dest": item["dest_rel"],
            }
            for idx, item in enumerate([r for r in results if r.get("low_confidence")], start=1)
        ],
        # 真缺口：文件完全读不出来（空文件 / 损坏 / 加密），已按兜底类别入档。
        "data_gaps": [
            {
                "no": f"G{idx:02d}",
                "file": item["rel"],
                "kind": "材料不可读",
                "reason": item["unreadable_reason"],
                "dest": item["dest_rel"],
            }
            for idx, item in enumerate([r for r in results if r.get("unreadable")], start=1)
        ],
        # 上游成果认领：AHP（权重/评分依据）、OSGB（量化指标/体检依据）。
        "upstream": [
            {
                "file": item["rel"],
                "kind": item["upstream"]["kind"],
                "item": item["upstream"]["item"],
                "inBody": item["upstream"]["inBody"],
                "dest": item["dest_rel"],
            }
            for item in results if item.get("upstream")
        ],
        "actions_executed": meta.get("executed", []),
    }
    write_text(path, json.dumps(payload, ensure_ascii=False, indent=2) + "\n")


# ---------------------------------------------------------------- 执行


def execute_actions(project_dir: Path, results: list[dict], move: bool) -> list[dict]:
    executed: list[dict] = []
    for item in results:
        if item["pending"] or not item["dest_rel"] or item.get("already_in_place"):
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

    files = list(iter_delivery_files(delivery_path, ignore_globs, _exclude_set(args)))
    if not files:
        print(f"投递区为空：{delivery_path}")
        print("（没有可分类文件；请确认用户是否还没投递材料。）")
        return 0

    files, dupes = dedupe_by_content(files)
    if dupes:
        print(f"[去重] 同一份材料在投递区出现多次，已按内容去重 {len(dupes)} 个"
              f"（压缩包解开后与原目录重复）")

    # 上游成果整包优先认领（AHP / OSGB），其余走普通材料分类。
    results = [classify_upstream(path, rel, categories) or
               classify_one(path, rel, rules, categories)
               for path, rel in files]
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

    # 派生：有文本层的 PDF 逐页抽全文（只供事实抽取与引用，不进主线）
    try:
        _, pdf_stats = write_pdf_text(project_dir, results)
    except OSError as exc:
        print(f"[提示] PDF 文本抽取失败（不影响分类）：{type(exc).__name__}: {exc}")
        pdf_stats = {"pdfs": 0}
    if pdf_stats.get("pdfs"):
        tail = f"；扫描件 {pdf_stats.get('scanned', 0)} 份、CAD 图幅 {pdf_stats.get('cad', 0)} 份未解析"
        if pdf_stats.get("noLib"):
            tail += f"；{pdf_stats['noLib']} 份因本机无 PDF 解析库未抽"
        print(f"[PDF] 共 {pdf_stats['pdfs']} 份：有文本层 {pdf_stats.get('withTextLayer', 0)} 份"
              f"已抽全文 → {PDF_TEXT_DIR}{tail}")

    # ---- 控制台摘要 ----
    print(f"项目：{project}")
    print(f"投递区：{delivery_label}（{disp_path(delivery_path) if user_mode else delivery_path}）")
    print(f"模式：{mode_text}")
    print(f"文件总数 {summary['total']}：已入档 {summary['auto']}"
          f"（其中低置信度 {summary.get('low_confidence', 0)}），"
          f"文件不可读 {summary.get('unreadable', 0)}")
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
    if summary.get("low_confidence"):
        print("")
        print(f"[入档] {summary['low_confidence']} 个文件按低置信度入档（不阻断流水线，已登记缺口）：")
        for item in [r for r in results if r.get("low_confidence")]:
            print(f"  - {item['rel']} → {item['category_label']}：{item['low_confidence_reason']}")
    if summary.get("unreadable"):
        print("")
        print(f"[真缺口] {summary['unreadable']} 个文件完全不可读（空/损坏/加密），已登记数据缺口：")
        for item in [r for r in results if r.get("unreadable")]:
            print(f"  - {item['rel']}：{item['unreadable_reason']}")
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