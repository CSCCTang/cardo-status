#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""pdf_render.py -- 成品 PDF：docx → A4 HTML → 无头 Edge 打印 → 目录页码 → 页脚。

为什么不用 Word 转 PDF：交付环境没有 Office，且目录页码/页脚必须在离线条件下可复现。
本脚本复用 AHP 线验证过的链路：python-docx 读成稿 → 生成 A4 HTML → Edge headless
--print-to-pdf → PyMuPDF 定位章标题页码 → 回填目录 → 二次打印 → 烘入页脚。

用法：
  python pdf_render.py --docx 成果.docx --out 成果.pdf [--title 标题]
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("[pdf] 缺少 python-docx", file=sys.stderr)
    raise SystemExit(1)

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


def iter_block_items(doc):
    """按文档顺序迭代段落与表格。"""
    from docx.oxml.ns import qn
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def heading_level(p: Paragraph) -> int:
    name = (p.style.name or "").lower()
    if name.startswith("heading"):
        m = re.search(r"(\d+)", name)
        if m:
            return int(m.group(1))
    # 兼容中文章节标题样式（标题 1/标题 2）
    m = re.search(r"标题\s*(\d+)", p.style.name or "")
    return int(m.group(1)) if m else 0


def esc(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def table_html(t: Table) -> str:
    rows = []
    for row in t.rows:
        cells = []
        for cell in row.cells:
            cells.append("<td>" + esc(cell.text.strip()).replace("\n", "<br>") + "</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    if not rows:
        return ""
    # 表头进 thead：长表跨页时表头重复，不靠「整表不跨页」硬塞（那样会把超页表格裁掉）
    return ('<table class="data"><thead>' + rows[0] + "</thead><tbody>"
            + "".join(rows[1:]) + "</tbody></table>")


def _is_bold_caption(p: Paragraph) -> bool:
    """短句全加粗 → 表题/图题的候选（避免表题落在页底、表格翻页后成孤儿）。"""
    text = p.text.strip()
    if not text or len(text) > 40:
        return False
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)


def docx_to_html(doc):
    parts, headings = [], []
    cover, in_cover = [], True
    blocks = list(iter_block_items(doc))
    skip = set()
    for i, block in enumerate(blocks):
        if i in skip:
            continue
        if isinstance(block, Table):
            parts.append(table_html(block))
            in_cover = False
            continue
        p = block
        # 图片段落（w:drawing→a:blip）先处理：p.text 为空但内容必须进 PDF
        try:
            blips = p._p.xpath(".//a:blip/@r:embed")
        except Exception:  # noqa: BLE001
            blips = []
        if blips:
            for rid in blips:
                try:
                    part = doc.part.related_parts[rid]
                except KeyError:
                    continue
                b64 = base64.b64encode(part.blob).decode("ascii")
                parts.append(f'<p class="fig"><img src="data:{part.content_type};base64,{b64}" alt=""/></p>')
            in_cover = False
            continue
        text = p.text.strip()
        if not text:
            continue
        # docx 里的「目录」标题与 TOC 域由本脚本重新生成目录页，不能原样再排一遍
        if text.replace("　", "").replace(" ", "") == "目录":
            continue
        # 表题紧接表格：包成不可分页的块，表题不会单独留在页底
        if _is_bold_caption(p) and i + 1 < len(blocks) and isinstance(blocks[i + 1], Table):
            parts.append('<div class="keep"><p class="tcap">' + esc(text) + "</p>"
                         + table_html(blocks[i + 1]) + "</div>")
            skip.add(i + 1)
            in_cover = False
            continue
        lvl = heading_level(p)
        if 1 <= lvl <= 3:
            headings.append((lvl, text))
            parts.append(f'<h{lvl} data-h="{esc(text)}">{esc(text)}</h{lvl}>')
            in_cover = False
        else:
            parts.append("<p>" + esc(text) + "</p>")
            if in_cover:
                cover.append(parts.pop())
    return "\n".join(parts), headings, "\n".join(cover)


CSS = """
@page { size: A4; margin: 37mm 26mm 35mm 28mm; }
* { box-sizing: border-box; }
body { font-family: "仿宋", "FangSong", "FangSong_GB2312", "SimSun", serif;
       font-size: 16pt; line-height: 29pt; color: #000; margin: 0; }
h1 { font-family: "黑体", "SimHei", "Microsoft YaHei", sans-serif;
     font-size: 16pt; font-weight: 400; line-height: 29pt;
     margin: 0; page-break-before: always; }
h1:first-of-type { page-break-before: auto; }
h2 { font-family: "楷体", "KaiTi", "KaiTi_GB2312", serif;
     font-size: 16pt; font-weight: 400; line-height: 29pt; margin: 0; }
h3 { font-family: "仿宋", "FangSong", serif;
     font-size: 16pt; font-weight: 700; line-height: 29pt; margin: 0; }
p { margin: 0; text-align: justify; text-indent: 2em; line-height: 29pt; }
p.fig { text-align: center; text-indent: 0; margin: 6pt 0; page-break-inside: avoid; }
p.fig img { width: 100%; max-width: 150mm; }
p.toc-line { text-indent: 0; margin: 0; display: flex; line-height: 29pt;
             justify-content: space-between; }
p.toc-line.lv1 { font-family: "黑体", "SimHei", sans-serif; }
p.toc-line .pg { color: #000; }
.keep { page-break-inside: auto; }
p.tcap { text-indent: 0; font-weight: 700; margin: 6pt 0 3pt; page-break-after: avoid; }
table.data { width: 100%; border-collapse: collapse; margin: 6pt 0 6pt;
             font-family: "仿宋", "FangSong", serif;
             font-size: 10.5pt; line-height: 16pt; page-break-inside: auto; }
table.data thead { display: table-header-group; }
table.data tr { page-break-inside: avoid; }
table.data td { border: 0.5pt solid #000; padding: 2pt 4pt; vertical-align: top;
                word-break: break-word; overflow-wrap: anywhere; }
table.data tr:first-child td { background: #f2f2f2; font-weight: 700; }
.toc-page { page-break-after: always; }
.cover { page-break-after: always; text-align: center; padding-top: 55mm; }
.cover p { text-align: center; text-indent: 0; margin: 0 0 14pt; }
.cover p:first-child { font-family: "黑体", "SimHei", sans-serif;
                       font-size: 26pt; font-weight: 400; line-height: 1.5; margin-bottom: 22pt; }
.cover p:nth-child(2) { font-size: 16pt; color: #000; }
.toc-title { font-family: "黑体", "SimHei", sans-serif;
             font-size: 22pt; font-weight: 400; text-align: center; margin: 0 0 14pt;
             line-height: 29pt; }
"""


def build_html(body: str, toc_html: str, cover_html: str = "") -> str:
    cover = f'<div class="cover">{cover_html}</div>' if cover_html else ""
    return ("<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>"
            f"<title>现状分析报告</title><style>{CSS}</style></head><body>"
            f"{cover}{toc_html}{body}</body></html>")


def toc_block(headings, pages: dict[int, int], sentinel: bool = True) -> str:
    if not headings:
        return ""
    lines = ['<div class="toc-page"><div class="toc-title">目录</div>']
    for i, (lvl, text) in enumerate(headings):
        if lvl > 2:
            continue
        pg = pages.get(i, "")
        lines.append(f'<p class="toc-line lv{lvl}"><span>{esc(text)}</span>'
                     f'<span class="pg">{pg}</span></p>')
    # 目录结束哨兵：定位正文页码时跳过目录页，避免目录条目本身被当成标题命中。
    # 哨兵必须放在目录页容器「内部」——放到容器外会落到正文第一页上，
    # 使定位从第二页开始，首页的 H1/H2 全部查不到页码。
    # 定稿那一遍换成 HTML 注释——不占版面高度，也不留下可被复制到的文字。
    if sentinel:
        lines.append('<div style="font-size:1pt;color:#ffffff">CARDOTOCEND</div>')
    else:
        lines.append("<!--CARDOTOCEND-->")
    lines.append("</div>")
    return "".join(lines)


def find_browser():
    candidates = [
        os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                     "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"),
                     "Microsoft", "Edge", "Application", "msedge.exe"),
        shutil.which("msedge.exe"),
        shutil.which("chrome.exe"),
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"),
                     "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


def render(html: str, pdf: Path, work: Path) -> bool:
    browser = find_browser()
    if not browser:
        print("[pdf] 找不到 Edge/Chrome，无法打印 PDF", file=sys.stderr)
        return False
    html_path = work / "report.html"
    html_path.write_text(html, encoding="utf-8")
    cmd = [browser, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
           f"--print-to-pdf={pdf}", html_path.as_uri()]
    try:
        subprocess.run(cmd, check=True, timeout=180,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as exc:
        print(f"[pdf] 打印失败：{exc}", file=sys.stderr)
        return False
    return pdf.exists()


def heading_pages(pdf: Path, headings) -> dict[int, int]:
    """在 PDF 里逐页找标题文字，返回 {标题序号: 页码}（页码从 1 起）。

    按「行首」匹配而不是整页子串：正文里引用章名（如「该口径在「现状详细分析」一章引用…」）
    会让整页子串先命中引用句所在页，目录页码就跳到引用页而不是标题页。
    标题在版面上独占一行，所以用行首比对既能命中标题，也不会被正文引用误伤。
    """
    if fitz is None:
        return {}
    out = {}
    doc = fitz.open(str(pdf))
    norm = lambda s: re.sub(r"\s+", "", s or "")
    pages = [[norm(line) for line in (p.get_text("text") or "").splitlines()] for p in doc]
    texts = [norm(p.get_text("text")) for p in doc]
    start = 0
    for idx, t in enumerate(texts):
        if "CARDOTOCEND" in t:
            start = idx + 1
    for i, (_lvl, text) in enumerate(headings):
        key = norm(text)
        probe = key[:14]
        if not probe:
            continue
        for idx in range(start, len(pages)):
            if any(line.startswith(probe) for line in pages[idx]):
                out[i] = idx + 1
                break
    doc.close()
    return out


def bake_footer(pdf: Path, body_start: int = 2) -> None:
    """页脚：公文样式「— N —」+ 右侧「共 M 页」（烘进 PDF，任何阅读器一致）。

    正文页码从 1 起（封面、目录不计），与 docx 的 PAGE / SECTIONPAGES 双域同口径。
    """
    if fitz is None:
        return
    doc = fitz.open(str(pdf))
    total = doc.page_count
    body_total = max(0, total - (body_start - 1))
    for idx, page in enumerate(doc):
        if idx < body_start - 1:
            continue
        w, h = page.rect.width, page.rect.height
        num = idx - (body_start - 1) + 1
        label = f"— {num} —"
        page.insert_text((w / 2 - 14, h - 22), label,
                         fontsize=10.5, fontname="china-s", color=(0, 0, 0))
        tail = f"共 {body_total} 页"
        page.insert_text((w - 78, h - 22), tail,
                         fontsize=10.5, fontname="china-s", color=(0, 0, 0))
    doc.saveIncr()
    doc.close()


def patch_app_pages(docx: Path, total: int) -> None:
    """把 docProps/app.xml 的 Pages 改成实际页数，避免属性页显示旧值（13 页）。"""
    if total <= 0:
        return
    try:
        with zipfile.ZipFile(str(docx)) as z:
            names = z.namelist()
            if "docProps/app.xml" not in names:
                return
            parts = {n: z.read(n) for n in names}
    except Exception:
        return
    xml = parts["docProps/app.xml"].decode("utf-8")
    new, n = re.subn(r"<Pages>\d+</Pages>", f"<Pages>{total}</Pages>", xml, count=1)
    if n == 0 or new == xml:
        return
    parts["docProps/app.xml"] = new.encode("utf-8")
    tmp = str(docx) + f".tmp-{os.getpid()}"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
        os.replace(tmp, str(docx))
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--title", default="")
    ap.add_argument("--toc-json", default="",
                    help="把目录条目与正文页码写到这里（供 docx 目录缓存使用）")
    ap.add_argument("--body-start", type=int, default=0,
                    help="正文起始页（1 起）；0 表示自动取首个标题页")
    args = ap.parse_args()

    docx, out = Path(args.docx), Path(args.out)
    if not docx.exists():
        print(f"[pdf] docx 不存在：{docx}", file=sys.stderr)
        return 2
    out.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(str(docx))
    body, headings, cover = docx_to_html(doc)

    def rel_pages(pages: dict) -> tuple[dict, int]:
        """把绝对 PDF 页换成「正文第几页」（封面与目录不计），与定稿目录口径一致。"""
        if not pages:
            return {}, 0
        start = args.body_start or min(pages.values())
        return {i: p - start + 1 for i, p in pages.items()}, start

    with tempfile.TemporaryDirectory(prefix="cardo_pdf_") as tmp:
        work = Path(tmp)
        # 第一遍：无目录，定位正文页码
        probe = work / "probe.pdf"
        if not render(build_html(body, "", cover), probe, work):
            return 1
        pages, _ = rel_pages(heading_pages(probe, headings))
        # 第二遍：带目录（占位页码），目录会固定占用页数
        toc_pdf = work / "toc.pdf"
        toc = toc_block(headings, pages)
        if not render(build_html(body, toc, cover), toc_pdf, work):
            return 1
        # 目录插入后正文页码整体后移：以带目录的 PDF 重新定位并回填
        pages2, start = rel_pages(heading_pages(toc_pdf, headings))
        toc = toc_block(headings, pages2, sentinel=False)
        if not render(build_html(body, toc, cover), out, work):
            return 1

    total = 0
    if fitz is not None:
        d = fitz.open(str(out))
        total = d.page_count
        d.close()
        bake_footer(out, start)
        patch_app_pages(docx, total)
    if args.toc_json:
        entries = [{"level": lvl, "text": text, "page": pages2.get(i, "")}
                   for i, (lvl, text) in enumerate(headings) if lvl <= 2]
        try:
            Path(args.toc_json).write_text(
                json.dumps({"schema": "cardo/toc/v1",
                            "bodyPages": max(0, total - (start - 1)),
                            "entries": entries}, ensure_ascii=False, indent=1),
                encoding="utf-8")
        except OSError:
            pass
    print(f"[pdf] ok：{out.name}（{total} 页，正文 {max(0, total - (start - 1))} 页）"
          if total else f"[pdf] ok：{out.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
