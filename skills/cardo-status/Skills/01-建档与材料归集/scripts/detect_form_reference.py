# -*- coding: utf-8 -*-
"""形式参考件识别、隔离与串项黑名单生成。

判定不依赖固定文件名：docx、体量较大、段落数多、含本项目模板章节词、
且标题/开头不含本项目专名 → 判为“同类报告成品（形式参考）”。

隔离口径：
- 原件留在投递区不动（零写入），登记 excluded 路径，供 01 分类与 03 汇编跳过；
- 复制一份到 过程文件/00-总控/形式参考/参考件/ 存档，并记 sha256；
- 只抽取章节骨架、表格表头、依据排列方式等结构信息，不落正文段落；
- 生成“专有名词 + 数值”黑名单：先扣掉本项目材料里也出现的词/数，
  剩下的是该参考件独有的内容，供 09 对成品做零串入扫描。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

CHAPTER_HITS = ["项目概述", "现状分析", "现状", "建设必要性", "必要性", "编制依据",
                "更新思路", "建设方案", "投资估算", "效益分析", "风险", "结论"]
GENERIC = {"某市", "某省", "某区", "老旧小区", "老旧街区", "城市更新", "片区",
           "社区", "街道", "项目", "工程", "改造", "更新", "综合", "示范", "提升",
           "院落", "公寓", "花园", "广场", "公园", "小区", "建设", "实施", "推进"}
_BASE_PREFIX = ["某市", "某省", "某区",
                "关于", "实施", "进行", "推进", "开展", "位于", "涉及", "以及", "及", "和"]


def _load_prefixes():
    """通用前缀 + 可选的外部地名词表（一行一个，`#` 开头为注释）。

    词表位置：技能目录下 assets/地名词表.txt，或环境变量 CARDO_GAZETTEER 指定的文件。
    词表缺失时不做地名枚举，改由 _place_suffix_only 按“区/县/市”后缀兜底。
    """
    prefixes = list(_BASE_PREFIX)
    path = Path(__file__).resolve().parents[3] / "assets" / "地名词表.txt"
    env = os.environ.get("CARDO_GAZETTEER")
    if env:
        path = Path(env)
    if not path.is_file():
        return prefixes, False
    for line in path.read_text(encoding="utf-8").splitlines():
        name = line.strip()
        if name and not name.startswith("#"):
            prefixes.append(name)
    return prefixes, True


GENERIC_PREFIX, GAZETTEER_LOADED = _load_prefixes()
PLACE_SUFFIX = re.compile(r"^[\u4e00-\u9fa5]{2,6}(区|县|市)$")


def _place_suffix_only(tok: str) -> bool:
    """词表缺失时的兜底：以行政区后缀结尾的短词按地名处理。"""
    return (not GAZETTEER_LOADED) and bool(PLACE_SUFFIX.fullmatch(tok))


def sha256(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def docx_text(p: Path) -> str:
    try:
        from docx import Document
        d = Document(p)
        parts = [x.text for x in d.paragraphs if x.text and x.text.strip()]
        for t in d.tables:
            for row in t.rows:
                for c in row.cells:
                    if c.text and c.text.strip():
                        parts.append(c.text)
        return "\n".join(parts)
    except Exception:
        return ""


def pptx_text(p: Path) -> str:
    try:
        with zipfile.ZipFile(p) as z:
            names = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                           key=lambda s: int(re.search(r"(\d+)", s).group(1)))
            out = []
            for n in names:
                x = z.read(n).decode("utf-8", errors="ignore")
                out.extend(re.findall(r"<a:t>(.*?)</a:t>", x))
            return "\n".join(out)
    except Exception:
        return ""


def xlsx_text(p: Path) -> str:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        out = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                out.extend(str(v) for v in row if v is not None)
        return "\n".join(out)
    except Exception:
        return ""


def read_text(p: Path, limit: int = 3_000_000) -> str:
    suf = p.suffix.lower()
    if suf == ".docx":
        t = docx_text(p)
    elif suf == ".pptx":
        t = pptx_text(p)
    elif suf in (".xlsx", ".xlsm"):
        t = xlsx_text(p)
    else:
        return ""
    return t[:limit]


def outline_of(p: Path) -> dict:
    """只取结构：标题、H1/H2/H3、表格表头、依据排列样式。"""
    try:
        from docx import Document
        d = Document(p)
    except Exception:
        return {}
    paras = [x for x in d.paragraphs if x.text and x.text.strip()]
    headings = []
    for x in paras:
        s = x.text.strip()
        st = (x.style.name if x.style else "") or ""
        is_head = ("Heading" in st or "标题" in st or re.match(r"^[一二三四五六七八九十]+、", s)
                   or re.match(r"^（[一二三四五六七八九十]+）", s)
                   or re.match(r"^\d+(\.\d+){0,2}\s+\S", s))
        if is_head and len(s) <= 80:
            headings.append({"style": st, "text": s[:80]})
        if len(headings) >= 120:
            break
    tables = []
    for t in d.tables[:8]:
        try:
            header = [c.text.strip().replace("\n", " ")[:24] for c in t.rows[0].cells]
        except Exception:
            header = []
        tables.append({"rows": len(t.rows), "cols": len(t.columns), "header": header})
    basis = []
    for x in paras:
        s = x.text.strip()
        if "依据" in s and len(s) <= 120:
            basis.append({"style": (x.style.name if x.style else ""),
                          "pattern": re.sub(r"[《》].*?[》]", "《…》", s)[:80]})
        if len(basis) >= 20:
            break
    title = next((x.text.strip() for x in paras[:5] if 6 <= len(x.text.strip()) <= 90), "")
    return {"title": title, "headings": headings, "tables": tables, "basis_samples": basis,
            "paragraphs": len(paras)}


def normalize_num(s: str) -> str:
    return s.replace("％", "%").replace(" ", "").replace(",", "")


# 地名后缀：两类（多字后缀优先匹配，单字后缀兜底）
PLACE_SUFFIX_RE = re.compile(r"(街道|社区|小区|公寓|花园|大厦|广场|公园|巷|路|街|里|村|坊|苑|桥)")
# 判定“专名”时剔除的动词/泛化词：命中即说明该片段是句子而不是地名
VERB_WORDS = (
    "做好", "促进", "保持", "提升", "改善", "解决", "加强", "推动", "实施", "建设",
    "改造", "更新", "完成", "进行", "组织", "统筹", "覆盖", "位于", "属于", "涉及",
    "包括", "主要", "部分", "最终", "控制", "拆除", "打造", "形成", "实现", "满足",
    "达到", "提供", "保障", "维护", "管理", "服务", "完善", "优化", "整合", "盘活",
    "利用", "开发", "运营", "发展", "增长", "增加", "减少", "降低", "提高", "增强",
    "拓宽", "畅通", "整治", "修缮", "加固", "消除", "清理", "规范", "统一", "协调",
    "配合", "支持", "引导", "鼓励", "吸引", "凝聚", "体现", "展示", "呈现", "反映",
    "包含", "具有", "具备", "成为", "作为", "用于", "用来", "便于", "利于", "不断",
    "持续", "积极", "切实", "全面", "加快", "进一步", "同步", "分别", "共同", "重点",
    "多数", "少数", "大量", "启动", "开展", "推进", "一步", "纳入", "列入", "提出",
    "明确", "确定", "落实", "执行", "安排", "计划", "预计", "以及", "同时", "根据",
    "按照", "结合", "针对", "对于", "关于", "通过", "由此", "因此", "其中", "目前",
    "当前", "现已", "已经", "正在", "将要", "即将", "能够", "可以", "需要", "应当",
    "必须", "不得", "不应", "是否", "相关", "上述", "现有", "原有", "既有", "本次",
    "此次", "面积", "户数", "人口", "项目", "工程", "片区", "街区", "小区内", "小区外",
    "老旧", "城镇", "城市", "居住", "住宅", "成熟", "共享", "依托", "指导", "涉农",
    "农业", "处置", "施工", "排查", "用电", "前期", "后期", "同期", "后续", "同类",
    "工作", "道路", "两个", "每个", "涵盖", "特色", "线路", "疏散", "确保", "设计",
    "负责", "遗撒", "污染", "背街", "全区", "市政",
)
# 文档/标准名称特征：命中说明这是可共享的引用名，不是某项目的专名
DOC_WORDS = ("法", "条例", "规定", "办法", "标准", "规范", "导则", "规程", "准则",
             "规划", "意见", "方案", "通知", "定额", "纲要", "计划", "规则", "政策",
             "文件", "指南", "技术要求", "批复", "报告", "合同", "协议", "章程", "细则")
# 片段里出现这些虚词/代词，基本可判定不是地名
BAD_CHARS = re.compile(r"[的了着过们为该各等将对把被让使并且或则其是在于与及从按由]")
BAD_START = set("省市县区镇乡村街路巷两个每各多少全本该其此同前后以从在于个内年和")
# 正文取词时的“左边界”：前缀之前必须是标点/空白/数字或这些虚词，才允许起词
BOUNDARY = set("\n\r\t 　，。、；：？！“”‘’（）()《》【】[]—…·-~/\\0123456789")
BOUNDARY |= set("的在于为是与和及至从到由被让把对按经沿临接邻靠省市区县镇乡村")


def _is_doc_name(tok: str) -> bool:
    return any(w in tok for w in DOC_WORDS)


def _clean_place(tok: str) -> bool:
    """判断一个候选片段是不是“像地名的专名”。"""
    tok = tok.strip()
    if not (3 <= len(tok) <= 12):
        return False
    if tok in GENERIC or _is_doc_name(tok):
        return False
    if any(w in tok for w in VERB_WORDS):
        return False
    if BAD_CHARS.search(tok):
        return False
    if tok[0] in BAD_START:
        return False
    if any(tok.startswith(p) and len(tok) > len(p) + 1 for p in GENERIC_PREFIX):
        return False
    if _place_suffix_only(tok):
        return False
    return True


def extract_places(text: str, structural: bool = False) -> set[str]:
    """按地名后缀取候选专名。

    structural=True 时用于标题/表头/表格单元格，取词不看左边界（整格多为地名）；
    structural=False 时用于正文，必须命中左边界，避免把“覆盖示例街”这类
    动词短语当成地名。
    """
    out: set[str] = set()
    for m in PLACE_SUFFIX_RE.finditer(text):
        suffix, end = m.group(1), m.end()
        for plen in range(4, 1, -1):
            start = end - len(suffix) - plen
            if start < 0:
                continue
            tok = text[start:end]
            if not all("\u4e00" <= c <= "\u9fff" for c in tok):
                continue
            if not structural:
                pre = text[start - 1] if start > 0 else "\n"
                if pre not in BOUNDARY:
                    continue
            if _clean_place(tok):
                out.add(tok)
    return out


def table_cells_of(p: Path, max_tables: int = 12, max_cols: int = 3) -> list[str]:
    """取表格前几列的单元格文本（小区/街道清单里的地名主要在这里）。"""
    try:
        from docx import Document
        d = Document(p)
    except Exception:
        return []
    out: list[str] = []
    for t in d.tables[:max_tables]:
        for row in t.rows:
            for c in row.cells[:max_cols]:
                s = (c.text or "").strip().replace("\n", " ")
                if s and len(s) <= 40:
                    out.append(s)
    return out


def build_blacklist(ref_texts: list[str], project_text: str, current_project: str,
                    ref_titles: list[str] | None = None,
                    ref_structures: list[dict] | None = None) -> dict:
    proj_norm = re.sub(r"\s+", "", project_text)
    cur_norm = re.sub(r"\s+", "", current_project)
    proper = set(ref_titles or [])
    for text in ref_texts:
        for m in re.finditer(r"《([^》]{4,80})》", text):
            name = m.group(1).strip()
            if not _is_doc_name(name):     # 法规/标准名可共享，不算专名
                proper.add(name)
        proper |= extract_places(text, structural=False)
    for st in ref_structures or []:
        for h in st.get("headings", []):
            proper |= extract_places(h.get("text", ""), structural=True)
        for c in st.get("cells", []):
            proper |= extract_places(c, structural=True)
    # 去掉本项目材料里已出现的、本项目名内的、以及纯通用词
    proper = {t for t in proper if t and t not in GENERIC and len(t) >= 3
              and t not in proj_norm and t not in cur_norm}
    proper = {t for t in proper if _clean_place(t) or t in (ref_titles or [])}

    nums = set()
    keep_min = {"平方米": 100, "㎡": 100, "公顷": 0.1, "亩": 10, "万元": 10, "亿元": 0,
                "户": 50, "栋": 3, "人": 100, "个": 10, "处": 2, "米": 10,
                "公里": 0.1, "千米": 0.1}
    for text in ref_texts:
        for m in re.finditer(
                r"(\d+(?:\.\d+)?)\s*(平方米|㎡|公顷|亩|万元|亿元|户|栋|人|个|处|米|公里|千米|%|％)", text):
            val, unit = m.group(1), normalize_num(m.group(2))
            try:
                v = float(val)
            except ValueError:
                continue
            if unit == "%" or (1900 <= v <= 2035):
                continue
            if v < keep_min.get(unit, 0):
                continue
            tok = normalize_num(val + unit)
            if val in proj_norm:
                continue
            nums.add(tok)
    nums = {n for n in nums if not any(n != m and n in m for m in nums)}
    return {"proper_nouns": sorted(proper), "numbers": sorted(nums)}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--delivery", required=True, help="投递区目录（相对项目根或绝对路径）")
    ap.add_argument("--project", required=True)
    ap.add_argument("--out", required=True, help="形式参考登记输出目录")
    args = ap.parse_args()

    ws = Path(args.project_dir).resolve()
    delivery = Path(args.delivery)
    if not delivery.is_absolute():
        delivery = ws / delivery
    out_dir = Path(args.out)
    if not out_dir.is_absolute():
        out_dir = ws / out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    # 只认投递区顶层的独立 docx：上游成果包内的技术说明不是“同类报告成品”，
    # 也不把成套文件夹里的附件误判成形式参考。
    files = [p for p in sorted(delivery.iterdir())
             if p.is_file() and p.suffix.lower() == ".docx" and not p.name.startswith("~$")
             and p.stat().st_size >= 2 * 1024 * 1024]
    cur_tokens = [x for x in re.split(r"片区|项目|工程|更新|提升|老旧|街区|建议书|改造|综合|建设", args.project)
                  if len(x.strip()) >= 2]

    candidates = []
    for p in files:
        try:
            size = p.stat().st_size
            text = read_text(p, 3000_000)
            paras = text.count("\n") + 1
            head = text[:2500]
            same = any(tok and tok in head for tok in cur_tokens)
            hits = [k for k in CHAPTER_HITS if k in text]
            score = (1 if size >= 2 * 1024 * 1024 else 0) + (1 if paras >= 150 else 0) + \
                    (1 if len(set(hits)) >= 3 else 0) + (1 if not same else 0)
            if (not same) and len(set(hits)) >= 3 and score >= 3:
                candidates.append({"path": p, "size": size, "text": text, "score": score,
                                   "hits": sorted(set(hits))})
        except Exception:
            continue

    # 本项目材料文本（用于扣掉共享词/共享数）：投递区内所有可读文本件，
    # 排除识别出的参考件。上游成果目录里的指标说明、台账等也要算进来，
    # 否则别的项目与本项目共有的人数、栋数会被误当“参考件独有数值”。
    project_text_parts = []
    cand_paths = {c["path"] for c in candidates}
    for p in sorted(delivery.rglob("*")):
        if not p.is_file() or p.name.startswith("~$") or p in cand_paths:
            continue
        if p.suffix.lower() not in (".docx", ".pptx", ".xlsx", ".xlsm"):
            continue
        try:
            if p.stat().st_size > 40 * 1024 * 1024:
                continue
        except OSError:
            continue
        t = read_text(p, 300_000)
        if t:
            project_text_parts.append(t)
    project_text = "\n".join(project_text_parts)[:6_000_000]

    ref_texts = [c["text"] for c in candidates]
    ref_titles, ref_structures = [], []
    for c in candidates:
        st = outline_of(c["path"])
        st["cells"] = table_cells_of(c["path"])
        c["structure"] = st
        if st.get("title"):
            ref_titles.append(st["title"])
        ref_structures.append(st)
    black = build_blacklist(ref_texts, project_text, args.project, ref_titles, ref_structures) \
        if candidates else {"proper_nouns": [], "numbers": []}

    ref_dir = out_dir / "参考件"
    ref_dir.mkdir(parents=True, exist_ok=True)
    refs = []
    for c in candidates:
        p = c["path"]
        copy = ref_dir / p.name
        if not copy.exists():
            shutil.copy2(p, copy)
        refs.append({
            "path": str(p),
            "rel": p.relative_to(delivery).as_posix(),
            "copy": str(copy),
            "sha256": sha256(p),
            "size": c["size"],
            "score": c["score"],
            "chapter_hits": c["hits"],
            "structure": c["structure"],
        })

    data = {
        "schema": "cardo/form-reference/v1",
        "detected_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "current_project": args.project,
        "delivery": str(delivery),
        "reference": refs,
        "excluded": [r["rel"] for r in refs],
        "blacklist": black,
        "note": "形式参考件只学章节骨架、句式、表格样式和依据排列方式；"
                "专有名词与数值不得进入正文。excluded 由 01 分类与 03 汇编跳过。",
    }
    (out_dir / "形式参考.json").write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                            encoding="utf-8")

    md = ["# 形式参考件结构登记", "",
          f"- 识别时间：{data['detected_at']}",
          f"- 本项目：{args.project}",
          f"- 命中件数：{len(refs)}",
          "- 用途：只参考章节骨架、句式、表格样式、依据排列方式；正文零引用。", ""]
    for r in refs:
        st = r["structure"]
        md += [f"## {st.get('title') or Path(r['rel']).stem}", "",
               f"- 文件：{r['rel']}",
               f"- sha256：{r['sha256']}",
               f"- 段落数：{st.get('paragraphs')}；评分：{r['score']}",
               f"- 章节命中：{'、'.join(r['chapter_hits'])}", "",
               "### 章节骨架（只记标题）", ""]
        for h in st.get("headings", [])[:60]:
            md.append(f"- {h['text']}")
        md += ["", "### 表格样式（只记表头）", ""]
        for t in st.get("tables", [])[:6]:
            md.append(f"- {t['rows']} 行 × {t['cols']} 列：{' / '.join(t['header'])}")
        md += ["", "### 依据排列样式（不记正文）", ""]
        for b in st.get("basis_samples", [])[:12]:
            md.append(f"- {b['pattern']}")
        md.append("")
    (out_dir / "形式参考.md").write_text("\n".join(md) + "\n", encoding="utf-8")
    print(json.dumps({"references": len(refs), "proper_nouns": len(black["proper_nouns"]),
                      "numbers": len(black["numbers"])}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
