#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""extract_facts.py -- 现状事实底座薄抽取（T22 v0：从原文转储抽候选事实 + 来源定位）。

用法：
    python extract_facts.py --project-dir <项目目录> [--out <目录>] [--max-facts 400] [--no-timestamp]
    # 校准/回归（标准项目库布局）：--root <项目库根> --project <项目名> [--direction <方向>]

做：读 `<项目>/05_过程/原文转储/`（blocks.txt / tables.txt）+ `03_拆解成果/章节映射.yaml`，按「块号/表行号」
抽带单位数值与少量关键结论句，写 `05_过程/现状事实库/`：
  事实库.json（cardo/facts/v1：facts / conflicts / missing / rejected + stats）
  事实库.md（人读：按 tNN 分组 + 覆盖率 + 冲突/缺失）
  未采纳清单.md（未采纳候选及原因）
不做：不读 docx 正文、不做 OCR、不推断、不补数；冲突并列标 conflict（不静默选一个）。
幂等：同输入同参数逐字节一致（--no-timestamp 去掉 generatedAt）；只写 --out。
退出码：0 抽到事实；1 抽取/写入失败（0 条事实）；2 参数/输入坏（项目目录、转储、输出）。
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _contract_lib import norm_title, parse_yaml, read_text  # noqa: E402

__version__ = "0.1.0"
SCHEMA = "cardo/facts/v1"
DUMP_DIR = Path("05_过程") / "原文转储"
OUT_DIR = Path("05_过程") / "现状事实库"
PDF_TEXT_DIR = Path("02_分类材料") / "11_PDF文本"
PAGE_RE = re.compile(r"^\[PAGE\s+(\d+)\]$")
MAPPING = Path("03_拆解成果") / "章节映射.yaml"
PROJECT_MARKERS = ("00_项目索引.md", ".cardo")
UNIT_ALT = ("万平方米|万平方公里|万立方米|平方公里|平方米|立方米|千伏安|万元|亿元|万人|万吨|"
            "公里|千米|公顷|毫米|厘米|小时|分钟|分贝|米|亩|吨|个|户|人|处|栋|座|条|项|辆|"
            "千伏|千瓦|万|亿|%|个百分点|℃|度|天|秒")
NUM_RE = re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)\s*(" + UNIT_ALT + ")")
# 表格单元格里的数字**不带单位**——单位在列名里（"户数（户）"、"建筑面积（万㎡）"）。
# 正文用的 NUM_RE 要求数字后跟单位，对表格单元格一律抽不中：
# 实测楼栋台账整张表（户数/楼栋数/面积/建成年代几十个值）一条事实都抽不出来。
# 故表格走这个判据：整格就是一个数字（可带单位），单位另从列名取。
CELL_NUM_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*(.*)$")
# 纯数值格（用于识别表头续行）：整格就是一个数字，没有别的字
PURE_NUM_RE = re.compile(r"^\d+(?:\.\d+)?$")
# 列名里的单位：取括号内的部分（"户数（户）" → 户，"建筑面积（万㎡）" → 万㎡）
COL_UNIT_RE = re.compile(r"[（(]([^）)]+)[）)]")
STYLE_RE = re.compile(r"^\[([^\]]*)\]\s*(.*)$")
TABLE_RE = re.compile(r"^=+\s*TABLE\s+(\d+)\s*\((\d+)x(\d+)\)")
ROW_RE = re.compile(r"^\s*R(\d+):\s*(.*)$")
KEY_SENTENCE = re.compile(r"(据统计|现状|共计|合计|约|总体|主要|存在|达到|占比)")

class BadInput(Exception):
    """参数/输入错误 → exit 2。"""

def resolve_project_dir(args) -> Path:
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
        candidates: list[Path] = []
        if str(args.direction or "").strip():
            candidates.append(root / str(args.direction).strip() / project)
        else:
            candidates += sorted(d / project for d in root.iterdir() if d.is_dir() and (d / project).is_dir())
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

def doc_version(name: str) -> str:
    rev, date = re.search(r"修订版-(\d+)", name), re.search(r"(20\d{6})", name)
    if rev and date:
        return f"修订版-{rev.group(1)}-{date.group(1)}"
    return f"v-{date.group(1)}" if date else "v1"

def template_sections(project: Path) -> list[dict]:
    """章节映射.yaml → [{section: tNN, heading: 模板标题, sources: [源报告标题]}]（映射是唯一真相）。"""
    path = project / MAPPING
    if not path.is_file():
        return []
    try:
        data = parse_yaml(read_text(path))
    except ValueError:
        return []
    out = []
    for i, node in enumerate(data.get("nodes") or data.get("sections") or []):
        if not isinstance(node, dict):
            continue
        heading = str(node.get("template_heading") or node.get("title") or "").strip()
        section = str(node.get("id") or f"t{i + 1:02d}").strip()
        sources = [str(s).strip() for s in (node.get("sources") or []) if str(s).strip()]
        if heading and section:
            out.append({"section": section, "heading": heading, "sources": sources})
    return out

def pick_sources(project: Path) -> tuple[list[Path], list[Path]]:
    dump = project / DUMP_DIR
    if not dump.is_dir():
        raise BadInput(f"找不到原文转储目录：{DUMP_DIR.as_posix()}（先跑 dump_docx_outline.py）")
    blocks = sorted(p for p in dump.glob("*.blocks.txt") if "模板" not in p.name)
    tables = sorted(p for p in dump.glob("*.tables.txt") if "模板" not in p.name)
    if not blocks and not tables:
        raise BadInput(f"原文转储目录下没有 blocks.txt/tables.txt：{DUMP_DIR.as_posix()}")
    return blocks, tables

def find_sections(text: str, sections: list[dict]) -> list[str]:
    """标题 → 可能归属的 tNN 列表（同一源标题可被多个模板节引用，按映射顺序返回）。"""
    norm = norm_title(text)
    if not norm:
        return []
    exact = [sec["section"] for sec in sections if norm == norm_title(sec["heading"])]
    if exact:
        return exact[:1]
    hits: list[str] = []
    for sec in sections:
        for src in sec.get("sources") or []:
            nsrc = norm_title(src)
            if nsrc and (norm == nsrc or nsrc in norm or norm in nsrc):
                hits.append(sec["section"])
                break
    return hits

# 表头语义 → 模板节标题的对照：现状分析真正要用的数据多在表里
# （"现状基本情况"表逐行给出各小区的建成年代、户数、层数、建筑面积、产权、结构），
# 而原先表格行一律不归属任何节、只进补充池末尾，结果每节拿到的表格行是错位的，
# 甚至整节无数据。
# 判据取「表头列名」与「节标题」的语义交集；命中即归属，不再靠顺序碰运气。
TABLE_HEADER_SECTIONS = [
    (("建成年代", "层数", "单元数", "建筑面积", "产权单位", "建筑结构", "户数", "小区"),
     ["建筑现状与风貌保护现状", "土地与空间利用现状", "人口与社会人居现状"]),
    (("用地", "用地面积", "红线", "四至", "用地性质"),
     ["土地与空间利用现状"]),
    (("常住", "人口", "户籍", "年龄", "老年人"),
     ["人口与社会人居现状"]),
    (("商户", "业态", "经营", "商铺", "租金", "空置"),
     ["街区业态与经营现状"]),
    (("停车", "车位", "道路", "交通", "通行"),
     ["交通与通行现状"]),
    (("管线", "管网", "给水", "排水", "燃气", "电力", "消防", "照明"),
     ["市政基础设施现状"]),
    (("配套", "服务", "教育", "医疗", "商业", "养老"),
     ["公共服务配套现状"]),
    (("绿化", "绿地", "景观", "树木", "环境", "口袋"),
     ["生态与人居环境现状"]),
    (("投资", "费用", "金额", "造价", "资金", "概算"),
     ["项目概况"]),
    (("文件", "资料", "名称", "清单"),
     ["项目编制依据"]),
]


def table_sections(header: list[str], sections: list[dict]) -> list[str]:
    """按表头列名判断这张表的内容属于哪个模板节。

    返回该表可能归属的节 id 列表（按表头语义对照表的顺序，节 id 去重）。
    判不出则返回空列表，交由既有的补充池逻辑处理。
    """
    if not header:
        return []
    heads = " ".join(str(h) for h in header)
    out: list[str] = []
    for keys, want_titles in TABLE_HEADER_SECTIONS:
        if not any(k in heads for k in keys):
            continue
        for sec in sections:
            title = norm_title(sec.get("heading") or "")
            if not title:
                continue
            if any(norm_title(w) == title for w in want_titles):
                if sec["section"] not in out:
                    out.append(sec["section"])
        if out:
            break
    return out


def metric_phrase(quote: str) -> str:
    """去掉数值与单位后的句式（冲突判定用：只有同一句式不同数值才可能算冲突）。"""
    text = str(quote)
    for unit in sorted(UNIT_ALT.split("|"), key=len, reverse=True):
        text = text.replace(unit, "")
    return norm_title(re.sub(r"\d+(?:\.\d+)?", "", text))[:40]


def _unit_from_header(head: str) -> str:
    """从列名里取单位：取括号内的部分（"户数（户）" → 户，"建筑面积（万㎡）" → 万㎡）。

    取不到就返回空串——裸数字不入库：没有单位的事实在报告里没法用，
    写进正文只会变成一堆来路不明的数字。
    """
    if not head:
        return ""
    m = COL_UNIT_RE.search(head)
    if not m:
        return ""
    unit = m.group(1).strip()
    # 括号里不是单位（如"（最高）"、"（先不填）"）→ 丢弃
    if not unit or len(unit) > 6:
        return ""
    if any(bad in unit for bad in ("先不填", "最高", "序号", "填")):
        return ""
    return unit

def _row_context(cells: list[str], header: list[str], col: int) -> str:
    """取同一行里的**文本标识**（小区名/地址/社区），作为这个数的归属。

    没有它，事实库只有"122 户"这样的裸数，成文时说不出是哪个小区。
    实测楼栋台账的标识列是「小区名称」「具体地址」「所在社区」——
    取该行里最靠前的一个有名字的列。
    """
    if col < 0 or col >= len(cells):
        return ""
    # 标识列：**越具体越优先**（小区名称 > 具体地址 > 社区 > 街道）。
    # 实测楼栋台账里「街道」列排在「小区名称」前面，
    # 按列序取会拿到"某街道"这种粗粒度归属，说不清是哪个小区。
    # 每列只归入**最具体**的那一类（先匹配到的关键字胜出）。
    col_class: dict[int, str] = {}
    for kw in ("小区名称", "名称", "项目名称", "具体地址", "地址", "社区", "街道"):
        for i, h in enumerate(header):
            if not h or i in col_class:
                continue
            if kw in h:
                col_class[i] = kw
    for kw in ("小区名称", "名称", "项目名称", "具体地址", "地址", "社区", "街道"):
        for i, cls in col_class.items():
            if cls != kw:
                continue
            if i < len(cells) and cells[i] and not PURE_NUM_RE.match(cells[i]):
                return cells[i][:60]
    # 没有标识列 → 取该行**本列之前**最靠前的文本格（工程量表：小区名在左、工程量在右）
    for i in range(min(col, len(cells))):
        c = cells[i]
        if c and not PURE_NUM_RE.match(c) and not re.match(r"^\d{4}年?$", c):
            return c[:60]
    # 再兜底：整行第一个非数值文本格
    for c in cells:
        if c and not PURE_NUM_RE.match(c) and not re.match(r"^\d{4}年?$", c):
            return c[:60]
    return ""


def add_candidate(bucket: list, *, file: str, locator: str, sections: list[str], label: str,
                  quote: str, value, unit: str, confidence: str, order: tuple,
                  doc_version: str = "v1", context: str = "",
                  origin: str = "", origin_locator: str = "") -> None:
    # origin / origin_locator：这条事实在**原始材料**里的出处（文件名 + 表号），
    # 供 06 成文时写锚点用。汇编文件（素材汇编.blocks.txt）是中间产物，
    # 09 依据校验按文件名在工作区找不到它，也无法在转储里定位。
    src = {"file": file, "locator": locator, "quote": quote, "docVersion": doc_version}
    if origin:
        src["origin"] = origin
    if origin_locator:
        src["originLocator"] = origin_locator
    bucket.append({"source": src,
                   "section": sections[0] if sections else "", "sections": list(sections),
                   "label": label, "value": value, "unit": unit,
                   "confidence": confidence, "_order": order, "context": context})

# 汇编产出的块带来源前缀：「【来源：xx.xlsx】表（现状基本情况）数据行：…」
# 故正则允许前缀存在，并把前缀捕获出来做表头缓存的键——
# 不同文件的工作表都叫 "Sheet1"，只按表名索引会让后一个文件的表
# 继承前一个文件的列名，抽出来的事实就张冠李戴了。
SRC_PREFIX_RE = re.compile(r"^(?:【([^】]*)】)?")
ROW_BLOCK_RE = re.compile(r"^(?:【[^】]*】)?表（([^）]*)）数据行：(.*)$")
HEAD_BLOCK_RE = re.compile(r"^(?:【[^】]*】)?表（([^）]*)）表头：(.*)$")


def clean_origin(prefix: str) -> str:
    """来源前缀只留原文件名：去掉「来源：/出处：」等标记。"""
    return re.sub(r"^(?:来源|出处)\s*[：:]\s*", "", (prefix or "").strip()).strip()


def harvest_blocks(path: Path, sections: list[dict], bucket: list) -> int:
    rel = (Path(DUMP_DIR) / path.name).as_posix()
    current: list[str] = []
    scanned = 0
    # 表格表头缓存：汇编产物里「表（X）表头：…」在「表（X）数据行：…」之前，
    # 记住它才能把数据行的第 N 格对上「第 N 列名」。
    tbl_headers: dict[str, list[str]] = {}
    # 刚读到表头块、正在等「表头续行」的那张表
    tbl_pending_header = None
    for lineno, line in enumerate(read_text(path).splitlines(), start=1):
        match = STYLE_RE.match(line)
        style, text = (match.group(1), match.group(2)) if match else ("", line)
        low = style.strip().lower()
        if low.startswith("toc") or "目录" in low:
            continue
        if "heading" in low or low in ("title", "标题"):
            hits = find_sections(text, sections)
            if hits:
                current = hits
            continue
        if not text.strip():
            continue
        scanned += 1
        stripped = text.strip()
        # 表头块：记下来，供后续数据行按列名对齐
        hm = HEAD_BLOCK_RE.match(stripped)
        if hm:
            # 键 = 来源前缀 + 表名：不同文件的工作表可能都叫 "Sheet1"
            srcm = SRC_PREFIX_RE.match(stripped)
            prefix = clean_origin((srcm.group(1) or "") if srcm else "")
            tbl_headers["{}::{}".format(prefix, hm.group(1))] = [
                c.strip() for c in hm.group(2).split("|")]
            # 有些表的**列名在第二行**（第一行是大标题，第二行才是"序号|街道|户数…"）。
            # 实测楼栋台账正是如此：列名不对齐，户数/面积就全抽不出来。
            # 判据：**格数多得多的那行才是列名行**（大标题只有 2 格，列名有 22 格）。
            # 不用"是不是数值"来判断——列名行里常有"1 立面材料修缮或翻新"这类带序号的说明格。
            tbl_pending_header = "{}::{}".format(prefix, hm.group(1))
            continue
        # 表格数据行：按表头逐格抽值，保留「列名｜值」的对应关系
        rm = ROW_BLOCK_RE.match(stripped)
        if rm:
            srcm = SRC_PREFIX_RE.match(stripped)
            prefix = clean_origin((srcm.group(1) or "") if srcm else "")
            key = "{}::{}".format(prefix, rm.group(1))
            header = tbl_headers.get(key, [])
            cells = [c.strip() for c in rm.group(2).split("|")]
            # 表头续行：格数明显多于当前表头 → 这才是真正的列名行，整体替换
            if (tbl_pending_header == key and header
                    and len(cells) > len(header) * 2):
                tbl_headers[key] = cells
                tbl_pending_header = None
                continue
            tbl_pending_header = None
            # 同一个 sheet 里常并着**两张列结构不同的表**（楼栋基本信息表 + 工程量表）。
            # 实测"楼栋数｜0.315"这种错位就是这么来的：拿第一张表的列名去读第二张表。
            # 判据：格数与列名数对不上 → 这是新表的第一行，它自己就是列名行。
            if header and len(cells) != len(header):
                new_header = list(cells)
                # 与已有列名合并（前缀相同的那部分沿用旧名）
                old = tbl_headers.get(key, [])
                for i, c in enumerate(new_header):
                    if i < len(old) and old[i] and (not c or PURE_NUM_RE.match(c)):
                        new_header[i] = old[i]
                tbl_headers[key] = new_header
                header = new_header
                # 若这一行本身就全是数值，说明它其实是数据行，只是列数与表头不符
                # （缺列/多列）→ 此时不能拿它当表头，退回旧表头并按位置对齐。
                if any(PURE_NUM_RE.match(c) for c in cells if c) and \
                        sum(1 for c in cells if PURE_NUM_RE.match(c)) >= len(cells) // 2:
                    tbl_headers[key] = old
                    header = old
            for col, cell in enumerate(cells):
                head = header[col] if col < len(header) else ""
                # 列名本身是序号位/说明位 → 不是事实，跳过
                if head and any(bad in head for bad in ("序号", "先不填", "改造项目及序号")):
                    continue
                # 整格是数字（或数字+单位）→ 直接抽；单位优先取格内自带，其次取列名括号里的
                cm = CELL_NUM_RE.match(cell)
                if cm:
                    value = float(cm.group(1))
                    tail = cm.group(2).strip()
                    unit = tail if tail else _unit_from_header(head)
                    if not unit:
                        continue          # 没有单位的裸数字不入库（来源里意义不明）
                    # 建成年代这类"年"不是现状指标，抽了只会污染事实库
                    if unit in ("年", "年代"):
                        continue
                    add_candidate(bucket, file=rel, locator=f"block {lineno}",
                                  sections=current,
                                  label=(f"{head}｜{cell}" if head else cell)[:40],
                                  quote=cell[:160], value=value, unit=unit,
                                  confidence="high",
                                  order=(current[0] if current else "zz", lineno, cm.group(1), unit),
                                  doc_version=doc_version(path.name),
                                  context=_row_context(cells, header, col),
                                  origin=prefix, origin_locator=f"表（{rm.group(1)}）")
                    continue
                # 格里是带单位的文本（"36户"、"0.252万㎡"）→ 走正文判据
                for num, unit in NUM_RE.findall(cell):
                    add_candidate(bucket, file=rel, locator=f"block {lineno}",
                                  sections=current,
                                  label=(f"{head}｜{cell}" if head else cell)[:40],
                                  quote=cell[:160], value=float(num), unit=unit,
                                  confidence="high",
                                  order=(current[0] if current else "zz", lineno, num, unit),
                                  doc_version=doc_version(path.name),
                                  context=_row_context(cells, header, col),
                                  origin=prefix, origin_locator=f"表（{rm.group(1)}）")
            continue
        numbers = NUM_RE.findall(text)
        # 汇编块常带「【来源：原文件】」前缀：原文出处用它，别把中间产物写进依据块
        srcm = SRC_PREFIX_RE.match(stripped)
        origin = clean_origin((srcm.group(1) or "") if srcm else "")
        for num, unit in numbers:
            add_candidate(bucket, file=rel, locator=f"block {lineno}", sections=current,
                          label=text.strip()[:40], quote=text.strip()[:160],
                          value=float(num), unit=unit, confidence="high",
                          order=(current[0] if current else "zz", lineno, num, unit),
                          doc_version=doc_version(path.name), origin=origin)
        if numbers and KEY_SENTENCE.search(text) and 8 <= len(text) <= 160:
            if not any(text.strip() == c["source"]["quote"] for c in bucket if c["unit"] == ""):
                add_candidate(bucket, file=rel, locator=f"block {lineno}", sections=current,
                              label="关键结论句", quote=text.strip()[:160],
                              value=text.strip()[:120], unit="", confidence="medium",
                              order=(current[0] if current else "zz", lineno, "~", ""),
                              doc_version=doc_version(path.name))
    return scanned

def harvest_tables(path: Path, bucket: list, sections: list[dict] | None = None) -> int:
    rel = (Path(DUMP_DIR) / path.name).as_posix()
    table_no, header, scanned = 0, [], 0
    for line in read_text(path).splitlines():
        tm = TABLE_RE.match(line.strip())
        if tm:
            table_no, header = int(tm.group(1)), []
            continue
        rm = ROW_RE.match(line)
        if not rm or not table_no:
            continue
        cells = [c.strip() for c in rm.group(2).split("|")]
        if not header:
            header = cells
            continue
        scanned += 1
        rowno = int(rm.group(1))
        row_sections = table_sections(header, sections or [])
        for col, cell in enumerate(cells):
            for num, unit in NUM_RE.findall(cell):
                head = header[col] if col < len(header) else ""
                add_candidate(bucket, file=rel, locator=f"table {table_no} row {rowno}",
                              sections=row_sections,
                              label=(f"{head}｜{cell}" if head else cell)[:40],
                              quote=cell[:160], value=float(num), unit=unit, confidence="high",
                              order=((row_sections[0] if row_sections else "zz"), table_no, rowno, col, num),
                              doc_version=doc_version(path.name))
    return scanned

def harvest_pdf_evidence(pdf_dir: Path, limit: int = 200) -> list[dict]:
    """从 01 抽出的 PDF 文本（`[PAGE n]` 分页）收集「可引用证据」。

    **只进 pdfEvidence，不并入 facts**：PDF 正文不作章节结构来源、不生成正文，
    仅作为带页码的引用证据供写作任务使用。
    """
    if not pdf_dir.is_dir():
        return []
    out: list[dict] = []
    for path in sorted(pdf_dir.glob("*.blocks.txt")):
        page = 0
        rel = f"{PDF_TEXT_DIR.as_posix()}/{path.name}"
        for raw in read_text(path).splitlines():
            line = raw.strip()
            if not line:
                continue
            pm = PAGE_RE.match(line)
            if pm:
                page = int(pm.group(1))
                continue
            numbers = NUM_RE.findall(line)
            if numbers:
                num, unit = numbers[0]
                out.append({
                    "source": {"file": rel, "locator": f"page {page}",
                               "quote": line[:160], "docVersion": "pdf"},
                    "label": line[:40], "value": float(num), "unit": unit,
                    "sourceType": "pdf_text", "confidence": "medium",
                })
            elif KEY_SENTENCE.search(line) and 8 <= len(line) <= 160:
                out.append({
                    "source": {"file": rel, "locator": f"page {page}",
                               "quote": line[:160], "docVersion": "pdf"},
                    "label": "关键结论句", "value": line[:120], "unit": "",
                    "sourceType": "pdf_text", "confidence": "low",
                })
            if len(out) >= limit:
                return out
    return out


def select_facts(candidates: list[dict], sections: list[dict], max_facts: int) -> tuple[list[dict], list[dict]]:
    """按模板节配额挑选（保证覆盖各节），再按顺序补足；未选入的进未采纳清单。"""
    rejected: list[dict] = []
    pool = []
    for c in candidates:
        if not c["section"] and c["source"]["locator"].startswith("block "):
            rejected.append({"source": c["source"]["file"], "locator": c["source"]["locator"], "quote": c["source"]["quote"],
                             "reason": "未匹配到模板章节（封面/目录/章外），不采纳"})
        else:
            pool.append(c)
    block_pool = [c for c in pool if c["source"]["locator"].startswith("block ")]
    table_pool = [c for c in pool if not c["source"]["locator"].startswith("block ")]
    facts: list[dict] = []
    if sections:
        per_section = max(1, (max_facts * 3 // 5) // len(sections))
        for sec in sections:
            sid = sec["section"]
            # 先取本节的正文事实，再取本节的表格事实——表格里的数据
            # （如"现状基本情况"表逐行给出各小区建成年代/户数/建筑面积）
            # 往往正是现状分析的核心，不应因排在补充池末尾而拿不到。
            facts.extend([c for c in block_pool if c["section"] == sid][:per_section])
            facts.extend([c for c in table_pool if c["section"] == sid][:per_section])
    else:
        facts.extend(block_pool[: max_facts * 3 // 5])
    chosen = {id(c) for c in facts}
    for c in block_pool + table_pool:
        if len(facts) >= max_facts:
            break
        if id(c) not in chosen:
            facts.append(c)
            chosen.add(id(c))
    facts.sort(key=lambda c: c["_order"])
    for c in pool:
        if id(c) not in chosen:
            rejected.append({"source": c["source"]["file"], "locator": c["source"]["locator"], "quote": c["source"]["quote"],
                             "reason": f"超出配额/上限（--max-facts={max_facts}），保留候选可重跑"})
    return facts, rejected

def detect_conflicts(facts: list[dict]) -> list[dict]:
    """保守冲突判定：叙述段、同句式/同单位、恰好两处各一条且值不同 → 并列 conflict。"""
    groups: dict[tuple, list[dict]] = {}
    for f in facts:
        if not f["source"]["locator"].startswith("block ") or not f["unit"]:
            continue
        phrase = metric_phrase(f["source"]["quote"])
        if len(phrase) >= 12:
            groups.setdefault((f["section"], phrase, f["unit"]), []).append(f)
    conflicts = []
    for (section, phrase, unit), items in groups.items():
        values = {f["value"] for f in items}
        locators = {f["source"]["locator"] for f in items}
        if len(items) == 2 and len(locators) == 2 and len(values) == 2:
            for f in items:
                f["status"] = "conflict"
            conflicts.append({"section": section, "phrase": phrase, "unit": unit,
                              "values": sorted(values), "locators": sorted(locators)})
    return conflicts

def build(project: Path, opts) -> tuple[dict, list[dict]]:
    sections = template_sections(project)
    blocks, tables = pick_sources(project)
    bucket: list[dict] = []
    scanned = {"blocks": 0, "tables": 0}
    for path in blocks:
        scanned["blocks"] += harvest_blocks(path, sections, bucket)
    for path in tables:
        scanned["tables"] += harvest_tables(path, bucket, sections)
    seen, uniq = set(), []
    for c in bucket:
        key = (c["source"]["file"], c["source"]["locator"], str(c["value"]), c["unit"])
        if key not in seen:
            seen.add(key)
            uniq.append(c)
    uniq.sort(key=lambda c: (0, c["_order"]) if isinstance(c.get("_order"), (int, float))
              else (1, str(c.get("_order"))))
    facts, rejected = select_facts(uniq, sections, max(1, int(opts.max_facts)))
    conflicts = detect_conflicts(facts)
    raw_pdf_dir = str(getattr(opts, "pdf_dir", "") or "").strip()
    pdf_dir = Path(raw_pdf_dir).expanduser() if raw_pdf_dir else project / PDF_TEXT_DIR
    if raw_pdf_dir and not pdf_dir.is_absolute():
        pdf_dir = project / pdf_dir
    pdf_evidence = harvest_pdf_evidence(pdf_dir)
    for f in facts:
        f.setdefault("status", "confirmed")
    for i, f in enumerate(facts, start=1):
        f["id"] = f"F-{i:04d}"
    covered: set[str] = set()
    for f in facts:
        covered.update(f.get("sections") or ([f["section"]] if f["section"] else []))
    missing = [{"section": s["section"], "heading": s["heading"],
                "note": "该模板节未抽到可复核事实（不推断、不补数）"}
               for s in sections if s["section"] not in covered]
    stats = {"dumpFiles": len(blocks) + len(tables), "blocksScanned": scanned["blocks"],
             "tablesScanned": scanned["tables"], "candidates": len(uniq), "facts": len(facts),
             "numericFacts": sum(1 for f in facts if f["unit"]),
             "textFacts": sum(1 for f in facts if not f["unit"]),
             "unitMissing": sum(1 for f in facts if not f.get("unit")),
             "conflicts": len(conflicts), "missingSections": len(missing),
             "sectionsTotal": len(sections), "sectionsCovered": len(covered),
             "rejected": len(rejected), "pdfEvidence": len(pdf_evidence)}
    for f in facts:
        f.pop("_order", None)
    payload = {"schema": SCHEMA}
    if not opts.no_timestamp:
        payload["generatedAt"] = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    payload.update({"project": ".", "stats": stats, "facts": facts, "conflicts": conflicts,
                    "missing": missing, "rejected": rejected, "pdfEvidence": pdf_evidence})
    return payload, facts

def render_md(payload: dict) -> str:
    s = payload["stats"]
    lines = ["# 现状事实库", "", f"- schema：{SCHEMA}",
             f"- 覆盖率：有事实的模板节 {s['sectionsCovered']}/{s['sectionsTotal']}；"
             f"带单位事实 {s['numericFacts']} / 文本事实 {s['textFacts']}"
             f"（其中单位缺失 {s['unitMissing']}）",
             f"- 扫描：blocks {s['blocksScanned']} 段 / tables {s['tablesScanned']} 行；候选 {s['candidates']}，"
             f"采纳 {s['facts']}，未采纳 {s['rejected']}，冲突 {s['conflicts']}，缺失节 {s['missingSections']}", ""]
    by_section: dict[str, list[dict]] = {}
    for f in payload["facts"]:
        by_section.setdefault(f["section"] or "（未归属）", []).append(f)
    for section in sorted(by_section):
        lines += [f"## {section}", "",
                  "| id | 值 | 单位 | 标签 | 定位 | 置信 | 状态 |",
                  "| --- | --- | --- | --- | --- | --- | --- |"]
        for f in by_section[section]:
            lines.append(f"| {f['id']} | {f['value']} | {f['unit']} | {f['label']} | "
                         f"{f['source']['file']} · {f['source']['locator']} | {f['confidence']} | {f['status']} |")
        lines.append("")
    if payload["conflicts"]:
        lines += ["## 冲突（并列，不静默选一个）", ""]
        lines += [f"- {c['phrase']}（{c['unit']}）：值 {c['values']}；定位 {c['locators']}"
                  for c in payload["conflicts"]]
        lines.append("")
    if payload["missing"]:
        lines += ["## 缺失的模板节（出缺项，不补数）", ""]
        lines += [f"- {m['section']} {m['heading']}：{m['note']}" for m in payload["missing"]]
        lines.append("")
    if payload.get("pdfEvidence"):
        lines += ["## PDF 证据（供引用，不参与主线；引用须标文件名 + 页码）", "",
                  "| 值 | 单位 | 标签 | 出处 | 置信 |",
                  "| --- | --- | --- | --- | --- |"]
        for e in payload["pdfEvidence"]:
            lines.append(f"| {e['value']} | {e['unit']} | {e['label']} | "
                         f"{e['source']['file']} · {e['source']['locator']} | {e['confidence']} |")
        lines.append("")
    return "\n".join(lines)

def render_rejected(payload: dict) -> str:
    lines = ["# 未采纳清单（候选事实）", ""]
    if not payload["rejected"]:
        lines.append("（无）")
    lines += [f"- {row['source']} · {row['locator']}：{row['quote'][:80]} —— {row['reason']}"
              for row in payload["rejected"]]
    return "\n".join(lines) + "\n"

def write_if_changed(path: Path, text: str) -> bool:
    if path.is_file():
        try:
            if read_text(path) == text:
                return False
        except OSError:
            pass
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")
    return True

def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    parser = argparse.ArgumentParser(description="现状事实底座薄抽取（T22）")
    parser.add_argument("--project-dir", default="", help="项目目录（任意路径；最高优先级）")
    parser.add_argument("--root", default="", help="[校准/回归] 项目库根（与 --project 搭配）")
    parser.add_argument("--project", default="", help="[校准/回归] 项目名")
    parser.add_argument("--direction", default="", help="[校准/回归] 方向")
    parser.add_argument("--out", default="", help="输出目录（缺省 <项目>/05_过程/现状事实库/）")
    parser.add_argument("--max-facts", type=int, default=400, help="最多采纳事实数（缺省 400）")
    parser.add_argument("--dump-dir", default="", help="原文转储目录（缺省 <项目>/05_过程/原文转储/）")
    parser.add_argument("--pdf-dir", default="", help="PDF 文本目录（缺省 <项目>/02_分类材料/11_PDF文本/；只收证据，不进 facts）")
    parser.add_argument("--mapping", default="", help="章节映射 YAML（缺省 <项目>/03_拆解成果/章节映射.yaml）")
    parser.add_argument("--no-timestamp", action="store_true", help="不写 generatedAt（幂等测试用）")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    opts = parser.parse_args(argv)
    if str(opts.dump_dir).strip():
        global DUMP_DIR
        DUMP_DIR = Path(opts.dump_dir).expanduser()
        if not DUMP_DIR.is_absolute():
            DUMP_DIR = Path.cwd() / DUMP_DIR
    if str(opts.mapping).strip():
        global MAPPING
        MAPPING = Path(opts.mapping).expanduser()
        if not MAPPING.is_absolute():
            MAPPING = Path.cwd() / MAPPING
    try:
        project = resolve_project_dir(opts)
        out_dir = Path(opts.out).expanduser() if str(opts.out).strip() else project / OUT_DIR
        payload, facts = build(project, opts)
    except BadInput as exc:
        print(f"[facts] 错误：{exc}", file=sys.stderr)
        return 2
    if not facts:
        print("[facts] 抽取失败：未抽到任何可复核事实（检查原文转储是否为空/格式是否受支持）", file=sys.stderr)
        return 1
    try:
        files = (("事实库.json", json.dumps(payload, ensure_ascii=False, indent=2) + "\n"),
                 ("事实库.md", render_md(payload)), ("未采纳清单.md", render_rejected(payload)))
        changed = [name for name, text in files if write_if_changed(out_dir / name, text)]
    except OSError as exc:
        print(f"[facts] 写入失败：{exc}", file=sys.stderr)
        return 1
    s = payload["stats"]
    print(f"[facts] 覆盖率 {s['sectionsCovered']}/{s['sectionsTotal']} 节；事实 {s['facts']}"
          f"（带单位 {s['numericFacts']}）；冲突 {s['conflicts']}；缺失节 {s['missingSections']}")
    print(f"[facts] 产物：{out_dir.name}/{{事实库.json,事实库.md,未采纳清单.md}}"
          f"（{'已更新 ' + '、'.join(changed) if changed else '幂等：内容一致，未重写'}）")
    return 0

if __name__ == "__main__":
    sys.exit(main())