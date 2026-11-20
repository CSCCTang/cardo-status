#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""assemble_materials.py -- 把散材料按模板章节汇编成一份主线 docx。

设计前提：**结构由模板决定，材料只提供事实**。
用户不需要自己攒一份分节的 Word；只要把材料丢进 输入文件/，
本脚本按模板的 H1/H2 骨架分节，把材料里的文本块归到最匹配的节下，
产出 <out>，供 03/04/05 当作"源报告"继续跑。

用法:
    python assemble_materials.py --input <输入文件目录> --template <模板docx> --out <输出docx>

退出码: 0 成功；2 参数或输入错误；1 运行失败。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

# 形式参考件排除集（绝对路径，posix 风格）；由 --exclude-json 在 main 中填充
_EXCLUDE_ABS: set[str] = set()

try:
    import docx
    from docx.shared import Pt
except ImportError:
    print("[assemble] 缺少 python-docx，无法汇编", file=sys.stderr)
    raise SystemExit(1)

TEXT_SUFFIX = {".docx", ".txt", ".md", ".pptx", ".json"}
SHEET_SUFFIX = {".xlsx", ".csv"}
# 图片也要进汇编：项目信息卡、现状照片上的文字常写着建设单位、资金来源、
# 时间节点——这些是现状分析的第一手事实。
# 实测「项目信息.png」里写着某街道为业主、2027.06–2028.06 实施周期、
# 资金来源含中央预算内投资与专项债券，整张图此前被跳过，
# 「项目概况」节因此完全无材料。
IMAGE_SUFFIX = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
SKIP_DIR_PARTS = ("输入文件-模板", "提资数据", "过程文件", "成果文件", ".cardo", "02_分类材料")
PLACEHOLDER = "（本节材料不足，见缺项表）"
# 结构化缺口标记：汇编时无材料的小节写入本行，04 章节映射据此判 gap（不靠标题猜）。
GAP_MARK = "【缺口】本节无材料，未作推断；已登记缺项（见缺项表）。"


# 表名/工作表名 → 模板节标题 的语义路由。
# 表格内容与节标题的措辞常常对不上（「现状基本情况」表 vs「建筑现状与风貌保护现状」），
# 靠二元组重合度打分会把表派到措辞偶然相似的节里去——实测「现状基本情况」表
# 被派到了「综合现状研判」节，现状分析最核心的那张表整个用不上。
# 表格改按表名语义直接路由。
TABLE_NAME_SECTIONS = [
    ("现状基本情况", "建筑现状与风貌保护现状"),
    ("现状", "建筑现状与风貌保护现状"),
    ("基本", "建筑现状与风貌保护现状"),
    ("改造清单", "土地与空间利用现状"),
    ("清单", "土地与空间利用现状"),
    ("栋数", "建筑现状与风貌保护现状"),
    ("面积", "土地与空间利用现状"),
    ("能耗", "生态与人居环境现状"),
    ("规划", "规划要求"),
    ("片区", "区位条件"),
    ("投资", "项目概况"),
    ("估算", "项目概况"),
    ("资金", "项目概况"),
    ("进度", "项目概况"),
]

# 表格**表头列名** → 模板节标题。
# 工作表名常常是 "Sheet1"（Excel 默认），此时路由靠表名必然落空——
# 实测「改造小区基本信息及改造内容统计表」整张表被派到第一节，
# 「建筑现状与风貌保护现状」这类最该用它的章节反而空着。
# 故再按**表头列名**判一次：楼栋/建成年代/结构形式 → 建筑现状；
# 户数/人口/居住 → 人口与社会人居；工程量/面积 → 土地与空间利用。
TABLE_HEADER_SECTIONS = [
    ("工程量", "建筑现状与风貌保护现状"),
    ("改造项目", "建筑现状与风貌保护现状"),
    ("立面材料", "建筑现状与风貌保护现状"),
    ("屋面", "建筑现状与风貌保护现状"),
    ("外墙雨棚", "建筑现状与风貌保护现状"),
    ("冷凝水管", "建筑现状与风貌保护现状"),
    ("落水管", "建筑现状与风貌保护现状"),
    ("空调外机", "建筑现状与风貌保护现状"),
    ("楼栋", "建筑现状与风貌保护现状"),
    ("建成年代", "建筑现状与风貌保护现状"),
    ("结构形式", "建筑现状与风貌保护现状"),
    ("建筑性质", "建筑现状与风貌保护现状"),
    ("建筑层数", "建筑现状与风貌保护现状"),
    ("建筑高度", "建筑现状与风貌保护现状"),
    ("建筑面积", "建筑现状与风貌保护现状"),
    ("立面", "建筑现状与风貌保护现状"),
    ("风貌", "建筑现状与风貌保护现状"),
    ("户数", "人口与社会人居现状"),
    ("常住人口", "人口与社会人居现状"),
    ("人口", "人口与社会人居现状"),
    ("居住", "人口与社会人居现状"),
    ("权属", "人口与社会人居现状"),
    ("业态", "街区业态与经营现状"),
    ("经营", "街区业态与经营现状"),
    ("商铺", "街区业态与经营现状"),
    ("管网", "市政基础设施现状"),
    ("管线", "市政基础设施现状"),
    ("道路", "交通与通行现状"),
    ("停车", "交通与通行现状"),
    ("公交", "交通与通行现状"),
    ("配套", "公共服务配套现状"),
    ("服务设施", "公共服务配套现状"),
    ("绿化", "生态与人居环境现状"),
    ("环境", "生态与人居环境现状"),
    ("能耗", "生态与人居环境现状"),
]


def route_table_section(table_name: str, h2_sections) -> str:
    """按表名把表格派到最合适的节标题；判不出返回空串（走常规打分）。"""
    name = str(table_name or "")
    for key, want in TABLE_NAME_SECTIONS:
        if key in name:
            for sec in h2_sections:
                if sec[1] == want:
                    return want
    return ""


def route_by_header(header_text: str, h2_sections) -> str:
    """按**表头列名**把表格派到节标题。

    工作表名多半是 "Sheet1"，表名路由必然落空；表头里的列名
    （楼栋/建成年代/结构形式/户数/人口/工程量…）才真正说明这张表是什么。
    命中多个时，取命中列名最多、且节标题存在的那一个。
    """
    text = str(header_text or "")
    if not text:
        return ""
    best, best_hits = "", 0
    for key, want in TABLE_HEADER_SECTIONS:
        if key not in text:
            continue
        if not any(sec[1] == want for sec in h2_sections):
            continue
        hits = text.count(key)
        if hits > best_hits:
            best, best_hits = want, hits
    return best


def read_template_outline(template: Path) -> list[tuple[int, str]]:
    """读模板的 H1/H2 骨架，返回 [(level, title)]。"""
    doc = docx.Document(str(template))
    out = []
    for para in doc.paragraphs:
        style = (para.style.name or "").strip()
        text = para.text.strip()
        if not text:
            continue
        m = re.match(r"^(?:Heading|标题)\s*(\d)$", style)
        if m:
            level = int(m.group(1))
            if level <= 2:
                out.append((level, text))
    return out


# 中文地名里表示「一个地标」的常见结尾字。
# 「某街某片区片区」是三个地标连写：某街 / 某片区 / 片区。
# 材料里常常只出现其中一个（社区叫「某片区社区」，地址叫「某街64号」），
# 故按这些字切边界，把子片段一并作为关键词。
_NAME_BOUNDARY_CHARS = set("街巷路道台寺庙里村园苑号桥")

# 切出来的片段若本身是通用词，不能作归属依据——
# 「片区」「老旧街」这类词几乎每个项目都有，拿它判归属会把别的项目材料拉进来。
_NAME_FRAGMENT_STOPWORDS = {
    "片区", "街区", "小区", "厂区", "城区", "园区", "社区",
    "老旧街", "老旧小", "旧街区", "更新区", "改造区",
    "项目", "工程", "建设", "更新", "改造", "提升", "整治",
}


def _name_fragments(name: str) -> list[str]:
    """按地标边界字切出子片段（降序长度）。

    「某街某片区片区」→ 某街某片区 / 某街 / 某片区
    「某项目片区」→ 某项目
    只出长度 ≥2 且有区分度的片段；单字地标（太泛）与通用词一律不要。
    """
    text = str(name or "")
    if len(text) < 3:
        return []
    # 找所有边界字位置，按边界把串切成若干地标
    cuts = [i for i, ch in enumerate(text) if ch in _NAME_BOUNDARY_CHARS]
    if not cuts:
        return []
    # 边界字本身属于前一段的结尾（如「××**街**」的「街」），故切点取边界字之后
    pieces = []
    start = 0
    for cut in cuts:
        end = cut + 1
        seg = text[start:end]
        if len(seg) >= 2:
            pieces.append(seg)
        start = end
    tail = text[start:]
    if len(tail) >= 2:
        pieces.append(tail)
    out = []
    for seg in pieces:
        # 通用词片段直接丢（「片区」这种没有区分度）
        if seg in _NAME_FRAGMENT_STOPWORDS:
            continue
        # 含通用词且自身过长的片段也丢（「老旧街」这种）
        if len(seg) >= 3 and any(w in seg for w in _NAME_FRAGMENT_STOPWORDS):
            continue
        if seg not in out:
            out.append(seg)
    # 长片段在前，匹配时优先用长的
    return sorted(set(out), key=len, reverse=True)


def project_keywords(project: str) -> list[str]:
    """从项目全称里取出用于筛选材料的关键词。

    项目全称常是「某市某新区老旧小区改造项目」，而材料目录名是
    「3 某新区老旧小区改造-20240802」——两者并不逐字相同。
    故这里切成**多个实词**，任一命中即算同一项目（「某市」或「某新区」都能命中）。

    取法：先按常见分隔符切，再对「XX市YY区ZZ」这类连写串按地名后缀再切一刀。
    「项目」「工程」这类通用词不带区分度，全部去掉——否则会把所有项目都拉进来。

    连写串还要再切：**「某街某片区片区」是一个连写串，但材料里往往只出现
    「某片区」或「某街」其中之一**（社区名、地址名各叫一边）。
    只按行政区划后缀切会把它留成一个整体词，于是本项目的材料被误排除。
    故额外按「街/巷/路/道/台/寺/里/村/园/苑/号/苑」等地名常见字切出**子片段**，
    「某街某片区片区」→「某街某片区片区 / 某街某片区 / 某街 / 某片区」，
    任一命中即算同一项目。
    """
    if not project:
        return []
    text = str(project)
    # 先剥掉尾部通用词（"…改造项目" → "…改造"），否则整串会被当成一个词
    generic = {"项目", "工程", "改造", "提升", "更新", "建设", "片区", "本项目",
               "项目名称", "实施", "综合", "计划", "专项"}
    raw = re.split(r"[\s\-—_、，,。.（）()《》【】\[\]/]+", text)
    out = []
    for w in raw:
        w = w.strip()
        # 剥尾部通用词（"…改造项目" → "…改造"）：按整词 endswith 剥离，
        # 不能用 w[-1] in generic（那是「单字符是否在集合中」的判断，恒为假）。
        changed = True
        while changed and len(w) >= 2:
            changed = False
            for g in generic:
                if len(g) >= 2 and w.endswith(g) and len(w) - len(g) >= 2:
                    w = w[:-len(g)]
                    changed = True
                    break
        if len(w) < 2:
            continue
        # 连写串按行政区划后缀切：「某市某新区老旧小区」→ 某市 / 某新区 / 老旧小区
        m = re.match(r"^(.+?省|.+?市)?(.+?区|.+?县)?(.*)$", w)
        parts = [p for p in m.groups() if p] if m else [w]
        for part in parts:
            if len(part) < 2 or part in generic:
                continue
            out.append(part)
            # 再按地名常见字切子片段：材料里往往只出现其中一段
            # （「某街某片区片区」→ 材料写「某片区社区…」）。
            for frag in _name_fragments(part):
                if frag != part:
                    out.append(frag)
    # 去重并按长度降序
    seen, res = set(), []
    for w in sorted(set(out), key=len, reverse=True):
        if w not in seen:
            seen.add(w)
            res.append(w)
    # 分级：行政区/项目专名是「专属词」，方向词是「通用词」。
    # 只用专属词过滤才守得住项目边界；方向词（老旧小区/老旧街区…）同方向
    # 的其他项目全都命中，会把别的项目材料混进来。
    direction_words = {"老旧小区", "老旧街区", "老旧厂区", "城市更新", "片区更新"}
    specific = [w for w in res if w not in direction_words]
    generic = [w for w in res if w in direction_words]
    return specific, generic


def _same_project(rel: str, keywords: list[str]) -> bool:
    """判断某个材料路径是否属于本项目。

    匹配是**双向**的：
      * 关键词出现在路径里（路径含「某新区老旧小区改造」）；
      * 路径里的片段出现在关键词里（项目名含「某市某新区」，而目录名只有
        「某新区」——项目名多带行政区前缀，目录名往往省略）。
    只做单向匹配会漏掉大量本项目材料。

    注意：keywords 是**分级**的（最具区分度的词在前），调用方按序尝试，
    任一级命中即算同一项目——不是把所有词的并集一起判。
    并集会把「老旧小区」这种方向词也算进去，于是同方向的其他项目
    （某区、某市、某路…）的材料全部混进来，项目边界失效。
    """
    rel = rel.replace("\\", "/")
    parts = [seg for seg in rel.split("/") if seg]
    for k in keywords:
        if not k or len(k) < 2:
            continue
        if k in rel:
            return True
        for seg in parts:
            # 目录名常带序号前缀与年份后缀，比对前先剥掉
            bare = re.sub(r"^[0-9]+[.、\s]*", "", seg)
            bare = re.sub(r"[-—_]\d{4,}$", "", bare)
            for probe in (seg, bare):
                if len(probe) < 3:
                    continue
                if probe in k:
                    return True
                # 反向：项目名里的行政区段（"某市某新区"）出现在目录名里
                # （目录名常省略行政前缀，只写"某新区"）——取关键词的连续子串比对。
                # 注意子串长度下限必须 ≥3：Python 切片 k[i:j] 在 j>i+len(k) 时会截短，
                # 用 "区" 这种单字去比对会把所有带「院/区/市」字样的别的项目全放进来。
                n = len(probe)
                for i in range(n):
                    for j in range(i + 3, n + 1):
                        sub = k[i:j]
                        if len(sub) >= 3 and sub in probe:
                            return True
    return False


def iter_candidate_files(root: Path, keywords=None, exclude_keywords=None):
    """遍历候选材料文件。

    keywords 非空时按**项目边界**过滤：只取路径中命中项目关键词的文件。
    投递区常常装着多个项目的材料（如课题提资包同时含厂区/小区/街区三类项目），
    不做边界过滤会把别的项目的数据写进本项目的现状报告——
    这是取料纪律问题，不是格式问题。

    两级判定：
      ① 路径级 —— 路径命中项目专属词即算本项目（快，覆盖大多数）。
      ② 内容级 —— 路径没命中时，读文件内容找项目专属词（**只认最长最具体的**）。
         社区台账这类材料常常只在**内容里**写着本项目的地址（社区叫别的名字，
         但表里全是「某街64号」），只看路径会把它误排除。
         内容级判定只用**最长的那个专属词**（如「某街某片区片区」），
         且它必须是文件的**专属归属词**——避免「老旧街区」这种通用词把别的项目拉进来。

    exclude_keywords 是**其它项目的专属词**（如某项目）。内容级判定时：
    命中本项目词、但同时也命中别个项目词的材料 → 判为别个项目的，排除。
    （前期沟通 PPT 常把兄弟片区的内容列在同一页里，只看「命中了本项目词」
    会把某项目的 PPT 拉进某街的报告。）
    """
    content_cache: dict[Path, str] = {}

    def content_of(path: Path) -> str:
        if path not in content_cache:
            text = ""
            try:
                suffix = path.suffix.lower()
                if suffix == ".docx":
                    text = blocks_from_docx(path)
                elif suffix == ".pptx":
                    text = blocks_from_pptx(path)
                elif suffix == ".xlsx":
                    text = blocks_from_xlsx(path)
                elif suffix in IMAGE_SUFFIX:
                    text = blocks_from_image(path)
                elif suffix in TEXT_SUFFIX:
                    text = blocks_from_text(path)
                content_cache[path] = "\n".join(text) if isinstance(text, list) else str(text)
            except Exception:
                content_cache[path] = ""
        return content_cache[path]

    # 内容级判定可用**全部专属词**（专属词已排除方向/通用词，
    # 不会把同方向的其他项目拉进来）。只用最长的那个会漏：
    # 「某街某片区片区」的社区台账里写的是「某街64号」，不含整串。
    content_keywords = list(keywords or [])
    # 其它项目的词：内容里出现它们，说明这份材料主体属于别的项目
    other_keywords = list(exclude_keywords or [])

    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        if str(path.resolve()).replace("\\", "/") in _EXCLUDE_ABS:
            continue
        if any(part in str(path) for part in SKIP_DIR_PARTS):
            continue
        if path.name.startswith("~$") or "空白" in path.name:
            continue
        if path.suffix.lower() not in TEXT_SUFFIX | SHEET_SUFFIX | IMAGE_SUFFIX:
            continue
        if keywords:
            rel = path.relative_to(root).as_posix()
            if not _same_project(rel, keywords):
                # 路径没命中 → 内容级判定：内容里写了本项目的专名才算
                if not content_keywords:
                    continue
                body = content_of(path)
                if not any(kw in body for kw in content_keywords):
                    continue
                # 内容里同时也出现了别个项目的专名 → 这份材料主体是别个项目的
                if other_keywords and any(kw in body for kw in other_keywords):
                    continue
        yield path


def _clean_cell(value) -> str:
    """把单元格值压成单行。

    Excel 单元格里的换行（"序号\n（先不填）"）若原样写进块文本，
    整行会被转储工具按行切开——表头只剩「序号」，数据行也散架，
    05 步按列名抽事实就完全对不上列了。
    故这里把换行/回车/制表统一换成空格（多个空格压成一个）。
    """
    if value is None:
        return ""
    text = str(value).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    text = text.replace("\t", " ").strip()
    return re.sub(r"\s{2,}", " ", text)


def blocks_from_docx(path: Path) -> list[str]:
    doc = docx.Document(str(path))
    out = []
    for para in doc.paragraphs:
        text = para.text.strip()
        style = (para.style.name or "")
        if not text or re.match(r"^(?:Heading|标题)\s*\d$", style):
            continue
        out.append(text)
    for ti, table in enumerate(doc.tables, 1):
        rows = []
        for row in table.rows:
            cells = [_clean_cell(c.text) for c in row.cells]
            cells = [c for c in cells if c]
            if cells:
                rows.append(cells)
        if not rows:
            continue
        # 同 xlsx：表头一块、数据行每行一块，保住列与值的对应关系
        out.append("表（表%d）表头：%s" % (ti, " | ".join(rows[0])))
        for cells in rows[1:]:
            out.append("表（表%d）数据行：%s" % (ti, " | ".join(cells)))
    return out


def blocks_from_xlsx(path: Path) -> list[str]:
    """把 xlsx 拆成「表头一块 + 数据行每行一块」。

    两个关键点，都是实测踩出来的：

    ① **必须按列位置对齐**，不能把行首的空单元格丢掉。
       楼栋台账用了合并单元格：openpyxl 读出的每行格数不同（21/6/17/16…），
       若按「非空格」拼行，各行的「第 N 格」根本不是同一列——户数会落到
       建筑面积的列上。读全列（含空值）再按位置取值才对得上。

    ② **一张 sheet 里可能并着多张表**，各自的列名行位置不同。
       实测楼栋台账左半是基本信息（列名在第 1 行）、右半是工程量
       （列名在第 2 行，只占右侧几列）。若把所有行当成一张表，
       右表的数据就会用左表的列名去读——抽出"楼栋数｜0.315"这种错位事实。

       解法：**逐列找列名行**。对每一列，找第一行「该列有文字且是说明性文字
       （非纯数值）」的行作为该列的列名；列名行以下的数据行按列位置对齐取值。
       列区间由此自然切分：左表列名行与右表列名行不同，各自成表。
    """
    out = []
    try:
        import openpyxl
    except ImportError:
        return out
    try:
        wb = openpyxl.load_workbook(str(path), data_only=True)
    except Exception:
        return out
    try:
        for ws in wb.worksheets:
            rows = []
            for row in ws.iter_rows(values_only=True):
                cells = [_clean_cell(v) for v in row]
                # 行尾的空列裁掉（Excel 常有一堆空列），但行内的空保留（占位）
                while cells and not cells[-1]:
                    cells.pop()
                rows.append(cells)
                if len(rows) >= 80:
                    break
            # 去掉整行全空的行
            rows = [r for r in rows if any(c for c in r)]
            if not rows:
                continue
            width = max(len(r) for r in rows)
            # 补齐到统一宽度，保证按列位置取值
            rows = [r + [""] * (width - len(r)) for r in rows]

            # 全表标题行（大标题）：格数少、内容是长句。实测楼栋台账第 0 行是
            # 「各小区改造楼栋基本信息统计表… | 各小区改造工程量统计表…」，
            # 横跨整行、只有 2 格——它不是列名，只是表题。
            title_row = None
            for ri, r in enumerate(rows[:2]):
                nz = [c for c in r if c]
                if len(nz) <= 2 and any(len(c) >= 8 for c in nz):
                    title_row = ri
                    break

            # 逐列找列名行。**子表列名行**要特别处理：
            # 实测楼栋台账横向并着两张表——左表列名在第 1 行（序号|街道|户数…），
            # 右表（工程量）列名在第 2 行且只占最右 6 列（工程量（㎡）|…）。
            # 逐列取「第一个非空行」会把右表的列名行当成左表的**数据行**。
            # 判据：某行若非空列**全部集中在最右端**、且内容形如「XX（单位）」，
            # 它就是子表的列名行 → 单独成表，数据行只取对应列区间。
            sub_head_rows: dict[int, list[int]] = {}      # 行号 → 该子表占的列
            main_head = -1
            # 先定主表列名行：含中文列名最多、且非空列最多的行
            best = (-1, -1)
            for ri, r in enumerate(rows[:4]):
                if title_row is not None and ri == title_row:
                    continue
                named = sum(1 for c in r if c and re.search(r"[一-鿿]", c))
                if (named, len([c for c in r if c])) > best:
                    best = (named, len([c for c in r if c]))
                    main_head = ri
            for ri, r in enumerate(rows):
                if ri <= main_head or ri == title_row:
                    continue
                nz = [i for i, c in enumerate(r) if c]
                if not nz:
                    continue
                # 非空列全部在最右端连续一段 → 子表列名行
                if nz == list(range(nz[0], nz[-1] + 1)) and nz[0] > 0 \
                        and all(c and not _is_pure_num(c) for c in r if c):
                    sub_head_rows[ri] = nz
            # 主表列名行之后、子表列名行之前的那一行，常常是子表的列名行
            for ri in sorted(sub_head_rows):
                pass

            col_head: list[int] = [-1] * width
            for i in range(width):
                for ri, r in enumerate(rows):
                    if title_row is not None and ri == title_row:
                        continue          # 大标题行不作列名
                    c = r[i]
                    if not c or _is_pure_num(c):
                        continue
                    # 该列若属于某个子表，且子表列名行更靠后 → 用子表的
                    owners = [sr for sr, cols in sub_head_rows.items() if i in cols]
                    if owners:
                        col_head[i] = min(owners)
                        break
                    col_head[i] = ri
                    break

            # 按列名行把列切成若干「列区间」，每个区间一张表。
            # **列名行相同的相邻列要合并**——实测"序号 （先不填）"这一列的
            # 列名行是第 1 行，与它右边几列同属一张表，但首列因为
            # 「上一行该列为空」的取列名规则被切走了。
            # 故先按列名行分组，再按「组内列名行相同且相邻」合并。
            groups: list[tuple[int, list[int]]] = []
            cur_key: int | None = None
            cur_cols: list[int] = []
            for i in range(width):
                key = col_head[i]
                if key != cur_key:
                    if cur_cols:
                        groups.append((cur_key if cur_key is not None else -1, cur_cols))
                    cur_key, cur_cols = key, []
                cur_cols.append(i)
            if cur_cols:
                groups.append((cur_key if cur_key is not None else -1, cur_cols))

            # 合并：**只合并列名行完全相同**的相邻区间。
            # 不按「行号相近」合并——实测左表列名在第 1 行、右表在第 2 行，
            # 行号只差 1，但它们是两套完全不同的列，合并后右表数据会错位。
            merged_groups: list[tuple[int, list[int]]] = []
            for key, cols in groups:
                if merged_groups and merged_groups[-1][0] == key:
                    merged_groups[-1][1].extend(cols)
                    continue
                merged_groups.append((key, list(cols)))

            # 子表（右表）补上主表的**标识列**（小区名称/地址/社区）：
            # 工程量表只有"工程量（㎡）|…"几列，光看它说不出这些量对应哪个小区；
            # 补上标识列后，05 步抽出的事实才带归属，成文能写"XX小区立面修缮 500㎡"。
            id_cols: list[int] = []
            main_cols = merged_groups[0][1] if merged_groups else []
            main_key = merged_groups[0][0] if merged_groups else -1
            # 只补**最具体**的标识列（小区名称），别把街道/社区一起搬过来——
            # 粗粒度归属会盖住细粒度归属，成文时说不出是哪个小区。
            if main_key >= 0:
                for i in main_cols:
                    if i >= len(rows[main_key]):
                        continue
                    h = rows[main_key][i]
                    if h and ("小区名称" in h or "项目名称" in h or "具体地址" in h):
                        id_cols.append(i)
                        break

            for gi, (key, cols) in enumerate(merged_groups):
                if not cols:
                    continue
                name = ws.title if gi == 0 else "{}-{}".format(ws.title, gi + 1)
                head_cells = []
                own_cols = list(cols)
                if gi > 0 and id_cols:
                    # 子表：标识列在前，自己的列在后
                    own_cols = id_cols + [i for i in cols if i not in id_cols]
                    head_cells = [" | ".join(rows[main_key][i] if main_key >= 0 else ""
                                             for i in id_cols)]
                # 子表的列名是**上下两行**：主表列名行（第 1 行）写「1 立面材料修缮或翻新」，
                # 子表列名行（第 2 行）只写「工程量（㎡）」。只取子表那一行，项目名就丢了，
                # 下游只能按列序猜列——实测把雨棚当楼道、把空调披水板当消防设施。
                # 这里把两行拼成一行真实表头：项目名（单位），原文照录、不改写。
                sub_names = []
                for i in cols:
                    v = rows[key][i] if key >= 0 else ""
                    if key >= 0 and main_key >= 0 and i < len(rows[main_key]):
                        main_v = re.sub(r"\s+", " ", rows[main_key][i] or "").strip()
                        v_flat = re.sub(r"\s+", " ", v or "").strip()
                        unit = re.search(r"[（(]([^）)]+)[）)]", v_flat)
                        if main_v and unit and "工程量" in v_flat and "工程量" not in main_v:
                            v = "{}（{}）".format(main_v, unit.group(1))
                    sub_names.append(v)
                # 大标题**单独成块**，不能与列名拼在同一行：
                # 拼在一起会让列名整体右移（实测右移 2 格），
                # 于是"户数"对上了建筑面积的格子，抽出的事实列名全错。
                if title_row is not None and gi == 0:
                    out.append("表（{}）标题：{}".format(
                        name, " | ".join(c for c in rows[title_row] if c)))
                head_cells.append(" | ".join(sub_names))
                out.append("表（{}）表头：{}".format(name, " | ".join(head_cells)))
                # 数据行：列名行之后的所有行（各行只取本表的列区间）
                start = key + 1 if key >= 0 else 0
                for ri in range(start, len(rows)):
                    r = rows[ri]
                    vals = [r[i] for i in own_cols]
                    if not any(vals):
                        continue
                    out.append("表（{}）数据行：{}".format(name, " | ".join(vals)))
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return out


def _is_pure_num(text: str) -> bool:
    """整格是不是纯数值。"""
    return bool(re.match(r"^\d+(?:\.\d+)?$", str(text or "").strip()))


def _looks_like_column_name(text: str) -> bool:
    """像不像列名（含中文说明、不是纯数字）。"""
    t = str(text or "").strip()
    if not t or _is_pure_num(t):
        return False
    return bool(re.search(r"[一-鿿]", t))

def blocks_from_text(path: Path) -> list[str]:
    try:
        text = path.read_text(encoding="utf-8", errors="ignore")
    except OSError:
        return []
    out = []
    for line in text.splitlines():
        line = line.strip().strip("#").strip()
        if len(line) >= 8:
            out.append(line)
    return out


def blocks_from_pptx(path: Path) -> list[str]:
    """PPTX 幻灯片文本：每页一个文本块，保留页内分组。

    前期沟通 PPT 常是材料里信息量最大的一份（改造范围、管线、人居环境整治），
    此前根本不在白名单里，等于整份材料被丢弃——补上。
    """
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            names = sorted(
                (n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                key=lambda n: int(re.search(r"slide(\d+)\.xml$", n).group(1)))
            out = []
            for n in names:
                xml = z.read(n).decode("utf-8", "ignore")
                texts = [re.sub(r"<[^>]+>", "", t).strip()
                         for t in re.findall(r"<a:t>(.*?)</a:t>", xml, re.S)]
                texts = [t for t in texts if t]
                if texts:
                    out.append("幻灯片：" + " | ".join(texts))
            return out
    except (zipfile.BadZipFile, OSError):
        return []


def _ocr_engine() -> str:
    """定位 OCR 引擎标识；没有则返回空串（图片跳过，不报错）。"""
    try:
        import rapidocr_onnxruntime  # noqa: F401
        return "rapidocr"
    except ImportError:
        pass
    import shutil as _sh
    return _sh.which("tesseract") or ""


def blocks_from_image(path: Path) -> list[str]:
    """图片 OCR → 文本块。

    项目信息卡、现状照片上的文字是现状分析的第一手事实——
    实测「项目信息.png」里写着业主单位、实施周期、资金来源，
    整张图不进汇编，「项目概况」节就完全无材料。
    无 OCR 引擎时返回空（不报错）：图片只是可选来源，缺了不影响其它材料。
    """
    engine = _ocr_engine()
    if not engine:
        return []
    try:
        if engine == "rapidocr":
            from rapidocr_onnxruntime import RapidOCR
            res, _ = RapidOCR()(str(path))
            texts = [str(item[1]) for item in (res or []) if len(item) > 1 and str(item[1]).strip()]
        else:
            import subprocess
            proc = subprocess.run([engine, str(path), "stdout", "-l", "chi_sim"],
                                  capture_output=True, text=True,
                                  encoding="utf-8", errors="replace")
            texts = [ln.strip() for ln in (proc.stdout or "").splitlines() if ln.strip()]
    except Exception:
        return []
    if not texts:
        return []
    # 一张图出一个块：OCR 结果是碎片，分行成块会被打散到不同章节
    return ["图片文字（{}）：{}".format(path.name, " ".join(texts))]


def bigrams(text: str) -> set[str]:
    cleaned = re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9]", "", text)
    return {cleaned[i:i + 2] for i in range(len(cleaned) - 1)} if len(cleaned) > 1 else set()


def score(block: str, title: str) -> float:
    """材料块与章节标题的贴合度：二元组重叠数 / 标题二元组数。"""
    t = bigrams(title)
    if not t:
        return 0.0
    return len(bigrams(block) & t) / len(t)


def build(template: Path, input_dir: Path, out_path: Path, keywords=None,
          exclude_keywords=None) -> dict:
    outline = read_template_outline(template)
    if not outline:
        raise SystemExit("[assemble] 模板里没读到 H1/H2，无法汇编")

    sections = []          # [(level, title, [blocks])]
    for level, title in outline:
        if level == 1 or not sections:
            sections.append([level, title, []])
        else:
            sections.append([level, title, []])
    h2_sections = [s for s in sections if s[0] == 2]
    bucket = {id(s): [] for s in sections}

    # 子节 → 父章（H2 归属其前最近的 H1），用于 H1 覆盖度回填
    section_children: dict[str, str] = {}
    current_h1 = ""
    for level, title, _ in sections:
        if level == 1:
            current_h1 = title
        else:
            section_children[title] = current_h1

    sources = []
    # 表键 → 已定节标题（同一张表的表头与数据行必须同节）
    # 键用「文件名::表名」：不同文件的表可能都叫 "Sheet1"，
    # 只按表名缓存会让后一个文件的表继承前一个文件的归属节。
    tbl_route: dict[str, str] = {}
    for path in iter_candidate_files(input_dir, keywords, exclude_keywords):
        suffix = path.suffix.lower()
        try:
            if suffix == ".docx":
                blocks = blocks_from_docx(path)
            elif suffix == ".pptx":
                blocks = blocks_from_pptx(path)
            elif suffix in SHEET_SUFFIX:
                blocks = blocks_from_xlsx(path)
            elif suffix in IMAGE_SUFFIX:
                blocks = blocks_from_image(path)
            else:
                blocks = blocks_from_text(path)
        except Exception:
            continue
        if blocks:
            sources.append({"file": path.name, "blocks": len(blocks)})
        for block in blocks:
            if len(block) < 10:
                continue
            # 表格块走**表名语义路由**，正文块走二元组打分。
            # 表格的表名（如"现状基本情况"）与节标题措辞往往对不上，
            # 靠打分会被派到偶然相似的节里；同一张表的表头块与数据行块
            # 必须落进同一节，故用 tbl_route 记住该表已定的节。
            tname = ""
            tm = re.match(r"^表（([^）]*)）(表头|数据行)：", block)
            if tm:
                tname = tm.group(1)
            if tname:
                tkey = "{}::{}".format(path.name, tname)
                if tkey in tbl_route:
                    want = tbl_route[tkey]
                else:
                    want = route_table_section(tname, h2_sections)
                    if not want:
                        # 表名路由落空（工作表名常是 "Sheet1"）→ 用**表头列名**判。
                        # 实测改造小区统计表整张落到第一节，
                        # 「建筑现状与风貌保护现状」这类最该用它的章节反而空着。
                        want = route_by_header(block, h2_sections)
                    tbl_route[tkey] = want
                target = next((sec for sec in h2_sections if sec[1] == want), None) if want else None
                if target is None:
                    target = h2_sections[0] if h2_sections else sections[0]
            else:
                # 图片文字（项目信息卡）多写的是**项目级**信息：
                # 业主单位、实施周期、资金来源、面积总量 ——
                # 这些属于「项目概况」，不是靠二元组打分能命中的。
                if block.startswith("图片文字（"):
                    want = ""
                    for kw in ("项目概况", "项目概述"):
                        want = next((sec[1] for sec in h2_sections if sec[1] == kw), "")
                        if want:
                            break
                    target = next((sec for sec in h2_sections if sec[1] == want), None) if want else None
                    if target is None:
                        target = h2_sections[0] if h2_sections else sections[0]
                    bucket[id(target)].append((path.name, block))
                    continue
                best, best_score = None, 0.0
                for sec in h2_sections:
                    sc = score(block, sec[1])
                    if sc > best_score:
                        best, best_score = sec, sc
                target = best if best_score > 0 else (h2_sections[0] if h2_sections else sections[0])
            bucket[id(target)].append((path.name, block))

    doc = docx.Document()
    normal = doc.styles["Normal"]
    normal.font.name = "宋体"
    normal.font.size = Pt(10.5)

    # 覆盖度清单：逐节记录「是否真有材料」，供 04 章节映射判定真实缺口。
    # 没有这份清单，04 只能按标题匹配——而汇编产物的标题就是模板标题，
    # 会把「无料节」判成 exact，缺口版永远不会启用。
    coverage = []

    for sec in sections:
        level, title, _ = sec
        doc.add_heading(title, level=level)
        items = bucket.get(id(sec)) or []
        if not items:
            # 结构化标记：正文里显式说明本节无材料，04 可据此判 gap
            doc.add_paragraph(GAP_MARK)
            coverage.append({"title": title, "level": level, "hasMaterial": False,
                             "blockCount": 0, "chars": 0, "sources": [],
                             # 无 H2 子节的 H1（如「核心问题汇总」「附件」）：内容由流程段/
                             # 模板外附加节生成，不算材料缺口
                             "structural": level == 1 and not section_children.get(title)})
            continue
        kept = items[:40]
        for name, block in kept:
            doc.add_paragraph("【来源：{}】{}".format(name, block))
        coverage.append({
            "title": title, "level": level, "hasMaterial": True,
            "blockCount": len(kept), "chars": sum(len(b) for _n, b in kept),
            "sources": sorted({n for n, _b in kept}), "structural": False,
        })

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    covered = sum(1 for c in coverage if c["hasMaterial"])
    # H1 章标题本身不承载材料，其「有料」由子节决定：
    # 章下所有子节都无料 → 该章判为无料（否则 3 个 H1 恒为空，虚增缺口数）。
    by_title = {c["title"]: c for c in coverage}
    h2_titles = [c["title"] for c in coverage if c.get("level") == 2]
    for c in coverage:
        if c.get("level") != 1:
            continue
        kids = [by_title[t] for t in h2_titles
                if section_children.get(t, "").strip() == c["title"]]
        if kids:
            c["hasMaterial"] = any(k["hasMaterial"] for k in kids)
            c["blockCount"] = sum(k["blockCount"] for k in kids)
            c["chars"] = sum(k["chars"] for k in kids)
            c["childSections"] = [k["title"] for k in kids]
    covered = sum(1 for c in coverage if c["hasMaterial"])
    cov_path = out_path.with_suffix(out_path.suffix + ".覆盖.json")
    with open(cov_path, "w", encoding="utf-8") as f:
        json.dump({"schema": "cardo/assemble-coverage/v1",
                   "sections": coverage,
                   "summary": {"total": len(coverage), "covered": covered,
                               "empty": len(coverage) - covered}},
                  f, ensure_ascii=False, indent=1)

    return {"sections": len(coverage), "covered": covered,
            "coverage": str(cov_path),
            "sources": sources, "out": str(out_path)}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--project", default="", help="项目全称；用于按项目边界过滤材料")
    ap.add_argument("--other-projects", default="",
                    help="投递区里其它项目的全称（逗号分隔）；用于把它们的材料排除在本项目之外")
    ap.add_argument("--exclude-json", default="",
                    help="形式参考件登记 JSON；excluded 列出的文件不进材料汇编")
    args = ap.parse_args(argv)
    if args.exclude_json:
        try:
            _d = json.loads(Path(args.exclude_json).read_text(encoding="utf-8"))
            for _x in (_d.get("excluded") or []):
                _EXCLUDE_ABS.add(str(Path(args.input, _x).resolve()).replace("\\", "/"))
            for _r in (_d.get("reference") or []):
                if _r.get("path"):
                    _EXCLUDE_ABS.add(str(Path(_r["path"]).resolve()).replace("\\", "/"))
        except Exception as _e:
            print("[assemble] 形式参考登记读取失败：{}".format(_e), file=sys.stderr)
            return 2

    input_dir = Path(args.input)
    template = Path(args.template)
    if not input_dir.is_dir():
        print("[assemble] 输入目录不存在：{}".format(input_dir), file=sys.stderr)
        return 2
    if not template.is_file():
        print("[assemble] 模板不存在：{}".format(template), file=sys.stderr)
        return 2

    specific, generic = project_keywords(args.project)
    total = sum(1 for _ in iter_candidate_files(input_dir, None))
    keywords = specific
    matched = sum(1 for _ in iter_candidate_files(input_dir, keywords)) if keywords else 0
    if keywords and matched == 0:
        # 专属词一个没命中：放宽到方向词再试一次（材料目录名可能只写了方向）
        relaxed = specific + generic
        rmatch = sum(1 for _ in iter_candidate_files(input_dir, relaxed))
        if rmatch:
            print("[assemble] 专属词 {} 未命中，放宽到 {} 后命中 {} / {} 份材料"
                  .format("、".join(specific), "、".join(relaxed), rmatch, total))
            keywords, matched = relaxed, rmatch
        else:
            print("[assemble] 按项目关键词 {} 未命中任何材料，改用全部材料（请核对项目名是否正确）"
                  .format("、".join(specific + generic)))
            keywords = []
    elif keywords:
        print("[assemble] 项目边界过滤：专属词 {}，命中 {} / {} 份材料".format(
            "、".join(keywords), matched, total))

    # 其它项目的排他词：内容级判定时，命中它们就说明材料主体属于别个项目
    exclude = []
    for other in [x.strip() for x in (args.other_projects or "").split(",") if x.strip()]:
        ospec, _ = project_keywords(other)
        exclude.extend(ospec)
    exclude = [w for w in exclude if w not in keywords]

    info = build(template, input_dir, Path(args.out), keywords, exclude)
    print("[assemble] 汇编完成：{} 节，其中 {} 节有材料；来源 {} 个文件 → {}".format(
        info["sections"], info["covered"], len(info["sources"]), info["out"]))
    print("[assemble] " + json.dumps(info["sources"], ensure_ascii=False)[:400])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
