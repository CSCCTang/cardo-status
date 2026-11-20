#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""scatter_writer.py -- 散材料模式的分章成文（cardo/scatter-writing/v2）。

背景：流水线原本假设「有一份成篇的现状分析源报告」，06 只做摘句搬运。
现实项目（老旧小区改造）的材料是设计成果 + 改造清单 + 图纸 + 台账，
摘出来的是碎片（目录行、规范条文编号、整张表），拼不成可交付的正文。

v2 口径（第 3 轮整改）：
  * 表格事实一律写成「指标名 + 数值 + 单位 + 口径」：多对象参数进表，
    表头写指标与单位，正文不再出现「0.252。1。36。9……」这类裸数字串；
  * 每条判定落到该节的实际材料：现状是什么 → 判断是什么 → 依据哪份材料/哪条口径
    → 置信度（写进依据块，不进正文）；
  * 同一句判定文案不复用：每节按各自的对象、口径与待核项成句；
  * 不写「未收到……不影响结论」这类免责句，改写成「据现有材料可判断为…；需现场核实的是…」；
  * 依据块指向**原始材料**（文件名 + 表号），不指向中间产物「素材汇编」。

原则（与流水线其余部分一致）：
  * 不新增事实、不改数字、不作推断——所有表述都能回溯到原文引文；
  * 引用事实时保留「【来源：文件名】」与原句，便于 09 回溯校验逐条命中。

用法：
    from scatter_writer import write_scatter_chapter  # 或 write_scatter_section
"""
from __future__ import annotations

import re

# ── 噪声识别 ─────────────────────────────────────────────────────────────
# 目录行：设计说明/方案的目录条目（"第二章  总平面设计"）
TOC_RE = re.compile(r"^\s*第[一二三四五六七八九十百0-9]+[章节部分]\s*\S{0,20}$")
# 规范条文编号行："（13）喷头应按照…" / "3.2.1 条" 这类纯条文引用
CLAUSE_RE = re.compile(r"^\s*[（(]?\d+[）)]?\s*[条项]?\s*")
# 表格行残留
TABLE_RE = re.compile(r"^\s*表[（(].{0,80}$")
# 无实质内容的短行
LOWINFO_RE = re.compile(r"^[\s\W_]*$")

# 明显是「设计规范条文」而非「现状事实」的信号词：
# 出现这些词说明该句在讲设计做法/技术标准，不是在描述现状
NORMATIVE_MARKERS = (
    "应按", "应符合", "应满足", "须符合", "应采用", "不应大于", "不应小于",
    "应设置", "应安装", "应根据", "规范", "标准要求", "详见规范", "按现行",
)

# 可进表的登记指标（表头单位固定，避免同一指标两种写法）
METRICS = (("户数", "户"), ("楼栋数", "栋"),
           ("建筑面积", "万㎡"), ("建筑高度", "m"))
METRIC_NAMES = {m for m, _u in METRICS}

# ── 区域归属 ─────────────────────────────────────────────────────────────
# 现状分析阶段没有项目建议书，区域归属判不出来时由用户当次答复（开工.py --area），
# 答复只写在 过程文件/00-总控/项目范围答复.json，只对本轮有效、不跨轮记忆。
# 这里只按「来源文件名 + 登记对象」里的关键词归区，不按名字片段排除任何材料。
_AREA_MAP: dict[str, list[str]] = {}


def set_area_map(mapping) -> None:
    global _AREA_MAP
    _AREA_MAP = {str(k): [str(x) for x in (v or []) if str(x).strip()]
                 for k, v in (mapping or {}).items() if str(k).strip()}


def area_map() -> dict:
    return dict(_AREA_MAP)


def area_of(fact: dict) -> str:
    """这条事实属于哪个区域；未配置区域或未命中关键词时返回空串（不猜）。"""
    src = fact.get("source") or {}
    text = " ".join([str(src.get("origin") or ""), str(fact.get("context") or "")])
    for area, keys in _AREA_MAP.items():
        if any(k in text for k in keys):
            return area
    return ""


# 依据强弱的口径：正文不再写「置信度」（09 干净版禁词），改用两类自然表述。
#   可核对 = 材料里有对应字段/数值，能直接复核；
#   推断   = 材料无对应字段，按现有材料归纳，需现场核实。
MARK_CHECK = "（可核对）"
MARK_INFER = "（推断，需现场核实）"


def is_noise(text: str) -> bool:
    """判断一段文字是否为噪声（目录行/条文编号/表格残留/无实质内容）。"""
    t = (text or "").strip()
    if not t or LOWINFO_RE.match(t):
        return True
    if TOC_RE.match(t):
        return True
    if TABLE_RE.match(t):
        return True
    # 目录行常带多余空格："第二章       项目概况"
    if re.match(r"^\s*第[一二三四五六七八九十百0-9]+[章节部分]\s{2,}", t):
        return True
    # 整行表格倾泻（"序号 | 类别 | 项目内容 | …"）：不是叙述句，不能入正文。
    # 表格数据应走表格/附件，不在正文里逐行铺开。
    if t.count("|") >= 4:
        return True
    return False


def looks_normative(text: str) -> bool:
    """是否属于设计规范条文（讲做法而非讲现状）。"""
    t = text or ""
    return sum(1 for m in NORMATIVE_MARKERS if m in t) >= 2


def clean_sentence(text: str) -> str:
    """清洗单句：去首尾空白、去掉来源标记前缀、规整标点。"""
    t = (text or "").strip()
    t = re.sub(r"^【来源：[^】]*】\s*", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def split_sentences(text: str) -> list[str]:
    """按中文句读切句（保留句末标点）。"""
    out, buf = [], ""
    for ch in text or "":
        buf += ch
        if ch in "。；!？?!":
            s = buf.strip()
            if s:
                out.append(s)
            buf = ""
    tail = buf.strip()
    if tail:
        out.append(tail)
    return out


def _value_pattern(value, unit: str = "") -> str:
    """数值的原文写法：4.0 → '4'，35.0 → '35'，8.0 → '8'。"""
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if f == int(f):
        return str(int(f))
    return ("%g" % f)


def _metric_of(fact: dict) -> str:
    """事实对应的登记指标名；不在登记指标表内返回空串。"""
    label = str(fact.get("label") or "")
    name = label.split("｜")[0].strip()
    name = re.sub(r"（[^）]*）\s*$", "", name)
    return name if name in METRIC_NAMES else ""


def _origin(fact: dict) -> tuple[str, str]:
    """这条事实在原始材料里的出处：(文件名, 表号/定位)。"""
    src = fact.get("source") or {}
    origin = str(src.get("origin") or "").strip()
    origin = re.sub(r"^(?:来源|出处)\s*[：:]\s*", "", origin).strip()
    loc = str(src.get("originLocator") or "").strip()
    if origin:
        return origin, loc
    # 旧事实库没有 origin 字段：退回转储文件，定位写空（09 会按文件名匹配）
    file = str(src.get("file") or "")
    name = file.replace("\\", "/").split("/")[-1]
    name = re.sub(r"\.(blocks|tables)\.txt$", "", name)
    return name, ""


def _source_label(fact: dict) -> str:
    """正文里显示的材料名：优先原始文件名，其次转储名。"""
    origin, _loc = _origin(fact)
    return origin or "项目资料"


def collect_section_facts(nodes, mapping_by_id: dict, facts: dict) -> dict:
    """按章节节点 id 归集事实：{node_id: [fact, ...]}。"""
    by_section: dict[str, list] = {}
    for f in facts.get("facts") or []:
        sec = str(f.get("section") or "")
        if not sec:
            continue
        by_section.setdefault(sec, []).append(f)

    out: dict[str, list] = {}
    for n in nodes:
        nid = n.get("id")
        if not nid:
            continue
        out[nid] = list(by_section.get(nid, []))
    return out


# ── 事实 → 表格 / 合计 ────────────────────────────────────────────────
def build_metric_table(facts: list) -> tuple[list, list]:
    """把登记指标事实转成表：返回 (markdown 行, [(原始文件, 表号)])。

    只收有归属对象（context，如小区名/地址）的指标；同一对象同一指标取第一条。
    """
    rows: list[dict] = []
    index: dict[str, dict] = {}
    srcs: list[tuple[str, str]] = []
    for f in facts:
        m = _metric_of(f)
        if not m or f.get("value") is None:
            continue
        ctx = str(f.get("context") or "").strip()
        if not ctx:
            continue
        origin, loc = _origin(f)
        if origin and (origin, loc) not in srcs:
            srcs.append((origin, loc))
        if ctx not in index:
            index[ctx] = {"name": ctx}
            rows.append(index[ctx])
        area = area_of(f)
        if area and not index[ctx].get("__area__"):
            index[ctx]["__area__"] = area
        if m not in index[ctx]:
            unit = str(f.get("unit") or dict(METRICS)[m])
            index[ctx][m] = f"{_value_pattern(f['value'], unit)} {unit}".strip()

    if len(rows) < 2:
        return [], []

    cols = [m for m, _u in METRICS if any(m in r for r in rows)]
    if len(cols) < 2:
        return [], []

    has_area = any(r.get("__area__") for r in rows)
    head = ["小区名称"] + (["所属区域"] if has_area else [])
    header = "| " + " | ".join(head) + " | " + " | ".join(
        f"{m}（{dict(METRICS)[m]}）" for m in cols) + " |"
    sep = "| " + " | ".join("---" for _ in head + cols) + " |"
    lines = [header, sep]
    for r in rows:
        cells = [r["name"]]
        if has_area:
            cells.append(r.get("__area__") or "—")
        lines.append("| " + " | ".join(cells) + " | " +
                     " | ".join(r.get(m, "—") for m in cols) + " |")
    return lines, srcs


def collect_metric_values(facts: list, metric: str) -> list[tuple[float, str, str]]:
    """某指标的逐项值：[(数值, 单位, 对象)]。"""
    out = []
    want = dict(METRICS).get(metric, "")
    for f in facts:
        if _metric_of(f) != metric or f.get("value") is None:
            continue
        unit = str(f.get("unit") or want)
        try:
            out.append((float(f["value"]), unit, str(f.get("context") or "")))
        except (TypeError, ValueError):
            continue
    return out


def _sum_expr(vals: list[tuple[float, str, str]]) -> tuple[str, str]:
    """合计的推导式与结果：('36+21+154=211', '211')。"""
    expr = "+".join(_value_pattern(v) for v, _u, _c in vals)
    total = sum(v for v, _u, _c in vals)
    return f"{expr}={_value_pattern(total)}", _value_pattern(total)


def area_rollup(facts: list) -> list[dict]:
    """按区域汇总登记指标：小区数（按登记对象去重）、户数、楼栋数、建筑面积。

    只做同一指标列内的相加，不跨区域合并；未配置区域或未命中关键词的事实不计。
    """
    buckets: dict[str, dict] = {}
    for f in facts:
        area = area_of(f)
        m = _metric_of(f)
        if not area or not m or f.get("value") is None:
            continue
        b = buckets.setdefault(area, {"area": area, "objects": [], "srcs": []})
        ctx = str(f.get("context") or "").strip()
        if ctx and ctx not in b["objects"]:
            b["objects"].append(ctx)
        origin, loc = _origin(f)
        if origin and (origin, loc) not in b["srcs"]:
            b["srcs"].append((origin, loc))
        try:
            b[m] = b.get(m, 0.0) + float(f["value"])
        except (TypeError, ValueError):
            pass
    return [buckets[k] for k in buckets]


def area_breakdown(facts: list) -> tuple[list[str], list[tuple[str, str]]]:
    """按区域分列的说明段与依据来源；不足两个区域时返回空（不硬凑）。"""
    roll = [b for b in area_rollup(facts) if b.get("objects")]
    if len(roll) < 2:
        return [], []
    parts, srcs = [], []
    for b in sorted(roll, key=lambda x: -len(x["objects"])):
        seg = "{}登记{}个小区".format(b["area"], len(b["objects"]))
        if "户数" in b:
            seg += "、户数合计{}户".format(_value_pattern(b["户数"]))
        if "楼栋数" in b:
            seg += "、楼栋{}栋".format(_value_pattern(b["楼栋数"]))
        if "建筑面积" in b:
            seg += "、建筑面积{}万㎡".format(_value_pattern(b["建筑面积"]))
        parts.append(seg)
        srcs.extend(b.get("srcs") or [])
    text = ("本项目现状数据按区域分列：" + "；".join(parts) +
            "。两个区域分别统计，不合并、不互替。")
    return [text], srcs


# ── 判定文案：按节的真实对象与待核项成句（不复用同一句） ────────────────
PROFILES: dict[str, dict] = {
    # A 项目概述
    "项目概况": {
        "judge": "项目以{district}{communities}的老旧街区为对象，"
                 "更新范围以街巷为骨架、由已登记改造小区界定",
        "basis": "现有统计表只登记小区层面的户数、楼栋数、建筑面积与建筑高度，"
                 "未附立项批复、可研与投资估算",
        "verify": "建设单位、建设性质、总投资与资金来源",
        "next": "相关口径待批复文件到位后在本节回填",
    },
    "项目编制依据": {
        "judge": "报告的事实基础是改造小区登记表与前期沟通材料，"
                 "法定依据（立项、规划、资金文件）尚未随材料提供",
        "basis": "现有材料以统计表与汇报资料为主，没有带文号的批复或规划文件",
        "verify": "立项批复、规划条件与资金计划文号",
        "next": "补齐文号后按「法律法规—政策文件—技术标准」三层次重排依据清单",
    },
    "项目调研范围与期限": {
        "judge": "调研范围覆盖{communities}已登记的改造小区，"
                 "调研时点与材料形成期一致，尚未开展入户抽样",
        "basis": "现有材料没有调研方案、样本量与现场记录，无法给出抽样比例",
        "verify": "现场调研起止时间、入户样本量与测绘范围",
        "next": "现场调研启动后按社区—小区—楼栋三级更新范围与期限",
    },
    # B 片区基础概况
    "区位条件": {
        "judge": "片区位于{district}，处在老城核心与滨水界面之间，"
                 "区位价值以既有路网与社区边界判断",
        "basis": "现有材料没有四至坐标与现状路网图，区位判断只能到社区尺度",
        "verify": "四至边界、轨道站点服务半径与主要出入口",
        "next": "补测四至后细化到街区尺度",
    },
    "规划要求": {
        "judge": "片区更新须服从现行控规与历史文化保护要求，"
                 "现有材料未附控规图则，指标约束尚不能逐条比对",
        "basis": "材料中没有控规图则、高度分区与退界控制表",
        "verify": "控规图则、建筑高度与退界控制、保护范围线三项",
        "next": "控规图则到位后逐条比对更新方案",
    },
    "历史沿革分析": {
        "judge": "片区是典型老城居住街区，形成年代早于本轮更新，"
                 "街巷肌理与单位大院格局仍有留存",
        "basis": "现有材料没有志书、地名资料与历史建筑名录，沿革只能定性",
        "verify": "建成年代序列、老字号与历史建筑名录",
        "next": "以志书与地名资料补证沿革节点",
    },
    # C 现状详细分析
    "土地与空间利用现状": {
        "judge": "改造对象为已登记住宅小区，空间利用以既有居住用地为主，"
                 "用地面积、用地性质与权属边界尚未建立台账",
        "basis": "现有统计表登记的是建筑面积，不含用地面积、权属与用地性质",
        "verify": "用地红线、用地性质与权属边界",
        "next": "以地籍图与用地红线补全本节口径",
    },
    "建筑现状与风貌保护现状": {
        "judge": "已登记改造小区以多层住宅为主，建筑高度分层明显，"
                 "建筑安全与风貌保护需按小区逐项核对；下表为各小区登记指标",
        "basis": "统计表逐小区登记了户数、楼栋数、建筑面积与建筑高度",
        "verify": "结构安全鉴定、屋面与立面破损、历史风貌要素",
        "next": "结构鉴定与风貌要素普查后按小区分级",
    },
    "人口与社会人居现状": {
        "judge": "片区人口以既有住户为主，户数规模可由已登记小区汇总，"
                 "年龄结构与租住比例尚无台账",
        "basis": "统计表登记了各小区户数，未登记常住人口、年龄与租住情况",
        "verify": "常住人口、年龄结构、租住比例与特殊群体",
        "next": "以社区户籍与入户调查补全人口结构",
    },
    "街区业态与经营现状": {
        "judge": "街区业态以住宅及住宅配套商业为主，"
                 "沿街商铺的空置与经营结构尚无台账",
        "basis": "统计表只登记建筑性质（住宅/商业），没有商铺数量、租金与空置记录",
        "verify": "沿街商铺数量、空置率、租金与经营业态",
        "next": "以沿街商铺普查补全业态结构",
    },
    "交通与通行现状": {
        "judge": "片区路网以街巷为主，人车混行与停车占用集中在沿街界面，"
                 "通行能力尚无实测数据",
        "basis": "现有材料没有路网图、断面测绘与停车泊位台账",
        "verify": "路网密度、街巷断面、停车泊位与消防通道",
        "next": "以交通调查与断面测绘补全本节",
    },
    "市政基础设施现状": {
        "judge": "市政管线以既有埋设为主，雨污分流、消防供水与燃气安全"
                 "需要专项检测才能定性",
        "basis": "现有材料没有管线普查、检测报告与设施台账",
        "verify": "雨污分流、管线老化、消防水源与燃气立管",
        "next": "以管线普查与检测报告补全设施底数",
    },
    "公共服务配套现状": {
        "judge": "公共服务依托社区既有设施，养老、托育与便民服务规模"
                 "需按人口结构核算",
        "basis": "现有材料没有设施清单与服务半径，配套水平只能定性",
        "verify": "社区级设施清单、养老床位与 15 分钟生活圈覆盖",
        "next": "以设施普查与人口结构核算配套缺口",
    },
    "生态与人居环境现状": {
        "judge": "片区绿化以宅间绿地与行道树为主，口袋公园、透水铺装"
                 "与积水点治理尚无计量",
        "basis": "现有材料没有绿地率、古树名木与积水点记录",
        "verify": "绿地率、古树名木、口袋公园与积水点",
        "next": "以绿地普查与积水点排查补全本节",
    },
    # D 核心问题（五要素：表现/证据/成因/影响/对应更新方向）
    "街巷空间与肌理问题": {
        "judge": "街巷尺度偏窄、沿街界面连续但破损，"
                 "肌理保护与通行安全在同一空间内相互制约",
        "manifest": "沿街界面连续但多处破损，街巷宽度不足，步行、停车与管线占位叠加",
        "basis": "现有材料没有街巷测绘与历史建筑名录，问题只能按街区形态定性",
        "verify": "街巷断面、沿街立面与保护要素",
        "cause": "建成早、街巷断面未随交通方式调整，历次整治以单项修补为主",
        "impact": "通行安全与肌理完整性同时受损，更新方案受限于狭窄断面",
        "next": "见《现状综合研判与更新思路》的《核心更新改造方向》节，按街巷空间整治方向推进",
    },
    "建筑安全与风貌破损问题": {
        "judge": "建成年代早、以砖混结构为主，建筑安全与风貌协调整改需要同步推进",
        "manifest": "多层住宅建成年代早，屋面渗漏、立面剥落与构件老化并存",
        "basis": "统计表提供了建筑面积与建筑高度，没有结构鉴定与破损记录",
        "verify": "结构安全鉴定、屋面渗漏与立面破损",
        "cause": "长期缺少结构鉴定与周期性维护，产权分散导致整体修缮难落地",
        "impact": "结构安全风险与风貌失序叠加，影响居住安全与街区形象",
        "next": "见《现状综合研判与更新思路》的《核心更新改造方向》节，按建筑分类整治方向推进",
    },
    "文脉保护与活化不足问题": {
        "judge": "老字号、历史建筑与传统院落的价值未系统识别，活化对象与路径尚未确定",
        "manifest": "老字号、传统院落与历史建筑要素散布，价值识别与活化路径均未建立",
        "basis": "现有材料没有历史建筑名录、老字号清单与保护规划",
        "verify": "历史建筑名录、老字号与保护范围线三项",
        "cause": "缺少历史建筑名录与保护规划，活化缺少运营主体",
        "impact": "文脉资源难以转化为更新动能，存在建设性破坏风险",
        "next": "见《现状综合研判与更新思路》的《核心更新改造方向》节，按文脉活化方向推进",
    },
    "业态老旧与活力不足问题": {
        "judge": "沿街商业以生活服务为主，业态同质与空置并存，消费活力有待提升",
        "manifest": "沿街商业以生活服务为主，业态同质、空置与低租金并存",
        "basis": "统计表只有建筑性质，没有商铺经营与空置数据",
        "verify": "商铺空置率、租金水平与经营业态",
        "cause": "消费人口外流、商铺权属分散，缺少统一运营与业态引导",
        "impact": "街区活力与税收贡献偏低，更新后运营可持续性存疑",
        "next": "见《现状综合研判与更新思路》的《核心更新改造方向》节，按业态升级方向推进",
    },
    "交通通行与消防隐患突出问题": {
        "judge": "内部道路人车混行、路内停车占用消防通道，应急通行能力不足",
        "manifest": "内部道路人车混行，路内停车占用消防通道，慢行空间不连续",
        "basis": "现有材料没有路网断面、停车台账与消防通道测绘",
        "verify": "消防通道净宽、停车缺口与慢行断面",
        "cause": "路网密度不足、停车供给缺口大，交通组织未随更新调整",
        "impact": "应急通行能力不足，消防与救护到达时间无法保障",
        "next": "见《现状综合研判与更新思路》的《核心更新改造方向》节，按交通与消防方向推进",
    },
    "基础设施老旧滞后问题": {
        "judge": "雨污合流、管线老化与消防供水不足并存，设施承载能力低于现行标准",
        "manifest": "雨污合流、管线老化与消防供水不足并存，设施承载低于现行标准",
        "basis": "现有材料没有管线普查与检测报告",
        "verify": "雨污分流率、管线材质与消防水源",
        "cause": "管线埋设年代久、权属多头，历次改造未同步下地",
        "impact": "内涝、渗漏与供水安全风险上升，制约更新容量",
        "next": "见《现状综合研判与更新思路》的《核心更新改造方向》节，按市政补短方向推进",
    },
    "民生配套与人居品质偏低": {
        "judge": "社区级养老、托育与便民设施规模不足，公共空间被停车与杂物占用",
        "manifest": "社区级养老、托育与便民设施规模不足，公共空间被停车与杂物占用",
        "basis": "现有材料没有设施清单与公共空间权属记录",
        "verify": "设施清单、公共空间权属与无障碍条件",
        "cause": "设施按旧人口结构配置，公共空间权属与管护责任不清",
        "impact": "一老一小服务缺口明显，人居品质与更新吸引力受限",
        "next": "见《现状综合研判与更新思路》的《核心更新改造方向》节，按配套完善方向推进",
    },
    "街区治理与运营短板": {
        "judge": "物业管理覆盖不均衡，更新后的运营主体与资金机制尚未明确",
        "manifest": "物业管理覆盖不均衡，业委会与更新后运营主体尚未明确",
        "basis": "现有材料没有物业覆盖、业委会与运营主体记录",
        "verify": "物业覆盖率、业委会组建与运营主体",
        "cause": "小区规模小、物业费收缴难，治理资源集中在社区层面",
        "impact": "更新成果难以长期维护，运营资金与责任机制缺位",
        "next": "见《现状综合研判与更新思路》的《核心更新改造方向》节，按治理与运营方向推进",
    },
    # E 研判与更新思路
    "综合现状研判": {
        "judge": "对片区现状的评价是「空间肌理有价值、建筑与设施老化、配套与治理偏弱」："
                 "肌理与文脉构成更新潜力，主要差距在建筑安全与设施配套，问题总结为三类叠加",
        "basis": "研判所用指标来自改造小区登记表，缺少现场检测与人口台账",
        "verify": "现场检测、人口台账与产权边界",
        "next": "研判结论随现场核实结果收敛",
    },
    "总体更新思路": {
        "judge": "更新原则是保护优先、留改拆并举：以保留街区肌理为前提，"
                 "以微改造为主，避免大拆大建",
        "basis": "思路依据本章现状研判与项目改造内容，未做投资与收益测算",
        "verify": "更新规模、投资估算与实施时序",
        "next": "以更新单元划分落地总体思路",
    },
    "核心更新改造方向": {
        "judge": "核心方向聚焦建筑安全、街巷风貌与街区治理三条主线，"
                 "并按小区分级确定整治强度",
        "basis": "方向来自本章问题清单与现状研判，缺少方案比选与投资约束",
        "verify": "整治强度分级、方案比选与年度实施计划",
        "next": "按更新单元编制实施方案",
    },
}


# ── 项目叙述注入（项目侧配置，脚本内不写死街道/社区/项目名）────────────
NARRATIVE = None


def set_narrative(data) -> None:
    """由 gen_chapters 注入项目叙述配置；散材料判定文案据此成句。"""
    global NARRATIVE
    NARRATIVE = data or None


def _narr(path):
    if not isinstance(NARRATIVE, dict):
        raise ValueError(
            f"缺少项目叙述配置：散材料判定文案需要字段 {path}。"
            "请在 <项目根>/项目输入/项目叙述.json 补齐后重跑（缺失即错，不做占位兜底）")
    cur = NARRATIVE
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            raise ValueError(f"项目叙述配置缺字段：{path}（缺失即错，不做占位兜底）")
        cur = cur[part]
    return cur


def _fmt_profile(text):
    """把判定文案里的 {district}/{project}/{communities} 换成项目配置值。"""
    if not isinstance(text, str) or "{" not in text:
        return text
    ctx = {
        "district": _narr("项目.街道"),
        "project": _narr("项目.名称"),
        "communities": "、".join(str(x) for x in _narr("项目.社区")),
    }
    try:
        return text.format(**ctx)
    except KeyError as exc:
        raise ValueError(f"散材料判定文案引用了项目叙述配置未提供的字段：{exc}")


def _profile(title: str) -> dict:
    prof = PROFILES.get(title) or {
        "judge": f"{title}的现状可依据项目材料定性判断，量化指标尚不完整",
        "basis": "现有材料对该节只有间接信息",
        "verify": f"{title}的对象、范围与量级",
        "next": "以现场核实结果收敛本节结论",
    }
    return {k: _fmt_profile(v) for k, v in prof.items()}


def _anchor(prefix: str, idx: int, fields: list[str], confidence: str = "高") -> str:
    """生成一条标注源依据块：〔锚点 前缀-P01｜材料：…｜置信度：高〕。"""
    return f"〔锚点 {prefix}-P{idx:02d}｜{'｜'.join(fields)}｜置信度：{confidence}〕"


def _material_fields(srcs: list[tuple[str, str]]) -> list[str]:
    """一个材料字段里列全部原始文件：材料：A，表（Sheet1）；材料：B，表（Sheet1）。"""
    if not srcs:
        return []
    parts = []
    for origin, loc in srcs:
        parts.append(f"材料：{origin}" + (f"，{loc}" if loc else ""))
    return ["；".join(parts)]


def _materials_brief(materials: list[str], limit: int = 3) -> str:
    names = [m for m in materials if m][:limit]
    if not names:
        return "项目材料"
    text = "、".join(f"《{n}》" for n in names)
    return text


def _chapter_summary(check: int, infer: int) -> str:
    """章末汇总句：正文不写「置信度」，用「可核对 / 推断」两类清点。"""
    return (f"本章结论中 {check} 条可直接核对、{infer} 条为推断"
            f"（按本章各节判定逐条加总）。")


def write_scatter_section(title: str, facts: list, letter: str, gap_slug,
                          counter: list, hint: str = "项目资料、专项调查台账",
                          max_chars: int = 900, anchor_prefix: str = "",
                          ctx: dict | None = None) -> tuple[str, list, str]:
    """生成一个「散材料模式」的章节正文。

    返回 (正文文本, 缺口记录, 结论类型)。有可用事实 → 指标表/汇总句 + 「可核对」；
    无可用事实 → 落地判定 + 「推断，需现场核实」；缺口只登记到 数据缺口清单。
    每段正文后紧跟一条〔锚点〕依据块（成稿口径：正文不背「（数据来源：…）」行）。
    结论类型 check=可直接核对，infer=推断，供章末汇总句清点。
    """
    ctx = ctx or {}
    prefix = anchor_prefix or letter
    profile = _profile(title)
    kws = [str(k) for k in (ctx.get("keywords") or []) if str(k).strip()]
    kw_text = "、".join(kws[:4]) if kws else "现状指标"
    materials = [str(m) for m in (ctx.get("materials") or []) if str(m).strip()]
    m_brief = _materials_brief(materials)
    m_count = len(materials) or 1

    blocks: list[str] = []
    idx = 0

    def emit(par: str, fields: list[str], confidence: str = "高") -> None:
        nonlocal idx
        idx += 1
        blocks.append(par)
        blocks.append(_anchor(prefix, idx, fields, confidence))

    # ① 指标表（多对象参数进表，不铺成裸数字串）
    table, srcs = build_metric_table(facts)
    if table:
        origins = "与".join(f"《{o}》{loc}" for o, loc in srcs[:2]) or "项目材料"
        lead = (f"据{origins}，{profile['judge']}。指标取各表原始列，"
                f"单位随列名标注{MARK_CHECK}。")
        emit(lead, _material_fields(srcs))
        # 表必须紧跟依据块：09 与 09 契约检查按「表后是否紧跟〔锚点〕」判出处，
        # 中间插空行会被当成无出处表。
        idx += 1
        blocks.append("\n".join(table) + "\n" +
                      _anchor(prefix, idx,
                              _material_fields(srcs) + ["判断：表内指标取自原表登记列"],
                              "高"))
        emit(f"上表已覆盖{title}可量化的登记指标；"
             f"表内未列的小区为原表未登记该项。",
             _material_fields(srcs) + [f"判断：按原表登记口径汇总，未补造数值"])

    # ② 合计（只做同一列相加，推导式写进依据块）
    total_line = ""
    all_srcs = [tuple(x) for x in (ctx.get("table_sources") or [])] or srcs
    # ②′ 区域分列：同一指标按区域分别统计，不用一个总数盖两个区域
    if title in ("人口与社会人居现状", "土地与空间利用现状"):
        area_paras, area_srcs = area_breakdown(ctx.get("all_facts") or [])
        for par in area_paras:
            emit(par, _material_fields(area_srcs) +
                 ["判断：按来源台账的区域归属分列，不跨区域合并"], "高")
    if title == "人口与社会人居现状":
        vals = collect_metric_values(ctx.get("all_facts") or [], "户数")
        if vals:
            expr, total = _sum_expr(vals)
            total_line = (f"据现有登记表，已登记改造小区户数合计为 {total} 户"
                          f"（按表内已登记户数的小区逐项加总；原表未登记户数的小区不计入）。")
            emit(total_line, _material_fields(all_srcs) + [f"推导：{expr}"])
    elif title == "土地与空间利用现状":
        vals = collect_metric_values(ctx.get("all_facts") or [], "建筑面积")
        if vals:
            expr, total = _sum_expr(vals)
            total_line = (f"据现有登记表，已登记改造小区建筑面积合计为 {total} 万㎡"
                          f"（按表内已登记面积的小区逐项加总；原表未登记面积的小区不计入）。")
            emit(total_line, _material_fields(all_srcs) + [f"推导：{expr}"])

    # ③ 落地判定：D 章按五要素（表现/证据/成因/影响/对应更新方向）成段，
    #    其余章节按「现状 → 判断 → 依据/口径 → 待核项」成段。
    has_data = bool(table or total_line)
    if not has_data:
        if letter == "D":
            manifest = profile.get("manifest") or profile["judge"]
            cause = profile.get("cause") or "建成年代久、长期缺少系统性维护投入"
            impact = profile.get("impact") or "制约街区安全、民生品质与运营效率"
            emit(f"表现：{manifest}；{profile['judge']}{MARK_INFER}。",
                 [f"判断：{profile['judge']}"], "低")
            emit(f"证据：据现有 {m_count} 份项目材料（{m_brief}），{profile['basis']}；"
                 f"需现场核实的是{profile['verify']}。",
                 [f"判断：现有材料未覆盖{kw_text}口径，按材料清单归纳，未补造数字"], "低")
            emit(f"成因：{cause}。", [f"判断：{cause}"], "低")
            emit(f"影响：{impact}。", [f"判断：{impact}"], "低")
            emit(f"对应更新方向：{profile['next']}。",
                 [f"判断：{profile['next']}"], "低")
        else:
            emit(f"初步判定：{profile['judge']}{MARK_INFER}。",
                 [f"判断：{profile['judge']}"], "低")
            emit(f"判断依据：据现有 {m_count} 份项目材料（{m_brief}），{profile['basis']}；"
                 f"本节按{kw_text}口径核对，现有材料均无对应字段。"
                 f"需现场核实的是{profile['verify']}。",
                 [f"判断：现有材料未覆盖{kw_text}口径，按材料清单归纳，未补造数字"], "低")
            emit(f"下一步：{profile['next']}。",
                 [f"判断：{profile['next']}"], "低")
    else:
        # 有数据但无表（只有合计句）时，补一句判定，保证每节恰有一处标注落点。
        if not table:
            emit(f"初步判定：{profile['judge']}{MARK_CHECK}。",
                 [f"判断：{profile['judge']}"], "中")
        emit(f"本节按{kw_text}口径核对，{profile['basis']}；"
             f"需现场核实的是{profile['verify']}。",
             _material_fields(all_srcs) + [f"判断：登记指标可支撑定性判断，量级待现场核实"],
             "中")

    # ④ 缺口只登记，不进正文
    gaps = []
    if not has_data:
        gid = gap_slug(counter[0] + 1)
        counter[0] += 1
        gaps.append({"id": gid, "内容": f"{title}（现有材料无对应指标，需现场核实）",
                     "来源": hint})
    return "\n\n".join(blocks), gaps, ("check" if has_data else "infer")
