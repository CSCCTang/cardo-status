#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""build_project_narrative.py -- 项目叙述配置的校验与生成。

定位（与 SKILL 的分工）：
  * **正本**：项目侧稳定位置 `<项目根>/项目输入/项目叙述.json`，跨轮次不丢；
  * **快照**：每轮由本脚本生成 `<工作区>/记忆与规则/项目叙述.json`（只读、可重建），
    供 06-分章成文的 baseline_writer / scatter_writer 读取；
  * 工具侧只保留 `Skills/06-分章成文/规则/项目叙述.example.json`（字段清单），
    不含任何具体项目事实。

校验规则：
  * schema 必须是 cardo/project-narrative/v1；
  * 必填字段逐项检查，缺字段直接报错并列出字段路径（不静默兜底）；
  * 给了选填依据时交叉核对：台账社区识别必须在台账转储里命中；AHP 主导准则 id
    必须在权重表里存在。核不上即报错。

用法：
  python build_project_narrative.py --src 项目输入/项目叙述.json \
      --out 记忆与规则/项目叙述.json [--dump 过程文件/03-模板与源报告转储/源报告] \
      [--weights 过程文件/05-事实抽取与事实库/包内依据/包内依据.json]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

SCHEMA = "cardo/project-narrative/v1"

# 06 成文实际读取的字段（缺一项都不可成文），与 规则/项目叙述.example.json 保持一致
REQUIRED = [
    "项目.名称", "项目.城市", "项目.街道", "项目.社区", "项目.社区数量表述",
    "项目.台账社区识别", "项目.台账名称", "项目.台账名称简", "项目.台账简称",
    "项目.统计表简称", "项目.测绘范围四至", "项目.四至道路", "项目.东侧道路",
    "项目.南侧道路", "项目.贯穿街巷", "项目.外围道路", "项目.内部街巷",
    "项目.沿街商业街巷", "项目.建设性质", "项目.建设单位", "项目.实施主体",
    "项目.行业指导部门", "项目.建设内容", "项目.资金来源", "项目.资金来源简述",
    "项目.实施周期", "项目.调研期限", "项目.信息卡建筑面积", "项目.商业建筑",
    "项目.屋面翻新说明", "项目.AHP主导准则", "项目.AHP主导准则称谓",
    "项目.AHP主导准则称谓研判",
    "叙述.项目定位", "叙述.建筑类型演替", "叙述.产权构成", "叙述.产权构成简表",
    "叙述.产权形态并列", "叙述.产权形态三类", "叙述.产权形态数量",
    "叙述.产权形态数量表述", "叙述.院落式住宅数量", "叙述.文脉关键指标",
    "叙述.文脉记忆",
]


def get_path(data, path):
    cur = data
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            raise KeyError(path)
        cur = cur[part]
    return cur


# ── 材料侧推导（无正本时用） ─────────────────────────────────────────────
STREET_RE = re.compile(r"([\u4e00-\u9fa5]{2,6}?(?:街|路|巷|大道|胡同|里))")
CN_NUM = {1: "一", 2: "两", 3: "三", 4: "四", 5: "五", 6: "六"}


def _num(s, default=0.0):
    try:
        return float(re.sub(r"[^0-9.\-]", "", str(s)))
    except Exception:
        return default


def _year(s):
    m = re.search(r"(19|20)\d{2}", str(s))
    return int(m.group()) if m else 0


def project_tokens(name: str) -> set[str]:
    """从项目名取 3~4 字片段，用于判断台账社区/地址是否落在本项目范围。"""
    core = re.sub(r"(老旧街区|老街区|片区|更新|提升|改造|整治|项目|工程|建设|综合)", "", name or "")
    toks = set()
    for size in (4, 3):
        for i in range(len(core) - size + 1):
            t = core[i:i + size]
            if re.search(r"[\u4e00-\u9fa5]", t):
                toks.add(t)
    return toks


def read_ledger_rows(dump_dir):
    """从台账转储块里读小区行（与 baseline_writer 同一列序）。"""
    rows = []
    if not dump_dir or not os.path.isdir(dump_dir):
        return rows
    cur = ""
    last_comm = ""
    for fn in sorted(os.listdir(dump_dir)):
        if not fn.endswith(".blocks.txt"):
            continue
        with open(os.path.join(dump_dir, fn), encoding="utf-8-sig") as f:
            for ln in f:
                m = re.search(r"【来源：(.+?)】", ln)
                if m and m.group(1) != cur:
                    cur = m.group(1)
                    last_comm = ""
                if "表（Sheet1）数据行" not in ln or "小区名称" in ln:
                    continue
                cells = [c.strip() for c in ln.split("数据行：", 1)[-1].split("|")]
                if len(cells) < 13 or not cells[3] or "小区名称" in cells[3]:
                    continue
                # 台账「所在社区」是合并单元格：首行给值，其余行留空。
                # 留空行沿用上一行的社区，否则同一小区的其余门牌会被当成「无社区」
                # 而漏掉（实测示例片区B台账只有首行写「示例社区B」）。
                if cells[2]:
                    last_comm = cells[2]
                rows.append({
                    "file": cur, "street": cells[1], "community": cells[2] or last_comm,
                    "name": cells[3], "address": cells[4],
                    "households": int(_num(cells[5])),
                    "buildings": int(_num(cells[6])),
                    "areaWan": _num(cells[7]), "year": _year(cells[8]),
                    "floors": int(_num(cells[9])), "structure": cells[10],
                    "nature": cells[11], "height": _num(cells[12]),
                    "facade": cells[13] if len(cells) > 13 else "",
                })
    return rows


def run_areas() -> dict:
    """本轮确认的区域口径（开工.py --area 写进 项目识别.json），只对本次运行有效。"""
    path = os.path.join("过程文件", "00-总控", "项目识别.json")
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f) or {}
    except (OSError, ValueError):
        return {}
    return {str(a): [str(x) for x in (vals or []) if str(x).strip()]
            for a, vals in (data.get("areas") or {}).items()}


def run_area_keys() -> list[str]:
    """本轮确认的区域关键词（开工.py --area 写进 项目识别.json）。

    只对本次运行有效，不落 记忆与规则；没有区域口径时返回空。
    """
    keys = []
    for vals in run_areas().values():
        keys.extend(vals)
    return list(dict.fromkeys(keys))


def derive_narrative(project: str, dump_dir: str, baseline_json: str):
    """从项目识别 + 台账 + 包内依据推导一份可用叙述；推不出的登记缺口。"""
    gaps: list[dict] = []

    def gap(path, note):
        gaps.append({"field": path, "note": note})

    rows = read_ledger_rows(dump_dir)
    toks = project_tokens(project)
    # 区域（如示例片区B）已由范围闸门确认纳入本项目：它的台账社区也属于本项目，
    # 关键词一并参与归属判断，不再当邻片排除。
    for _keys in run_areas().values():
        for _k in _keys:
            toks |= project_tokens(_k)
    scope_toks = list(dict.fromkeys(toks))
    by_comm: dict[str, list[dict]] = {}
    for r in rows:
        by_comm.setdefault(r["community"], []).append(r)

    in_scope = {}
    for comm, rs in by_comm.items():
        # 归属看社区名与门牌，不拿来源文件名兜底：同一份台账里既有在册小区也有
        # 模板示例行（如「例：示例街道 | 示例巷」），按文件名归属会把示例行
        # 一起算进本项目，社区列表里冒出「示例巷社区」。
        if not comm:
            continue
        text = comm + " " + " ".join(r["address"] for r in rs)
        if any(t in text for t in scope_toks):
            in_scope[comm] = rs
    # 同文件归属：本项目台账文件（如示例社区C表）内的其余数据行一并纳入。
    # 示例社区C表首行「例：示例巷 示例路253号」社区名不含项目关键词，但它是
    # 该台账的在册行；只按社区/门牌匹配会漏掉 1 小区/36 户/1 栋/0.252 万平方米。
    scope_files = {r["file"] for rs in in_scope.values() for r in rs}
    for comm, rs in by_comm.items():
        if not comm or comm in in_scope:
            continue
        if rs and rs[0]["file"] in scope_files:
            # 其余行的「社区」列可能是院落名（示例社区C台账首行示例巷），不是社区。
            # 不能因此新增一个「示例巷社区」——挂到同一台账文件的已有社区下，
            # 只把行纳入统计，社区清单与「两社区台账」口径不变。
            owner = next((c for c, r2 in in_scope.items()
                          if r2 and r2[0]["file"] == rs[0]["file"]), "")
            if owner:
                in_scope[owner] = in_scope[owner] + rs
            else:
                in_scope[comm] = rs
    scope_rows = [r for rs in in_scope.values() for r in rs]
    if not scope_rows:
        gap("项目.社区", "台账里没有地址/文件名落在项目名范围内的社区")
    # 四至/街巷只按本项目主体区域的台账推；别的区域（如示例片区B）门牌另成一片，
    # 不能把两个区域的道路混成一条边界。主体区域 = 关键词落在项目名里的那个区域，
    # 其余区域的专属关键词才排除（不能把主体区域自己的社区也排掉）。
    other_keys = []
    for keys in run_areas().values():
        if any(k in project for k in keys):
            continue
        other_keys.extend(keys)
    geo_rows = [r for r in scope_rows
                if not any(k in (r["community"] + r["file"] + r["address"])
                           for k in other_keys)] or scope_rows

    # 街道：台账「街道」列众数
    streets_col = [r["street"] for r in geo_rows if r["street"]]
    street = ""
    if streets_col:
        street = max(set(streets_col), key=streets_col.count)
        if not street.endswith("街道"):
            street += "街道"
    else:
        gap("项目.街道", "台账未提供街道列取值")

    # 社区与片区：项目主体片区（项目名所指）与已确认纳入的区域分列，
    # 逐片区的四至分别从本片区台账门牌推，不把两个片区的道路混成一条边界。
    communities = [c if c.endswith("社区") else c + "社区" for c in in_scope]
    ledger_keys = [re.sub(r"社区$", "", c) for c in communities]

    def _area_short(text: str) -> str:
        """把项目/区域全称收成片区名：示例片区更新提升项目 → 示例片区A。"""
        t = re.sub(r"(区域|项目|工程)$", "", str(text or "").strip())
        t = re.sub(r"(老旧街区|老旧小区|老旧厂区)(更新提升|更新|改造|整治)?$", "", t)
        return re.sub(r"(更新提升|更新|改造|整治)$", "", t).strip() or str(text or "").strip()

    project_area = _area_short(project)
    area_defs = []  # [(片区名, 区域关键词, 区域词片)]
    for area_name, keys in run_areas().items():
        toks_a: list[str] = []
        for k in keys:
            toks_a.extend(sorted(project_tokens(k)))
        area_defs.append((_area_short(area_name), list(keys),
                          list(dict.fromkeys(toks_a))))
    comm_area: dict[str, str] = {}
    for c in communities:
        hit = ""
        for nm, keys, toks_a in area_defs:
            if any(k in c for k in keys) or any(t in c for t in toks_a):
                hit = nm
                break
        comm_area[c] = hit or project_area

    def _roads(rs) -> list[str]:
        cnt: dict[str, int] = {}
        for r in rs:
            m = STREET_RE.search(r["address"])
            if m:
                cnt[m.group(1)] = cnt.get(m.group(1), 0) + 1
        return sorted(cnt, key=lambda s: (-cnt[s], list(cnt).index(s)))

    comm_rows = {c if c.endswith("社区") else c + "社区": rs
                 for c, rs in in_scope.items()}
    main_rows = [r for c, rs in comm_rows.items()
                 if comm_area.get(c) == project_area for r in rs]
    geo_rows = main_rows or scope_rows

    # 街道：台账「街道」列众数
    streets_col = [r["street"] for r in geo_rows if r["street"]]
    street = ""
    if streets_col:
        street = max(set(streets_col), key=streets_col.count)
        if not street.endswith("街道"):
            street += "街道"
    else:
        gap("项目.街道", "台账未提供街道列取值")

    roads = _roads(geo_rows)
    if not roads:
        gap("项目.四至道路", "台账地址里没有可识别的道路名")
    outer = [s for s in roads if s.endswith("路") or s[-2:] in ("东街", "南街", "西街", "北街")]
    inner = [s for s in roads if s not in outer]
    if not inner:
        inner = outer[1:]
        outer = outer[:1]
        gap("项目.内部街巷", "台账地址可识别的街巷不足以区分外围/内部，按道路层级粗分")

    # 片区清单（区域在前，项目主体片区在后；各自四至只取本片区台账门牌）
    area_list = []
    for nm, _keys, _toks in area_defs + [(project_area, [], [])]:
        comms = [c for c in communities if comm_area.get(c) == nm]
        if not comms or any(a["名称"] == nm for a in area_list):
            continue
        rs = [r for c in comms for r in comm_rows.get(c, [])]
        area_roads = _roads(rs)
        area_list.append({"名称": nm, "社区": comms,
                          "四至": "、".join(area_roads),
                          "四至道路": area_roads,
                          "四至口径": "围合" if nm == project_area else "台账门牌涉及"})

    # 商业建筑
    biz = next((r for r in scope_rows if "商业" in (r["nature"] or "")), None)
    biz_info = {}
    if biz:
        biz_info = {"数量": 1, "名称": biz["name"], "地址": biz["address"],
                    "建成年份": biz["year"], "结构": biz["structure"] or "—",
                    "外立面": biz["facade"] or "—"}
    else:
        gap("项目.商业建筑", "台账「建筑性质」列没有商业建筑")

    # 产权形态（按小区名称特征归类，台账没有产权列）
    apt = [r for r in scope_rows if "公寓" in r["name"]]
    yard = [r for r in scope_rows if ("院" in r["name"] or "苑" in r["name"]) and "公寓" not in r["name"]]
    dorm = [r for r in scope_rows if ("宿舍" in r["name"] or "研究所" in r["name"])]
    forms = [x for x, lst in (("单位宿舍", dorm), ("院落", yard), ("公寓", apt)) if lst]
    if not forms:
        gap("叙述.产权形态并列", "台账小区名称里没有可识别的产权形态特征")

    years = sorted({r["year"] for r in scope_rows if r["year"]})
    n1980 = sum(1 for r in scope_rows if 1980 <= int(r["year"] or 0) < 1990)
    res = [r for r in scope_rows if "商业" not in (r["nature"] or "")]
    named = {id(r) for lst in (dorm, yard, apt) for r in lst}
    unclass = [r for r in res if id(r) not in named]

    # AHP 主导准则（按权重取前二）
    ahp_ids, ahp_names = [], []
    if baseline_json and os.path.isfile(baseline_json):
        with open(baseline_json, encoding="utf-8-sig") as f:
            base = json.load(f)
        crit = ((base.get("ahp") or {}).get("weights") or {}).get("criteria") or []
        crit = sorted(crit, key=lambda c: -float(c.get("weight") or 0))[:2]
        ahp_ids = [str(c.get("id")) for c in crit]
        ahp_names = [str(c.get("name")) for c in crit]
    if not ahp_ids:
        gap("项目.AHP主导准则", "包内依据里没有 AHP 权重表")

    now = __import__("datetime").datetime.now()
    qty = CN_NUM.get(len(communities), str(len(communities)))
    data = {
        "schema": SCHEMA,
        "note": "本轮由 06 从材料侧推导（项目识别 + 台账 + 包内依据），推不出的字段见 叙述缺口.json。",
        "项目": {
            "名称": project,
            "城市": "",
            "街道": street,
            "社区": communities,
            "社区数量表述": qty,
            "台账社区识别": ledger_keys,
            "台账名称": f"{qty}社区改造小区基本信息及改造内容统计表",
            "台账名称简": f"{qty}社区改造小区统计表",
            "台账简称": f"{qty}社区台账",
            "统计表简称": f"{qty}社区统计表",
            "片区": area_list,
            "测绘范围四至": "、".join(roads),
            "四至道路": roads,
            "东侧道路": "",
            "南侧道路": "",
            "贯穿街巷": inner[:2] or roads[:2],
            "外围道路": outer,
            "内部街巷": inner,
            "沿街商业街巷": ([STREET_RE.search(biz["address"]).group(1)]
                              if biz and STREET_RE.search(biz["address"]) else inner[:1]),
            "建设性质": "既有街区综合整治（非拆除重建）",
            "建设单位": (street + "办事处") if street else "",
            "实施主体": "街道" if street else "",
            "行业指导部门": "",
            "建设内容": "市政管线、消防通道、立面屋面、公共空间与社区服务设施五类",
            "资金来源": "",
            "资金来源简述": "",
            "实施周期": "",
            "调研期限": f"{now.year} 年 {now.month} 月",
            "信息卡建筑面积": 0,
            "商业建筑": biz_info,
            "屋面翻新说明": "",
            "AHP主导准则": ahp_ids,
            "AHP主导准则称谓": "与".join(ahp_names),
            "AHP主导准则称谓研判": "与".join(reversed(ahp_names)),
        },
        "叙述": {
            "项目定位": ("项目整体定位为老旧街区综合更新：以市政基础设施升级和人居环境提质为主线，"
                         "兼顾历史文脉延续与社区服务补短。"),
            "建筑类型演替": (
                f"从建筑类型看，片区住宅以"
                f"{'、'.join(x for x, lst in (('公寓', apt), ('院落式住宅', yard), ('单位宿舍', dorm)) if lst) or '居住建筑'}"
                f"为主，建成年代集中在 {years[0] if years else '—'}—{years[-1] if years else '—'} 年，"
                "空间上呈「院落 + 行列 + 点式」并存的肌理。"),
            "产权构成": (
                "；".join(filter(None, [
                    f"单位宿舍 {len(dorm)} 处" if dorm else "",
                    f"院落式住宅 {len(yard)} 处" if yard else "",
                    f"公寓类住宅 {len(apt)} 处" if apt else "",
                    (f"其余 {len(unclass)} 处以门牌号命名、台账未标注产权形态"
                     if unclass else ""),
                    f"商业 1 处（{biz_info.get('名称')}）" if biz_info else "",
                ])) + "。" if scope_rows else ""),
            "产权构成简表": "、".join(filter(None, [
                f"单位宿舍 {len(dorm)}" if dorm else "",
                f"院落式 {len(yard)}" if yard else "",
                f"公寓 {len(apt)}" if apt else ""])) or "—",
            "产权形态并列": "、".join(forms[:-1]) + "与" + forms[-1] if len(forms) > 1 else (forms[0] if forms else ""),
            "产权形态三类": "、".join(forms),
            "产权形态数量": len(forms),
            "产权形态数量表述": (CN_NUM.get(len(forms), str(len(forms))) + "类") if forms else "",
            "院落式住宅数量": len(yard),
            "文脉关键指标": (f"{n1980} 个 1980 年代小区，院落式住宅 {len(yard)} 处"
                             if years else f"院落式住宅 {len(yard)} 处"),
            "文脉记忆": (f"{years[0]}—{years[-1]} 年代单位宿舍与院落生活记忆"
                         if years else "单位宿舍与院落生活记忆"),
        },
    }
    for path in ("项目.城市", "项目.行业指导部门", "项目.资金来源", "项目.实施周期",
                 "项目.信息卡建筑面积", "项目.东侧道路", "项目.南侧道路"):
        if not get_path(data, path):
            gap(path, "材料侧推不出（需项目信息卡/申报材料），正文跳过或按模板口径处理")
    return data, gaps


def ledger_communities(dump_dir):
    """从台账转储里取「所在社区」列的全部取值（用于交叉核对配置）。"""
    if not dump_dir or not os.path.isdir(dump_dir):
        return None
    vals = set()
    for fn in sorted(os.listdir(dump_dir)):
        if not fn.endswith(".blocks.txt"):
            continue
        with open(os.path.join(dump_dir, fn), encoding="utf-8-sig") as f:
            for ln in f:
                if "数据行：" not in ln:
                    continue
                cells = [c.strip() for c in ln.split("数据行：", 1)[-1].split("|")]
                if len(cells) > 2 and cells[2] and "所在社区" not in cells[2]:
                    vals.add(cells[2])
    return vals


def weights_criteria(baseline_json):
    """从包内依据/权重表里取准则 id 与名称。"""
    if not baseline_json or not os.path.isfile(baseline_json):
        return None
    with open(baseline_json, encoding="utf-8-sig") as f:
        data = json.load(f)
    node = data.get("ahp") if isinstance(data.get("ahp"), dict) else data
    crit = node.get("criteria") or ((node.get("weights") or {}).get("criteria") or [])
    return {str(c.get("id")): str(c.get("name") or "") for c in crit}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default="", help="项目侧正本：项目输入/项目叙述.json")
    ap.add_argument("--out", required=True, help="本轮快照：工作区/记忆与规则/项目叙述.json")
    ap.add_argument("--derive", action="store_true",
                    help="无正本时从材料侧推导（项目识别 + 台账转储 + 包内依据），推不出的登记缺口")
    ap.add_argument("--baseline", default="", help="包内依据.json（推导 AHP 主导准则）")
    ap.add_argument("--project", default="", help="项目名（覆盖正本 项目.名称）")
    ap.add_argument("--dump", default="", help="台账转储目录（用于核对社区识别）")
    ap.add_argument("--weights", default="", help="包内依据.json（用于核对 AHP 准则 id）")
    args = ap.parse_args(argv)

    if args.derive:
        data, gaps = derive_narrative(args.project, args.dump, args.baseline)
        out = os.path.abspath(args.out)
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with open(out, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        gap_json = os.path.join(os.path.dirname(out), "叙述缺口.json")
        with open(gap_json, "w", encoding="utf-8") as f:
            json.dump({"schema": "cardo/narrative-gap/v1",
                       "project": args.project,
                       "narrativeGaps": gaps}, f, ensure_ascii=False, indent=1)
        print(f"项目叙述（材料侧推导）已生成：{out}")
        print(f"  项目：{args.project}；社区：{get_path(data, '项目.社区')}")
        if gaps:
            print(f"  推不出的字段 {len(gaps)} 项，已登记：{gap_json}")
            for g in gaps:
                print(f"    - {g['field']}：{g['note']}")
        return 0

    if not args.src:
        raise SystemExit("缺少 --src（项目叙述正本）；无正本时请用 --derive 从材料侧推导")

    if not os.path.isfile(args.src):
        raise SystemExit(f"项目叙述正本不存在：{args.src}\n"
                         f"请在项目根目录建立 项目输入/项目叙述.json"
                         f"（字段清单见 Skills/06-分章成文/规则/项目叙述.example.json）")
    with open(args.src, encoding="utf-8-sig") as f:
        data = json.load(f)

    if data.get("schema") != SCHEMA:
        raise SystemExit(f"项目叙述配置 schema 不符：{args.src} 为 {data.get('schema')!r}，"
                         f"应为 {SCHEMA!r}")

    missing = []
    for path in REQUIRED:
        try:
            v = get_path(data, path)
        except KeyError:
            missing.append(path)
            continue
        if v is None or (isinstance(v, str) and not v.strip()) or (isinstance(v, list) and not v):
            missing.append(path)
    if missing:
        raise SystemExit("项目叙述配置缺字段（缺失即错，不静默兜底）：\n  "
                         + "\n  ".join(missing))

    if args.project:
        data.setdefault("项目", {})["名称"] = args.project

    keys = [str(x) for x in get_path(data, "项目.台账社区识别")]
    vals = ledger_communities(args.dump)
    if vals is not None:
        unmatched = [k for k in keys if not any(k in v for v in vals)]
        if unmatched:
            raise SystemExit(
                "项目叙述配置与台账不符：项目.台账社区识别 "
                f"{unmatched} 在台账「所在社区」列中未命中。\n"
                f"台账现有社区值：{sorted(vals)}\n"
                "请核对 项目叙述.json 的 项目.台账社区识别，或确认放错台账。")

    crit = weights_criteria(args.weights)
    if crit is not None:
        bad = [i for i in [str(x) for x in get_path(data, "项目.AHP主导准则")] if i not in crit]
        if bad:
            raise SystemExit(f"项目叙述配置与 AHP 权重表不符：项目.AHP主导准则 {bad} "
                             f"不在权重表准则 {list(crit)} 中")

    out = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"项目叙述快照已生成：{out}")
    print(f"  社区：{get_path(data, '项目.社区')}；"
          f"AHP 主导准则：{get_path(data, '项目.AHP主导准则')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
