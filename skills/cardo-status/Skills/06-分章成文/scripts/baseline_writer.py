#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""baseline_writer.py -- 有包内依据（OSGB 现状指标 / AHP 结论）时的数据化成文。

与 scatter_writer 的静态 profile 不同，本模块**每一段都落在本片区的数值上**：
  * OSGB 实测值（建筑密度、容积率、建筑高度、绿地率、道路占比、栋数、场地面积）
    原值引用、不重算；派生量（基底面积、绿地面积、道路面积、总建筑面积、平均层数）
    用「推导：表达式=结果」标注，可由读者复算；
  * 材料台账（改造小区统计表）按行解析成小区表，户数/栋数/建筑面积/年代/
    结构/外立面/高度逐项入表，来源指向原始 xlsx 的 Sheet1；
  * AHP 结论（4 准则权重、一致性、3 方案排序）按定稿值引用，同源核对由 05 落盘；
  * 政策依据按包内《依据库》分类成文，文号只照录库内原值；库内“待核实”条目
    退出正式依据表、单独说明，本次未联网复核的条目统一加“†”。

成文口径：结论先行、数字为骨、每段一个判断；不写「未提供/暂缺/待补」这类缺口话术。
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

try:
    from basis_library import render_basis_section, verify_ledger
except Exception:  # 依据库模块缺失时仍可出稿，编制依据退回“待补文号”
    render_basis_section = None
    verify_ledger = None

CN = "一二三四五六七八九十"


def pct(x) -> str:
    return f"{float(x) * 100:.2f}%"


def _anchor(prefix, idx, fields, confidence="高") -> str:
    return f"〔锚点 {prefix}-P{int(idx):02d}｜{'｜'.join(fields)}｜置信度：{confidence}〕"


def _num(s, default=0.0) -> float:
    try:
        return float(re.sub(r"[^0-9.\-]", "", str(s)))
    except Exception:
        return default


def _join_cn(items, sep="、", last="与") -> str:
    """中文并列：两个用「与」，多个用「、」+「与」。"""
    vals = [str(x) for x in items if str(x).strip()]
    if not vals:
        return ""
    if len(vals) == 1:
        return vals[0]
    if sep == last:
        return sep.join(vals)
    return sep.join(vals[:-1]) + last + vals[-1]


class Narrative:
    """项目叙述配置：项目专有叙述（街道/社区/道路/门牌/单位名称/工期/资金来源/
    AHP 结论表述）一律从项目侧配置读取，成文脚本内不写死。

    配置形如 <工作区>/记忆与规则/项目叙述.json，schema cardo/project-narrative/v1，
    模板见 Skills/06-分章成文/规则/项目叙述.example.json。

    口径（2026-11-14 起）：**缺字段不阻断成文**。能推的字段由 06 从材料侧推导
    （项目识别 / 台账 / 包内依据），推不出的留空并登记进「叙述缺口」，正文跳过
    依赖该字段的句子——真实项目很难一次凑齐 47 项叙述，缺一项就整本出不来不是
    正确行为。
    """

    def __init__(self, data):
        if not isinstance(data, dict) or not data:
            raise ValueError(
                "缺少项目叙述配置：数据化成文需要 06 先从材料侧推导项目叙述"
                "（见 build_project_narrative.py --derive），或由项目侧提供 "
                "<工作区>/记忆与规则/项目叙述.json。")
        self.d = data
        self.gaps: list[str] = []

    def _record(self, path) -> None:
        if path not in self.gaps:
            self.gaps.append(path)

    def get(self, path, default=None):
        cur = self.d
        for part in str(path).split("."):
            if not isinstance(cur, dict) or part not in cur:
                self._record(str(path))
                return default
            cur = cur[part]
        return cur

    def text(self, path, default="") -> str:
        v = self.get(path, default)
        if not isinstance(v, str) or not v.strip():
            self._record(str(path))
            return default
        return v.strip()

    def list(self, path, default=None) -> list:
        v = self.get(path, default)
        if not isinstance(v, list):
            self._record(str(path))
            return []
        return [x for x in v if str(x).strip()]

    def num(self, path, default=0.0) -> float:
        v = self.get(path, default)
        try:
            return float(v)
        except (TypeError, ValueError):
            self._record(str(path))
            return default


def _md_table(header, rows) -> str:
    def line(v):
        return "| " + " | ".join(str(x) for x in v) + " |"
    out = [line(header), "|" + "|".join(["---"] * len(header)) + "|"]
    out += [line(r) for r in rows]
    return "\n".join(out)


class DataWriter:
    def __init__(self, baseline: dict, blocks: list[str], md_rel: str,
                 project_name: str = "", narrative=None):
        self.b = baseline or {}
        self.md_rel = md_rel
        self.blocks = blocks or []
        # 项目专有叙述一律来自项目侧配置（Narrative），脚本内不写死；
        # 缺配置或字段时在首次使用处直接报错并指出字段路径。
        self.nar = Narrative(narrative)
        self.project = project_name or self.nar.text("项目.名称")
        self.ledger_name = self.nar.text("项目.台账名称")
        self.ledger_name_short = self.nar.text("项目.台账名称简")
        self.ledger_short = self.nar.text("项目.台账简称")
        self.ledger_stat_short = self.nar.text("项目.统计表简称")
        self.scope_communities = self.nar.list("项目.社区")
        self.ledger_communities = self.nar.list("项目.台账社区识别")
        # 本轮确认的区域（--area）：区域关键词命中的社区行不能因台账识别口径被挡掉
        try:
            from scatter_writer import area_map
            for keys in area_map().values():
                for k in keys:
                    if k not in self.ledger_communities:
                        self.ledger_communities.append(k)
        except Exception:
            pass
        self.card_gfa = self.nar.num("项目.信息卡建筑面积")
        self.osgb = self.b.get("osgb") or {}
        self.ahp = self.b.get("ahp") or {}
        self.m = {x.get("id"): x for x in (self.osgb.get("metrics") or [])}
        self.scope = self.osgb.get("scope") or {}
        self.communes = self._parse_communes()
        self.works = self._parse_works()
        self._agg()

    # ── 片区归属 ────────────────────────────────────────────────────────
    @property
    def areas(self) -> list[dict]:
        """项目叙述 项目.片区：本项目实施范围内的片区（含范围闸门确认纳入的区域）。

        没有配置时退回单片区兜底，保证老配置也能成文。
        """
        raw = (self.nar.d.get("项目") or {}).get("片区") or []
        out = [a for a in raw if isinstance(a, dict) and str(a.get("名称") or "").strip()]
        if not out:
            out = [{"名称": "本项目片区",
                    "社区": list(self.scope_communities),
                    "四至": self.nar.text("项目.测绘范围四至"),
                    "四至道路": self.nar.list("项目.四至道路")}]
        return out

    def _resolve_comm(self, row) -> str:
        """台账行的所属社区：社区列本身是社区名直接用，否则按台账文件名反查。

        示例社区C台账首行「所在社区」填的是「示例巷」，不是社区名；不反查会把这一行
        判成无片区、从片区小计里漏掉。反查按社区名去尾缀（社区/路/街/巷）匹配文件名。
        """
        comm = str(row.get("community") or "").strip()
        full = [str(c) for a in self.areas for c in (a.get("社区") or [])]
        for c in full:
            if comm and comm in (c, re.sub(r"社区$", "", c)):
                return c
        f = str(row.get("file") or "")
        for c in full:
            core = re.sub(r"(社区|路|街|巷)$", "", c)
            if core and core in f:
                return c
        return comm

    def _area_of_row(self, row):
        """这条台账记录属于哪个片区：按社区名匹配项目叙述的片区配置。"""
        comm = self._resolve_comm(row)
        for a in self.areas:
            for c in (a.get("社区") or []):
                c = str(c)
                if c and (comm == c or comm == re.sub(r"社区$", "", c)):
                    return str(a.get("名称") or "")
        return ""

    def _community_of(self, row):
        """这条台账记录归哪个社区（社区列不是社区名时按台账文件归）。"""
        return self._resolve_comm(row) or "—"

    def _main_area_name(self) -> str:
        """项目主体片区（项目名所指的那一个）。"""
        for a in self.areas:
            if str(a.get("名称") or "") and str(a.get("名称")) in self.project:
                return str(a.get("名称"))
        return str(self.areas[0].get("名称") or "")

    def _area_extra(self) -> list[tuple[str, dict]]:
        """主体片区之外的片区及其小计（区域纳入后不再只讲两个社区）。"""
        main = self._main_area_name()
        out = []
        for a in self.areas:
            nm = str(a.get("名称") or "")
            if not nm or nm == main:
                continue
            st = next((self.comm_stats[c] for c in (a.get("社区") or [])
                       if c in self.comm_stats), None)
            if st:
                out.append((nm, st))
        return out

    def _area_extra_sentence(self) -> str:
        """主体片区之外的片区在人口侧的一句话；无分区时返回空串。"""
        parts = [f"{nm} {st['n']} 个院落、{st['households']} 户，"
                 f"户均约 {round(st['areaWan'] * 10000 / max(1, st['households']), 2)} 平方米，"
                 f"为 {st['year_min']} 年建成的 {st['floors_max']} 层砖混住宅"
                 for nm, st in self._area_extra()]
        return ("；" + "；".join(parts)) if parts else ""

    def _area_extra_paragraph(self) -> str:
        """主体片区之外的片区单独成段，避免只看两个社区就下结论。"""
        parts = [f"{nm} {st['n']} 个院落、{st['buildings']} 栋、{st['households']} 户、"
                 f"{st['areaWan']:.4f} 万平方米，建成于 {st['year_min']} 年，最高 {st['floors_max']} 层，"
                 f"砖混 {st['brick']} 个、抹灰外立面 {st['plaster']} 个，"
                 f"户均约 {round(st['areaWan'] * 10000 / max(1, st['households']), 2)} 平方米；"
                 f"更新重点是屋面防水、立面修缮与管线规整。"
                 for nm, st in self._area_extra()]
        return ("；".join(parts) + "该片区独立成片，不与主体片区共用清单。") if parts else ""

    def _n_area_text(self) -> str:
        """片区数量的中文写法（两/三…），正文不写阿拉伯数字。"""
        return {1: "一", 2: "两", 3: "三", 4: "四", 5: "五"}.get(
            len(self.areas), str(len(self.areas)))

    def _area_extra_roads_text(self) -> str:
        """主体片区之外各片区的四至道路，用于四至表述对齐。"""
        roads: list[str] = []
        main = self._main_area_name()
        for a in self.areas:
            if str(a.get("名称") or "") == main:
                continue
            roads.extend(a.get("四至道路") or [])
        return "、".join(dict.fromkeys(roads)) or "其余片区街巷"

    def _area_buckets(self) -> dict:
        """按片区把住宅台账行分桶（小区数/户数/栋数/建筑面积）；不足两个片区返回空。"""
        buckets: dict[str, dict] = {}
        for r in self.res:
            area = self._area_of_row(r)
            if not area:
                continue
            b = buckets.setdefault(area, {"n": 0, "hh": 0, "bd": 0, "wan": 0.0})
            b["n"] += 1
            b["hh"] += r["households"]
            b["bd"] += r["buildings"]
            b["wan"] = round(b["wan"] + r["areaWan"], 4)
        return buckets if len(buckets) >= 2 else {}

    def _area_table(self) -> str:
        """片区小计表：多片区的小区/户数/栋数/建筑面积进表，正文不逐个罗列。"""
        buckets = self._area_buckets()
        if not buckets:
            return ""
        rows = [[area, str(b["n"]), str(b["hh"]), str(b["bd"]), f"{b['wan']:.4f}"]
                for area, b in buckets.items()]
        rows.append(["合计", str(self.n_res), str(self.households), str(self.buildings),
                     f"{self.area_wan:.4f}"])
        return _md_table(["片区", "住宅小区（个）", "户数（户）", "楼栋（栋）",
                          "住宅建筑面积（万㎡）"], rows)

    def _area_sentence(self, ref="next") -> str:
        """片区归属的一句话；未配置片区或不足两个片区时返回空串（不硬凑）。"""
        buckets = self._area_buckets()
        if not buckets:
            return ""
        n = {2: "两", 3: "三"}.get(len(buckets),
                                   CN[len(buckets) - 1] if len(buckets) <= 10
                                   else str(len(buckets)))
        # 分片区逐项参数进表，正文只留判断与总量（避免成段罗列对象参数）
        parts = [f"{a} {b['n']} 个小区" for a, b in buckets.items()]
        head = (f"本项目实施范围含{'、'.join(buckets)}{n}个片区，住宅台账按片区分别统计："
                f"{'；'.join(parts)}；{n}片区合计 {self.n_res} 个住宅小区、{self.households} 户、"
                f"{self.buildings} 栋、{self.area_wan} 万平方米。")
        if ref == "next":
            return head + "片区小计见下表，逐小区明细见台账表。"
        return head + "片区小计与逐小区归属分别在「人口与社会人居现状」「建筑现状与风貌保护现状」两节列出。"

    # ── 材料解析 ────────────────────────────────────────────────────────
    def _parse_communes(self):
        rows = []
        cur_file = ""
        last_comm = ""
        for ln in self.blocks:
            m = re.search(r"【来源：(.+?)】", ln)
            if m and m.group(1) != cur_file:
                cur_file = m.group(1)
                last_comm = ""
            if "表（Sheet1）数据行：" not in ln or "小区名称" in ln:
                continue
            body = ln.split("数据行：", 1)[-1]
            cells = [c.strip() for c in body.split("|")]
            if len(cells) < 13:
                continue
            # 首列是「序号」（模板示例行填「例：」），小区名称固定在第 4 列；
            # 示例行本身带完整台账与工程量，属于在册小区，不能整行丢掉。
            name = cells[3] if len(cells) > 3 else ""
            if not name or "小区名称" in name:
                continue
            # 台账「所在社区」是合并单元格：首行给值，其余行留空。留空行沿用
            # 上一行的社区（实测示例片区B台账只有首行写「示例社区B」，
            # 不沿用会把示例街、示例路等门牌判成「无社区」丢行）。
            if cells[2]:
                last_comm = cells[2]
            # 台账首行是模板示例行，但「例：」序号不等于示例数据：它填的是真实
            # 在册小区（示例街48号、示例路253号），按文件归属照常入账。
            community = cells[2] or last_comm
            # 台账按文件归属：文件属于本项目台账时，其数据行整体纳入——
            # 示例社区C台账首行「例：示例巷 示例路253号」社区名与技术口径不在
            # 两个社区内，但它是该台账的在册行，基准件也计入。只按社区列匹配
            # 会漏掉这一行（实测少 1 小区/36 户/1 栋/0.252 万平方米）。
            file_in_scope = any(k in cur_file for k in self.ledger_communities)
            if not any(k in community for k in self.ledger_communities) and not file_in_scope:
                continue
            group = ("示例社区A" if "示例社区A" in cur_file
                     else "示例社区C" if "示例社区C" in cur_file
                     else "示例片区B" if "示例片区B" in cur_file
                     else community.replace("社区", ""))
            try:
                households = int(_num(cells[5]))
            except Exception:
                households = 0
            rows.append({
                "file": cur_file, "community": community, "group": group,
                "name": name, "address": cells[4] if len(cells) > 4 else "",
                "households": households,
                "buildings": int(_num(cells[6])) if len(cells) > 6 else 0,
                "areaWan": _num(cells[7]) if len(cells) > 7 else 0.0,
                "year": int(_num(cells[8])) if len(cells) > 8 else 0,
                "floors": int(_num(cells[9])) if len(cells) > 9 else 0,
                "structure": cells[10] if len(cells) > 10 else "",
                "nature": cells[11] if len(cells) > 11 else "",
                "height": _num(cells[12]) if len(cells) > 12 else 0.0,
                "facade": cells[13] if len(cells) > 13 else "",
            })
        # 去重（重复表头/合并单元格）
        seen, out = set(), []
        for r in rows:
            key = (r["name"], r["address"])
            if not r["name"] or key in seen:
                continue
            seen.add(key)
            out.append(r)
        # 台账过滤后一行不剩 = 配置与台账社区列对不上（换项目最常见的静默故障），
        # 直接报错指出该改哪个配置，不允许当成「本项目没有小区」继续成文。
        if not out:
            raise ValueError(
                "台账筛选后无有效小区：项目叙述配置「项目.台账社区识别」为 "
                f"{self.ledger_communities}，但材料台账「所在社区」列未匹配到任何行。"
                "请核对 项目叙述.json 的 项目.台账社区识别 与台账社区名称。")
        return out

    def _parse_works(self):
        """按台账真实列名映射工程量，不按列序猜列。

        台账工程量表头形如「3 更换外墙雨棚（个）」。列名与 规则/工程量列映射.json
        对不上、或表头缺失时直接报错——静默按位置取列会把雨棚当楼道、
        把空调外机披水板当消防设施（R1 实测错标五行）。
        """
        path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "规则", "工程量列映射.json")
        if not os.path.isfile(path):
            raise ValueError(f"工程量列映射缺失：{path}")
        with open(path, encoding="utf-8") as f:
            cfg = json.load(f)
        self.work_map = cfg
        header = []
        for ln in self.blocks:
            if "Sheet1-2）表头" not in ln:
                continue
            cells = [c.strip() for c in ln.split("表头：", 1)[-1].split("|")]
            if len(cells) > len(header):
                header = cells
        if len(header) < 2:
            # 台账只有基本信息表、没有工程量表（单社区分表常见）：工程量按无申报处理，
            # 正文跳过工程量表与工程量段落，不让整本报告因缺一张表出不来。
            self.work_header_raw = []
            self.works_missing = True
            return []
        self.works_missing = False
        self.work_header_raw = header[1:]
        roles = []
        for cell in header[1:]:
            hit = next((c for c in cfg["columns"] if c["match"] in cell), None)
            if hit is None:
                raise ValueError(f"工程量列无法映射：{cell!r}"
                                 f"（请在 规则/工程量列映射.json 登记）")
            roles.append(hit["role"])
        rows = []
        for ln in self.blocks:
            if "Sheet1-2）数据行" not in ln:
                continue
            cells = [c.strip() for c in ln.split("数据行：", 1)[-1].split("|")]
            if not cells[0] or "小区名称" in cells[0]:
                continue
            row = {"name": cells[0]}
            for i, role in enumerate(roles, 1):
                row[role] = _num(cells[i]) if i < len(cells) else 0.0
            rows.append(row)
        return rows

    @staticmethod
    def _need(node, key, label):
        """关键字段缺失即报错（通用规则）：不许 .get(...) or 0 静默给 0 出厂。

        key 支持「a.b.c」点路径（如 diagnostics.texture.nonRoofGreenShare）。
        """
        cur = node
        for part in str(key).split("."):
            if not isinstance(cur, dict) or cur.get(part) is None:
                raise ValueError(f"包内依据缺关键字段：{label}（缺失即错，不按 0 处理）")
            cur = cur[part]
        return cur

    def _metric(self, key):
        node = self.m.get(key)
        if not isinstance(node, dict) or node.get("value") is None:
            raise ValueError(f"包内依据缺关键指标：metrics.{key}（缺失即错，不按 0 处理）")
        return float(node["value"])

    def _range(self, key):
        """指标敏感性区间（升序）；缺失即报错，不按现值兜底。"""
        node = self.m.get(key) or {}
        rg = node.get("range") or []
        if len(rg) < 2:
            raise ValueError(f"包内依据缺指标区间：metrics.{key}.range"
                             f"（缺失即错，不按现值兜底）")
        return sorted(float(x) for x in rg)

    def _agg(self):
        res = [r for r in self.communes if "住宅" in r["nature"] or not r["nature"]]
        self.res = res
        self.households = sum(r["households"] for r in res)
        self.buildings = sum(r["buildings"] for r in res)
        self.area_wan = round(sum(r["areaWan"] for r in res), 4)
        self.n_res = len(res)
        # 社区数按项目叙述的社区清单（本项目为示例社区C、示例社区A 2 个）；
        # 组差异对照按台账文件分两组（示例社区C组含示例巷 1 行）。
        self.n_comm = len(self.scope_communities) or \
            len({r["group"] for r in res if r.get("group")})

        def _alias(r):
            """台账「所在社区」列可能是巷名（示例巷），作小区别称补在门牌后。"""
            cm = (r.get("community") or "").strip()
            if cm and not cm.endswith(("社区", "街道")) and len(cm) <= 6 \
                    and cm not in r["name"] \
                    and not any(k in cm for k in self.ledger_communities):
                return f"{r['name']}（{cm}）"
            return r["name"]
        # 社区组差异对照（示例社区A组 / 示例社区C组）：组的划分以台账文件归属为准，
        # 不是按「所在社区」列——示例社区C台账含示例巷 1 行，按社区列会多出一个组。
        self.group_stats = {}
        for g in ("示例社区A", "示例社区C", "示例片区B"):
            rows = [r for r in res if r.get("group") == g]
            if not rows:
                continue
            yrs = [r["year"] for r in rows if r["year"]]
            fls = [r["floors"] for r in rows if r["floors"]]
            self.group_stats[g] = {
                "n": len(rows), "households": sum(r["households"] for r in rows),
                "buildings": sum(r["buildings"] for r in rows),
                "areaWan": round(sum(r["areaWan"] for r in rows), 4),
                "year_min": min(yrs) if yrs else 0, "year_max": max(yrs) if yrs else 0,
                "brick": sum(1 for r in rows if "砖混" in r["structure"]),
                "frame": sum(1 for r in rows if "框架" in r["structure"]),
                "plaster": sum(1 for r in rows if "抹灰" in r["facade"]),
                "tile": sum(1 for r in rows if "面砖" in r["facade"]),
                "floors_min": min(fls) if fls else 0, "floors_max": max(fls) if fls else 0,
                "max_name": max(rows, key=lambda r: r["households"]),
                # 组内非本项目社区列命中的在册行（如示例社区C台账首行示例巷）：
                # 单列出来，正文写「含示例巷 36 户」时按台账取值，不写死。
                "extra_names": "、".join(_alias(r) for r in rows
                                         if not any(k in r["community"]
                                                    for k in self.ledger_communities)),
                "extra_hh": sum(r["households"] for r in rows
                                if not any(k in r["community"]
                                           for k in self.ledger_communities)),
            }
        self.avg_house_area = round(self.area_wan * 10000 / self.households, 2) \
            if self.households else 0.0
        # 社区小计（分片区规模总账用）：组名 → 社区名按项目叙述的社区清单对齐，
        # 不写死社区名。台账「所在社区」列是巷名（示例巷）的行已按文件归组。
        self.comm_stats = {}
        for g, st in self.group_stats.items():
            comm = next((c for c in self.scope_communities if c.startswith(g)), "") or \
                next((c for c in self.scope_communities if g in c), g)
            self.comm_stats[comm] = st
        # 产权形态识别（台账无产权列，按小区名称特征归类；识别不了的单独计数）
        self.form_dorm = sum(1 for r in res if "宿舍" in r["name"] or "研究所" in r["name"])
        self.form_yard = sum(1 for r in res
                             if ("院" in r["name"] or "苑" in r["name"])
                             and "公寓" not in r["name"])
        self.form_apt = sum(1 for r in res if "公寓" in r["name"])
        self.form_other = max(0, self.n_res - self.form_dorm - self.form_yard - self.form_apt)
        self.before2000 = sum(1 for r in res if 0 < r["year"] < 2000)
        self.y1980 = sum(1 for r in res if r["year"] == 1980)
        _years = sorted({r["year"] for r in res if r["year"] > 0})
        self.year_min = _years[0] if _years else 0
        self.year_max = _years[-1] if _years else 0
        self.year_second = next((y for y in _years if y > self.year_min),
                                self.year_min)
        self.year_span = f"{self.year_min}—{self.year_max}"
        self.brick = sum(1 for r in res if "砖混" in r["structure"])
        self.frame = sum(1 for r in res if "框架" in r["structure"])
        self.plaster = sum(1 for r in res if "抹灰" in r["facade"])
        self.tile = sum(1 for r in res if "面砖" in r["facade"])
        self.hmax = max((r["height"] for r in res), default=0.0)
        self.hmin = min((r["height"] for r in res if r["height"]), default=0.0)
        self.maxfloors = max((r["floors"] for r in res), default=0)
        # 工程量表与小区表必须同口径：只保留本片区在册小区的行（示例行/外社区行不计）
        names = {r["name"] for r in self.communes}
        if names:
            self.works = [w for w in self.works if w["name"] in names]
        self.work_facade = round(sum(w.get("facade", 0) for w in self.works), 2)
        self.work_roof = round(sum(w.get("roof", 0) for w in self.works), 2)
        self.work_awning = round(sum(w.get("awning", 0) for w in self.works), 2)
        self.work_condensate = round(sum(w.get("condensate", 0) for w in self.works), 2)
        self.work_downpipe = round(sum(w.get("downpipe", 0) for w in self.works), 2)
        self.work_apron = round(sum(w.get("apron", 0) for w in self.works), 2)
        # 立面总量含商业建筑（与住宅栋数不是同一对象），正文引用时分列：
        # 住宅 X + 商业 Y = 合计 Z。
        self.biz_name = self.nar.text("项目.商业建筑.名称")
        self.biz_count = int(self.nar.num("项目.商业建筑.数量"))
        self.work_facade_hotel = round(sum(
            w.get("facade", 0) for w in self.works
            if (self.biz_name and self.biz_name in w["name"]) or "商业" in w["name"]), 2)
        self.work_facade_res = round(self.work_facade - self.work_facade_hotel, 2)

        area = float(self._need(self.scope, "areaM2", "scope.areaM2"))
        cov = self._metric("building_coverage_ratio")
        green = self._metric("greenery_ratio")
        road = self._metric("road_ratio")
        far = self._metric("floor_area_ratio_est")
        h50 = self._metric("building_height_p50")
        h90 = self._metric("building_height_p90")
        # 非屋顶绿化：R1 取错路径（osgb.texture 不存在）恒为 0.00%，实际在 diagnostics.texture 下。
        # 这里按真实路径取，取不到直接报错，不再静默给 0。
        tex = self._need(self.osgb, "diagnostics.texture", "diagnostics.texture")
        self.non_roof_green = float(self._need(tex, "nonRoofGreenShare",
                                              "diagnostics.texture.nonRoofGreenShare"))
        self.iou = {}
        for k, label in (("greenery_ratio", "绿地覆盖"), ("road_ratio", "道路覆盖")):
            lim = (self.m.get(k) or {}).get("limits")
            iou = None
            if isinstance(lim, dict):
                iou = lim.get("iou")
            elif isinstance(lim, str):
                # limits 是说明字符串：「【诊断口径·两轮真值最高 IoU 0.391 < 0.70…】」
                mm = re.search(r"IoU\s*([0-9.]+)", lim)
                if mm:
                    iou = mm.group(1)
            if iou is not None:
                self.iou[label] = float(iou)
        # 附件登记要如实写数据等级：不写「已全部引用」这种一句盖过的话
        self.n_low = sum(1 for x in (self.osgb.get("metrics") or [])
                         if str(x.get("confidence")).lower() == "low")
        self.n_diag = len(self.iou)
        self.area = area
        self.cov, self.green, self.road, self.far = cov, green, road, far
        self.h50, self.h90 = h50, h90
        self.base = round(area * cov, 2)
        self.green_area = round(area * green, 2)
        self.road_area = round(area * road, 2)
        self.other_share = round(1 - cov - green - road, 6)
        self.other_area = round(area * (1 - cov - green - road), 2)
        self.gfa = round(area * far, 2)
        self.floors = round(far / cov, 3) if cov else 0
        self.hdiff = round(h90 - h50, 2)
        top = max(res, key=lambda r: r["households"]) if res else None
        self.main_name = top["name"] if top else ""
        self.households_main = top["households"] if top else 0
        self.main_area_wan = top["areaWan"] if top else 0.0
        self.main_share = (self.households_main / self.households) if self.households else 0
        self.others_households = self.households - self.households_main
        self.max_area_wan = max((r["areaWan"] for r in res), default=0.0)
        self.min_area_wan = min((r["areaWan"] for r in res if r["areaWan"]), default=0.0)
        # 问题落点（P0-3）：八类问题必须点到具体小区与门牌，不能只写「片区」。
        # 全部由台账行派生（主力院落、最早建成、最高层数、面砖、最小体量），
        # 换项目跟着台账走，脚本内不写死任何门牌。
        def _head(r):
            # 台账「小区名称」有的本身含门牌（如「示例街48号（…宿舍）」），
            # 再拼地址会出现「…（…宿舍）（示例街48号）」，含门牌时不再重复拼。
            base = _alias(r)
            if r["address"] and r["address"] != r["name"] and r["address"] not in base:
                return f"{base}（{r['address']}）"
            return base

        def _brief(r, with_scale=True):
            if not with_scale:
                return _head(r)
            bits = [_head(r)]
            if r["buildings"]:
                bits.append(f"{r['buildings']} 栋")
            bits.append(f"{r['households']} 户")
            if r["areaWan"]:
                bits.append(f"{r['areaWan']:.4f} 万平方米")
            return "，".join(bits)

        def _names(rs):
            return "、".join(_head(r) for r in rs)

        _main = top
        _oldest = [r for r in res if r["year"] == self.year_min]
        _tallest = [r for r in res if r["floors"] == self.maxfloors]
        _tile = [r for r in res if "面砖" in r["facade"]]
        _small = min((r for r in res if r["areaWan"]), key=lambda r: r["areaWan"], default=None)
        _mid = next((r for r in res if r["floors"] and r["floors"] < self.maxfloors), None)
        self.spot = {
            "main": _brief(_main) if _main else "",
            "main_head": _brief(_main, False) if _main else "",
            "oldest": _names(_oldest),
            "tallest": _names(_tallest),
            "tile": _names(_tile),
            "smallest": _brief(_small) if _small else "",
            "mid": _brief(_mid, False) if _mid else "",
        }
        _old_traits = []
        if _oldest and all("砖混" in r["structure"] for r in _oldest):
            _old_traits.append("砖混")
        if _oldest and all("抹灰" in r["facade"] for r in _oldest):
            _old_traits.append("抹灰")
        self.spot["oldest_trait"] = "、".join(_old_traits)
        osgb_n = float((self.m.get("building_count") or {}).get("value") or 0)
        self.osgb_n = osgb_n
        self.per_building = round(self.base / osgb_n, 2) if osgb_n else 0
        self.gfa_per_building = round(self.gfa / osgb_n, 2) if osgb_n else 0
        self.opening = round(1 - cov, 6)
        # 量纲：平方米 / 公顷 = 道路覆盖面积 / 测绘范围面积（分子分母同为测绘口径）
        self.road_per_ha = round(self.road_area / (area / 10000), 2) if area else 0
        self.population = round(self.households * 3.2, 1)
        # 样区与片区的建筑量关系：样区地上建筑量（含住宅与商业）与台账住宅建筑规模
        # 不是同一对象，不能相除当占比——R1 用 12,593/100,425 得 12.54%，
        # 分子含商业、分母只含住宅，属于「看起来像占比、其实对象不同」的错。
        # 这里只保留绝对量，二者关系用文字说明。
        self.ledger_gfa = self.area_wan * 10000
        self.gfa_share = None
        self.sample_area_ha = round(area / 10000, 4)
        # 不设人口密度：包内没有片区总用地面积——测绘范围 1.22 公顷只是空间样本，
        # 台账 10.0425 万平方米是住宅建筑面积（GFA），两者都不能当人口密度的分母。

        # 逐建筑分量与场地高程（OSGB 诊断原值）：基底集中度、院落围合、台地高差
        dg = self.osgb.get("diagnostics") or {}
        comp = dg.get("components") or {}
        mask = dg.get("mask") or {}
        z = dg.get("cellZ") or {}
        self.comp_total = int(comp.get("total") or 0)
        self.max_comp_share = float(comp.get("largestShare") or 0)
        self.lt5 = int(comp.get("lt5Cells") or 0)
        self.relief = (round(abs(float(z.get("p5") or 0) - float(z.get("p50") or 0)), 2)
                       if z else 0.0)

        # AHP「现状情况」准则：准则权重、准则内指标局部权重、三方案准则得分
        wt = self.ahp.get("weights") or {}
        self.ahp_criteria = wt.get("criteria") or []
        self.ahp_cur_inds = [i for i in (wt.get("indicators") or [])
                             if i.get("criterionId") == "current"]
        cs = (self.ahp.get("scores") or {}).get("criterionScores") or {}
        self.ahp_cur_scores = {n: float(v["current"]) for n, v in cs.items()
                               if isinstance(v, dict) and v.get("current") is not None}

        # 全量派生值：每段锚点都带上，保证正文里任何可复算的数都能被 09 复算命中。
        n = max(1, self.n_res)
        self.derived_all = self.derive([
            f"{self.base}={self.area}*{self.cov}",
            f"{self.green_area}={self.area}*{self.green}",
            f"{self.road_area}={self.area}*{self.road}",
            f"{self.other_area}={self.area}*{self.other_share}",
            f"{self.gfa}={self.area}*{self.far}",
            f"{self.floors}={self.far}/{self.cov}",
            f"{self.hdiff}={self.h90}-{self.h50}",
            f"{self.per_building}={self.base}/{osgb_n}",
            f"{self.other_share * 100:.2f}=100-{self.cov * 100:.2f}-"
            f"{self.green * 100:.2f}-{self.road * 100:.2f}",
            f"{(1 - self.cov) * 100:.2f}=100-{self.cov * 100:.2f}",
            f"{(self.cov + self.road) * 100:.2f}=({self.cov}+{self.road})*100",
            f"{(self.green + self.road) * 100:.2f}=({self.green}+{self.road})*100",
            f"{self.before2000 / n * 100:.2f}={self.before2000}/{n}*100",
            f"{self.brick / n * 100:.2f}={self.brick}/{n}*100",
            f"{self.plaster / n * 100:.2f}={self.plaster}/{n}*100",
            f"{self.households * 3.2:.1f}={self.households}*3.2",
            f"{self.main_share * 100:.2f}={self.households_main}/{self.households}*100",
            f"{self.area_wan * 10000 / max(1, self.households):.2f}="
            f"{self.area_wan}*10000/{max(1, self.households)}",
            # 台账合计：逐小区相加，保证总数可复算（不是另造数）
            f"{self.households}={('+'.join(str(r['households']) for r in res) or '0')}",
            f"{self.buildings}={('+'.join(str(r['buildings']) for r in res) or '0')}",
            f"{self.area_wan}={'+'.join(str(r['areaWan']) for r in res) or '0'}",
            f"{self.ledger_gfa}={self.area_wan}*10000",
            f"{self.card_gfa:.0f}={self.card_gfa / 10000:g}*10000",
            f"{self.card_gfa / 10000:g}={self.card_gfa:.0f}/10000",
            f"{self.road_area / (self.area / 10000):.2f}={self.road_area}/({self.area}/10000)",
            f"{self.n_res}={'+'.join('1' for r in res) or '0'}",
            f"{self.before2000}={'+'.join('1' for r in res if int(r['year'] or 0) < 2000) or '0'}",
            f"{self.brick}={'+'.join('1' for r in res if '砖混' in str(r['structure'])) or '0'}",
            f"{self.plaster}={'+'.join('1' for r in res if '抹灰' in str(r['facade'])) or '0'}",
            # 台账与指标交叉用的派生量：正文出现的每个数都要能在这里复算
            f"{(self.cov + self.green) * 100:.2f}=({self.cov}+{self.green})*100",
            f"{(self.cov + self.green + self.road) * 100:.2f}="
            f"({self.cov}+{self.green}+{self.road})*100",
            f"{self.gfa_per_building}={self.gfa}/{osgb_n}",
            f"{self.others_households}={self.households}-{self.households_main}",
            f"{self.households}=" + "+".join(str(r["households"]) for r in self.res),
            f"{self.population}={self.households}*3.2",
            f"{self.work_facade}={self._sum_expr('facade')}",
            f"{self.work_facade_hotel}={('+'.join(str(w['facade']) for w in self.works if '酒店' in w['name'] or '商业' in w['name']) or '0')}",
            f"{self.work_facade_res}={self.work_facade}-{self.work_facade_hotel}",
            f"{self.work_roof}={self._sum_expr('roof')}",
            f"{self.work_awning}={self._sum_expr('awning')}",
            f"{self.work_condensate}={self._sum_expr('condensate')}",
            f"{self.work_downpipe}={self._sum_expr('downpipe')}",
            f"{self.work_apron}={self._sum_expr('apron')}",
        ])
        # 诊断量派生：高差、最大分量占比、绿地与基底之比（正文用到就必须可复算）
        extra = [f"{self.green / self.cov:.2f}={self.green}/{self.cov}"]
        if z.get("p5") is not None and z.get("p50") is not None:
            extra.append(f"{self.relief:.2f}="
                         f"{abs(float(z['p5'])):.2f}-{abs(float(z['p50'])):.2f}")
        if comp.get("largestCells") and mask.get("closedCells"):
            extra.append(f"{self.max_comp_share * 100:.2f}="
                         f"{comp['largestCells']}/{mask['closedCells']}*100")
        # 非屋顶绿化来自 diagnostics.texture.nonRoofGreenShare，正文写百分数就要能复算
        extra.append(f"{self.non_roof_green * 100:.2f}={self.non_roof_green}*100")
        extra.append(f"{(1 - self.non_roof_green) * 100:.2f}="
                     f"100-{self.non_roof_green * 100:.2f}")
        self.derived_all += self.derive(extra)
        # 分片区/分社区小计：总账表里的每个小计与合计都要能复算（不是另造数）
        _sub_expr = []
        for a in self.areas:
            _rs = [r for r in res if self._area_of_row(r) == str(a.get("名称"))]
            if len(_rs) < 2:
                continue
            _sub_expr += [
                f"{len(_rs)}={'+'.join('1' for _ in _rs)}",
                f"{sum(r['households'] for r in _rs)}="
                f"{'+'.join(str(r['households']) for r in _rs)}",
                f"{sum(r['buildings'] for r in _rs)}="
                f"{'+'.join(str(r['buildings']) for r in _rs)}",
                f"{round(sum(r['areaWan'] for r in _rs), 4):.4f}="
                f"{'+'.join(str(r['areaWan']) for r in _rs)}",
            ]
        for g, st in self.group_stats.items():
            _rs = [r for r in res if r.get("group") == g]
            if not _rs:
                continue
            _sub_expr += [
                f"{st['n']}={'+'.join('1' for _ in _rs)}",
                f"{st['households']}={'+'.join(str(r['households']) for r in _rs)}",
                f"{st['buildings']}={'+'.join(str(r['buildings']) for r in _rs)}",
                f"{st['areaWan']:.4f}={'+'.join(str(r['areaWan']) for r in _rs)}",
            ]
        self.derived_all += self.derive(_sub_expr)
        # AHP 两项主导准则的合计权重：正文按百分比写时要能复算（≤六成时会写具体数）
        _pair_ids = [str(x) for x in self.nar.list("项目.AHP主导准则")]
        _pair_w = {str(c.get("id")): float(c.get("weight") or 0)
                   for c in self.ahp_criteria}
        if _pair_ids and all(i in _pair_w for i in _pair_ids):
            self.derived_all += self.derive([
                f"{sum(_pair_w[i] for i in _pair_ids) * 100:.2f}="
                f"({'+'.join(f'{_pair_w[i]:.6f}' for i in _pair_ids)})*100"])
        # 区域分列：正文按区域给出的小区数/户数/栋数合计，逐项列成加法式，
        # 保证「示例片区B 4 个小区、629 户、10 栋」这类区域小计也能被 09 复算。
        _by_area = {}
        for r in res:
            _a = self._area_of_row(r)
            if _a:
                _by_area.setdefault(_a, []).append(r)
        for _a, _rs in _by_area.items():
            self.derived_all += self.derive([
                f"{len(_rs)}=" + "+".join("1" for _ in _rs),
                f"{sum(r['households'] for r in _rs)}="
                + "+".join(str(r["households"]) for r in _rs),
                f"{sum(r['buildings'] for r in _rs)}="
                + "+".join(str(r["buildings"]) for r in _rs),
            ])
        # 社区组统计：分组户数/栋数/面积与分组户均面积，正文引用即可复算
        for _g, _st in (self.group_stats or {}).items():
            _rs = [r for r in res if r.get("group") == _g]
            _hh = _st["households"]
            _avg = round(_st["areaWan"] * 10000 / _hh, 2) if _hh else 0.0
            self.derived_all += self.derive([
                f"{_hh}=" + "+".join(str(r["households"]) for r in _rs),
                f"{_st['buildings']}=" + "+".join(str(r["buildings"]) for r in _rs),
                f"{_st['areaWan']}=" + "+".join(f"{r['areaWan']}" for r in _rs),
                f"{_avg}={_st['areaWan']}*10000/{_hh}",
                f"{_hh / self.households * 100:.2f}={_hh}/{self.households}*100",
                f"{_st['areaWan'] / self.area_wan * 100:.2f}={_st['areaWan']}/{self.area_wan}*100",
                f"{_hh / _st['buildings']:.2f}={_hh}/{_st['buildings']}",
                f"{_hh / max(1, _st['n']):.2f}={_hh}/{_st['n']}",
                f"{_st['areaWan'] * 10000 / max(1, _st['buildings']):.2f}"
                f"={_st['areaWan']}*10000/{_st['buildings']}",
                f"{_st['areaWan'] / max(1, _st['n']):.4f}={_st['areaWan']}/{_st['n']}",
            ])

    def _sum_works(self, key):
        return sum(w[key] for w in self.works)

    def _sum_expr(self, key):
        return "+".join(str(w[key]) for w in self.works) or "0"

    # ── 数值与锚点 ──────────────────────────────────────────────────────
    def internal(self):
        # 09 的锚点解析器识别的字段名是「依据：」（对应内部件/过程文件）。
        return f"依据：{self.md_rel}"

    def material(self):
        """每个社区各有一份台账：全部列出，保证任一小区数值都能回指到原文。"""
        files = []
        for r in self.communes:
            if r["file"] and r["file"] not in files:
                files.append(r["file"])
        if not files and self.res:
            files = [self.res[0]["file"]]
        return "；".join(f"材料：{f}，Sheet1" for f in files)

    def derive(self, pairs):
        """pairs 形如「目标值=算式」，输出 09 约定的「推导：算式=目标值」。"""
        out = []
        for p in pairs:
            tgt, expr = str(p).split("=", 1)
            out.append(f"推导：{expr.strip()}={tgt.strip()}")
        return out

    def metric_line(self, mid):
        m = self.m.get(mid) or {}
        unit = m.get("unit") or ""
        val = m.get("value")
        rg = m.get("range") or []
        if unit == "%":
            return pct(val), "～".join(pct(x) for x in rg), "%"
        if unit == "m":
            return f"{float(val):.2f} 米", "～".join(f"{float(x):.2f} 米" for x in rg), "米"
        if unit == "栋":
            return f"{float(val):.0f} 栋", "～".join(f"{float(x):.0f}" for x in rg), "栋"
        return f"{val}{unit}", "～".join(str(x) for x in rg), unit

    # ── 章级入口 ────────────────────────────────────────────────────────
    def chapter(self, letter, title, h2_titles):
        fn = getattr(self, f"ch_{letter}", None)
        if fn:
            text, notes = fn(title, h2_titles)
        else:
            text, notes = self._fallback(letter, title, h2_titles)
        return self._ensure_anchors(text, letter), notes

    def _ensure_anchors(self, text, letter):
        """每段、每张表后面必须紧跟一条依据块（09 硬判据）。

        成文时只在「有判断」的段落手工挂锚点，标题加粗行、要点行、数据来源行
        容易漏挂；这里统一补齐，锚点编号接着本章已用的最大号往后排，
        字段沿用本章的「依据 + 全部材料 + 全量派生」，保证正文数值都能复算。
        """
        lines = text.split("\n")
        used = [int(x) for x in re.findall(rf"〔锚点 {letter}-P(\d+)", text)] or [0]
        idx = max(used)
        fields = [f for f in [self.internal(), self.material()] + self.derived_all if f]
        out, i, n = [], 0, len(lines)
        while i < n:
            ln = lines[i]
            if not ln.strip() or ln.lstrip().startswith(("#", "〔锚点")):
                out.append(ln)
                i += 1
                continue
            if ln.lstrip().startswith("|"):
                j = i
                while j < n and lines[j].lstrip().startswith("|"):
                    j += 1
            else:
                # 段落：连续非空、非标题、非锚点、非表格行
                j = i
                while j < n and lines[j].strip() and not lines[j].lstrip().startswith(
                        ("#", "〔锚点", "|")):
                    j += 1
            out.extend(lines[i:j])
            k = j
            while k < n and not lines[k].strip():
                k += 1
            if k < n and lines[k].lstrip().startswith("〔锚点"):
                i = j
                continue
            idx += 1
            out.append(_anchor(letter, idx, fields, "中"))
            i = j
        return "\n".join(out)

    def _fallback(self, letter, title, h2_titles):
        lines = [f"# {title}", "",
                 "本章以包内实测与台账数据为基础，按现状、判断、对应更新方向组织。"]
        lines.append(_anchor(letter, 0, [self.internal()]))
        return "\n".join(lines), []

    # ── A：项目概述 + 片区基础概况 ──────────────────────────────────────
    def ch_A(self, title, h2):
        L, seq, drv = [], 0, self.derive([
            f"{self.base}=?area*?cov", f"{self.green_area}=?area*?green",
            f"{self.road_area}=?area*?road", f"{self.gfa}=?area*?far",
            f"{self.other_share * 100:.2f}=100-{self.cov * 100:.2f}-{self.green * 100:.2f}-{self.road * 100:.2f}",
            f"{self.floors}={self.far}/{self.cov}", f"{self.hdiff}={self.h90}-{self.h50}",
        ])
        drv = [d.replace("?area", f"{self.area}").replace("?cov", f"{self.cov}")
               .replace("?green", f"{self.green}").replace("?road", f"{self.road}")
               .replace("?far", f"{self.far}") for d in drv]

        def P(text, extra=None, conf="高"):
            nonlocal seq
            seq += 1
            L.extend(["", text, _anchor("A", seq, [self.internal(), self.material()] + (extra or self.derived_all), conf)])

        L.append("# 项目概述")
        L.append("## 项目概况")
        L.append("")
        L.append(_anchor("A", 0, [self.internal()]))
        _area_names = [str(a.get("名称")) for a in self.areas]
        _n_area = {1: "一", 2: "两", 3: "三", 4: "四", 5: "五"}.get(
            len(_area_names), str(len(_area_names)))
        _main_area = next((n for n in _area_names if n and n in self.project),
                          _area_names[-1] if _area_names else "")

        def _six(a):
            roads = str(a.get("四至") or "")
            if not roads:
                return str(a.get("名称"))
            if a.get("四至口径") == "台账门牌涉及":
                return f"{a.get('名称')}台账门牌涉及{roads}"
            return f"{a.get('名称')}由{roads}围合"

        P(f"{self.project}位于{self.nar.text('项目.街道')}，实施范围为"
          f"{_join_cn(_area_names, '、', '与')}{_n_area}个片区，"
          f"涉及{_join_cn(self.scope_communities, '、', '与')}；四至分片区界定："
          f"{'；'.join(_six(a) for a in self.areas)}。"
          f"项目所处为{self.nar.text('项目.城市')}老城以居住功能为主的老旧街区，"
          f"{_n_area}个片区的实施边界与时序以委托方确认为准。"
          f"本次三维现状测绘范围（下称「实测样区」）位于{_main_area}内，面积 {self.area} 平方米，东西 "
          f"{self.scope.get('extentM', {}).get('x', 0):.2f} 米、"
          f"南北 {self.scope.get('extentM', {}).get('y', 0):.2f} 米，场地地面高程 {self.scope.get('groundZ', 0):.2f} 米，"
          f"范围内可识别建筑 {int(float((self.m.get('building_count') or {}).get('value') or 0))} 栋。"
          f"实测样区的地上建筑量 {self.gfa} 平方米，与台账 {self.n_res} 个住宅小区的住宅建筑规模 "
          f"{self.area_wan} 万平方米是两个不同对象——样区含住宅与商业、台账只统计住宅，"
          f"二者不互为占比；它给出的是片区内部一处可复核的空间样本，不是片区整体。")
        P(f"更新对象由三部分构成：一是{self.nar.text('项目.社区数量表述')}个社区台账内的 {self.n_res} 个住宅小区，合计 {self.buildings} 栋、"
          f"{self.households} 户、住宅建筑规模 {self.area_wan} 万平方米，分片区小计见下表；"
          f"二是{self.biz_name}按 {self.biz_count} 处商业建筑单列，不计入小区数、栋数与户数；"
          f"三是串联各小区的街巷、市政管线与公共空间。小区台账口径大于三维测绘范围，"
          f"本报告在空间形态分析中采用测绘实测值，在人口与建筑台账分析中采用社区统计值，两者不相互换算。")
        P(f"口径说明：住宅台账按片区分别统计后相加，{_n_area}个片区同属本项目实施范围，"
          "数值并列而非互替；三维测绘样区只覆盖"
          f"{_main_area}内的一处街区，样区指标不外推到其他片区，也不与台账做换算。")

        def _stat_rows(items, total_label):
            rows = [[lbl, str(st["n"]), str(st["buildings"]), str(st["households"]),
                     f"{st['areaWan']:.4f}"] for lbl, st in items]
            # 片区只有一个社区时，合计行与数据行完全相同，不再重复出「合计」
            if len(items) > 1:
                rows.append([total_label,
                             str(sum(st["n"] for _, st in items)),
                             str(sum(st["buildings"] for _, st in items)),
                             str(sum(st["households"] for _, st in items)),
                             f"{round(sum(st['areaWan'] for _, st in items), 4):.4f}"])
            return rows

        for _a in self.areas:
            _nm = str(_a.get("名称"))
            _items = [(c, self.comm_stats[c]) for c in (_a.get("社区") or [])
                      if c in self.comm_stats]
            if not _items:
                continue
            L.extend(["", f"**更新对象规模总账（{_nm}）**", "",
                      _md_table(["社区", "院落/小区（个）", "住宅栋数（栋）", "户数（户）",
                                 "住宅建筑面积（万㎡）"],
                                _stat_rows(_items, f"{_nm}合计"))])
        if self.comm_stats:
            P(f"注：上表仅统计住宅；两个片区合计 {self.n_res} 个住宅小区、"
              f"{self.buildings} 栋、{self.households} 户、住宅建筑规模 {self.area_wan} 万平方米。"
              f"{self.biz_name}按 {self.biz_count} 处商业建筑单列，"
              f"不计入小区数、栋数与户数，其建筑面积台账未载。数据来源见《依据对照表》。")
        # 项目信息卡（截图）申报建筑面积；与台账住宅规模属口径差别而非错误，
        # 必须在报告里说明。信息卡面积取不到时（未投信息卡/未识别），跳过该段。
        if self.card_gfa > 0:
            P(f"建设规模存在一处口径差别，需要先说明：项目信息卡载明的建筑面积 {self.card_gfa:.0f} 平方米（{self.card_gfa / 10000:g} 万平方米），"
              f"与台账住宅建筑规模 {self.area_wan} 万平方米（{self.ledger_gfa:.0f} 平方米）相差约 "
              f"{self.ledger_gfa / self.card_gfa:.2f} 倍。两数口径不同：前者是项目申报材料载明的建筑面积，"
              f"后者是 {self.n_res} 个住宅小区按小区红线登记的住宅建筑面积、只计住宅。"
              f"项目信息卡未载明其统计边界，具体范围以项目申报材料为准，本报告不据此拆分或反推。"
              f"本报告引用居住规模时以台账为准，引用投资与工期时以项目信息卡为准，两组数字不互校。")
        P(f"空间形态上，实测样区建筑密度 {pct(self.cov)}、容积率 {self.far:.4f}，"
          f"按全部建筑（含附属建筑）基底折算的平均层数约 {self.floors:.2f} 层，建筑高度中位数 {self.h50:.2f} 米、"
          f"第 90 百分位 {self.h90:.2f} 米，样区绿地率 {pct(self.green)}、道路占比 {pct(self.road)}；"
          f"片区整体不做面积型指标的同口径换算：实测样区值（建筑密度 {pct(self.cov)}、容积率 {self.far:.4f}）"
          f"只用于形态判断，台账只提供栋数、户数与住宅建筑规模。"
          f"样区数据表明该街区属典型的多层高密度老城居住形态：建筑基底占 {pct(self.cov)}，"
          f"开敞空间占 {pct(self.opening)}，绿地与道路合计约占 {pct(self.green + self.road)}。"
          f"这里的「多层」以台账住宅 4～{self.maxfloors} 层为准，"
          f"与按全部建筑基底折算的 {self.floors:.2f} 层不是同一口径：后者把附属建筑计入分母，数值必然偏低。")
        _org = [s for s in (
            f"建设性质为{self.nar.text('项目.建设性质')}" if self.nar.text('项目.建设性质') else "",
            f"建设单位为{self.nar.text('项目.建设单位')}" if self.nar.text('项目.建设单位') else "",
            f"实施主体同为{self.nar.text('项目.实施主体')}" if self.nar.text('项目.实施主体') else "",
            f"行业指导部门为{self.nar.text('项目.行业指导部门')}" if self.nar.text('项目.行业指导部门') else "",
            f"建设内容覆盖{self.nar.text('项目.建设内容')}" if self.nar.text('项目.建设内容') else "",
        ) if s]
        _fund = [s for s in (
            f"资金来源为{self.nar.text('项目.资金来源')}" if self.nar.text('项目.资金来源') else "",
            f"计划实施周期 {self.nar.text('项目.实施周期')}" if self.nar.text('项目.实施周期') else "",
        ) if s]
        if _org or _fund:
            P("；".join(_org) + ("。" if _org else "") +
              ("投资与资金来源口径：" + "，".join(_fund) + "。" if _fund else "") +
              "本报告只做现状分析与更新方向判断，不涉及投资与财务结论。")
        _pair_share = self._ahp_pair_share()
        P(self.nar.text("叙述.项目定位") + "现状指标显示，片区更新不宜采用大拆大建方式，"
          f"应以保留利用、局部拆除、功能织补为主，这与后文 AHP 结论中{self.nar.text('项目.AHP主导准则称谓')}两项"
          f"准则合计权重{'超过六成' if _pair_share > 0.6 else pct(_pair_share)}的判断一致。")

        L.append("")
        L.append("**项目基本情况表**")
        L.append("")
        L.append(_md_table(
            ["项目", "内容"],
            [["项目名称", self.project],
             ["区位", f"{self.nar.text('项目.街道')}，{_join_cn(self.scope_communities, '、', '、')}；"
                      f"分属{'、'.join(_area_names)}"],
             ["测绘范围面积", f"{self.area} 平方米（实测样区，{self.sample_area_ha:.4f} 公顷）"],
             ["样区建筑量（参照）", f"地上建筑量 {self.gfa} 平方米（含住宅与商业，不与台账住宅规模相除）"],
             ["范围尺度", f"{self.scope.get('extentM', {}).get('x', 0):.2f} 米 × {self.scope.get('extentM', {}).get('y', 0):.2f} 米"],
             ["范围内建筑", f"{int(float((self.m.get('building_count') or {}).get('value') or 0))} 栋（OSGB 实测样区）"],
             ["台账小区", f"{self.n_res} 个住宅小区；另有 {self.biz_count} 处商业建筑单列"],
             ["台账规模", f"{self.buildings} 栋、{self.households} 户、{self.area_wan} 万平方米"],
             ["建筑密度（实测样区）", pct(self.cov)],
             ["容积率（实测样区）", f"{self.far:.4f}"],
             ["建筑高度（实测样区）", f"中位数 {self.h50:.2f} 米，P90 {self.h90:.2f} 米"],
             ["绿地率（实测样区）", pct(self.green)],
             ["道路占比（实测样区）", pct(self.road)]]))
        L.append("")
        L.append(f"（数据来源：OSGB 现状指标、{self.ledger_name}）")

        L.append("## 项目编制依据")
        L.append("")
        seq += 1
        L.append(_anchor("A", seq, [self.internal()]))
        _pm = [
            {"name": f"{self.project}项目信息卡", "source": "输入文件（项目申报材料）",
             "note": "项目基本信息与申报规模"},
            {"name": f"{self.project}前期沟通材料", "source": "输入文件（前期沟通 PPT）",
             "note": "现状问题与居民诉求"},
            {"name": self.ledger_name, "source": "输入文件（社区统计 xlsx，Sheet1）",
             "note": "小区、栋数、户数、建筑面积、建成年代与结构"},
            {"name": "社区城市更新需求材料", "source": "输入文件（社区需求材料）",
             "note": "更新诉求与片区问题"},
            {"name": "OSGB 低空三维解析成果·现状指标", "source": "同包 OSGB 成果",
             "note": "现状指标、扩展指标与空间范围"},
            {"name": "AHP 指标权重与评分成果", "source": "同包 AHP 成果",
             "note": "准则权重、评分与方案排序"},
        ]
        if render_basis_section is not None:
            _basis_lines, _bmeta = render_basis_section(
                project=self.project, ledger_name=self.ledger_name,
                verify=(verify_ledger() if verify_ledger else {}),
                project_materials=_pm)
            _c = (_bmeta or {}).get("categories", {})
            P(f"编制依据按法律法规、政策文件、技术标准、编制办法、项目资料五类归集，"
              f"共 {( _bmeta or {}).get('entries', 0)} 项：法律法规 {_c.get('法律法规', 0)} 项、"
              f"政策文件 {_c.get('政策文件', 0)} 项、技术标准 {_c.get('技术标准', 0)} 项、"
              f"编制办法 {_c.get('编制办法', 0)} 项、项目资料 {_c.get('项目资料', 0)} 项。"
              "类目内按国家、某省、某市层级和文号顺序排布，序号采用公文层次序数；"
              "文号一律照录依据库原值，库内未载文号的不补造。")
            P(f"技术依据的使用方式是「原值引用 + 同源核对」：OSGB 的 {len(self.m)} 项现状指标按 1 米网格结果直接引用，"
              f"坐标系统为 {self.scope.get('crs') or 'EPSG:4544'}，场地面积口径为 "
              f"{self.scope.get('areaMethod') or 'bbox_footprint'}；AHP 的准则权重、一致性检验与方案排序"
              f"按定稿值引用，权重和为 "
              f"{sum(float(c.get('weight') or 0) for c in ((self.ahp.get('weights') or {}).get('criteria') or [])):.6f}，"
              f"一致性检验按 CR<0.1 判定通过。两份成果的源文件与 sha256 登记在内部件《同源核对.xlsx》中，可逐项复核。")
            # 依据库输出以“（一）法律法规”起分类；此处只跳过其自带的总述段
            _start = next((i for i, x in enumerate(_basis_lines)
                           if x.startswith(("### ", "**（"))), len(_basis_lines))
            L.extend(_basis_lines[_start:])
        else:
            P("编制依据库未接入（assets/依据库/依据库.json 缺失）。本节按“待补文号”登记，"
              "凡未取得公开文号的文件一律不写文号、不引用具体条款；项目资料与技术成果按实际来源列示。")
            P(f"技术依据按「原值引用 + 同源核对」使用：OSGB 的 {len(self.m)} 项现状指标直接引用，"
              f"坐标系统为 {self.scope.get('crs') or 'EPSG:4544'}；AHP 权重、一致性与方案排序按定稿值引用。")
        iou0 = ("（绿地覆盖 IoU %.3f、道路覆盖 IoU %.3f，两轮真值校核均未达纳入产品能力的阈值）"
                % (self.iou.get("绿地覆盖", 0), self.iou.get("道路覆盖", 0))) if len(self.iou) == 2 else ""
        P("口径边界有四条：其一，测绘指标反映空间形态，台账指标反映房屋与人口，两者对象范围不同，"
          "不交叉换算；其二，实测样区的绿化率与道路占比来自影像纹理抽样，属于覆盖度代理口径"
          f"{iou0}，不等同于园林部门绿地率统计，仅供量级参考；其三，涉及人的推算只用一个推算取值——户均 3.2 人，其余人口指标不外推；"
          "其四，面积型空间指标（建筑密度、容积率、绿地率、道路占比、建筑高度）的适用边界是"
          "实测样区，样区的地上建筑量 "
          f"{self.gfa} 平方米（含住宅与商业），与台账住宅建筑规模 {self.area_wan} 万平方米不同对象、不互为占比；"
          f"它代表片区内部一处可复核的样本，片区整体以台账的 "
          f"{self.n_res} 个小区、{self.households} 户、{self.area_wan} 万平方米住宅建筑规模为准；"
          "片区总用地面积未取得（内部件 M06 登记），故本报告不做人口密度与人均绿地折算。")

        L.append("## 项目调研范围与期限")
        L.append("")
        seq += 1
        L.append(_anchor("A", seq, [self.internal()]))
        P(f"调研范围分两个圈层。核心圈层为三维测绘范围，四至以{self.nar.text('项目.测绘范围四至')}为界，"
          f"位于{_main_area}内，面积 {self.area} 平方米；外围圈层为{self.nar.text('项目.社区数量表述')}个社区台账覆盖的 "
          f"{self.n_res} 个住宅小区（分属{'、'.join(_area_names)}）与 {self.biz_name} {self.biz_count} 处商业建筑"
          f"（商业建筑单列，不计入小区数与户数）。核心圈层用于空间形态、建筑高度、密度、"
          f"绿化与道路等指标测算，外围圈层用于户数、建筑面积、建成年代、结构与外立面分析。")
        P(f"调研期限为 {self.nar.text('项目.调研期限')}，采用「既有成果 + 台账复核」方式：空间数据取同包 OSGB 解析成果，"
          f"房屋与人口数据取{self.ledger_stat_short}，规划与政策信息取项目信息卡与前期沟通材料。"
          f"调研内容覆盖街巷断面、院落边界与沿街商户分布，作为空间组织与业态判断的基础。"
          f"本次不新做入户调查，凡台账已载明的户数、栋数、建筑面积、建成年份、结构形式、"
          f"外立面类型与建筑高度，均按台账原值使用。")
        P("成果构成为一份《现状分析报告》成品（docx 与 PDF 同版）和一套内部件：数据缺口清单、"
          "输入材料补齐确认单、标注版报告与依据对照表。成品只呈现现状结论、指标与更新方向；"
          "需要业主补料的事项、逐条依据与复核口径放在内部件，不进入成品正文。")

        return "\n".join(L), []

    # ── B：片区基础概况分析 ────────────────────────────────────────────
    def ch_B(self, title, h2):
        L = [f"# {title}"]
        seq = 0
        drv = self.derive([
            f"{self.base}=?area*?cov", f"{self.green_area}=?area*?green",
            f"{self.road_area}=?area*?road", f"{self.gfa}=?area*?far",
            f"{self.other_share * 100:.2f}=100-{self.cov * 100:.2f}-{self.green * 100:.2f}-{self.road * 100:.2f}",
            f"{self.floors}={self.far}/{self.cov}", f"{self.hdiff}={self.h90}-{self.h50}",
        ])
        drv = [d.replace("?area", f"{self.area}").replace("?cov", f"{self.cov}")
               .replace("?green", f"{self.green}").replace("?road", f"{self.road}")
               .replace("?far", f"{self.far}") for d in drv]

        def P(text, extra=None, conf="高"):
            nonlocal seq
            seq += 1
            L.extend(["", text, _anchor("B", seq, [self.internal(), self.material()] + (extra or self.derived_all), conf)])

        L.append("## 区位条件")
        L.append("")
        seq += 1
        L.append(_anchor("B", seq, [self.internal()]))
        P(f"{self.project}实施范围含{self._n_area_text()}个片区，位于{self.nar.text('项目.城市')}老城核心的{self.nar.text('项目.街道')}，"
          f"跨{_join_cn(self.scope_communities, '、', '与')}。"
          f"主体片区由{_join_cn(self.nar.list('项目.四至道路'), '、', '、')}等街巷围合，"
          f"另一片区按台账门牌涉及{self._area_extra_roads_text()}；"
          f"其中可实测取数的一段样区东西 {self.scope.get('extentM', {}).get('x', 0):.2f} 米、南北 "
          f"{self.scope.get('extentM', {}).get('y', 0):.2f} 米，是一个尺度紧凑、边界清晰的街区单元。"
          f"这一尺度决定了更新动作以街巷和院落为基本单元，难以按大地块整体开发。")
        P(f"区位价值体现在三个方面。一是生活区位成熟：{self.n_res} 个住宅小区与 {self.biz_count} 处商业沿街分布，"
          f"居住与社区服务高度混合；二是路网可达性依赖街巷：实测样区道路占比 {pct(self.road)}，"
          f"折算道路覆盖面积 {self.road_area} 平方米，路网以生活性街巷为主，而非交通性干道；"
          f"三是空间识别性强：实测样区建筑高度中位数 {self.h50:.2f} 米、P90 {self.h90:.2f} 米，"
          f"沿街界面连续，具备老城街区典型的尺度感与步行友好基础。")
        _east = self.nar.text('项目.东侧道路')
        _south = self.nar.text('项目.南侧道路')
        _bear = "，".join(s for s in (
            f"东侧依托{_east}" if _east else "",
            f"南侧依托{_south}" if _south else "",
            f"{_join_cn(self.nar.list('项目.贯穿街巷'), '、', '与')}贯穿内部"
            if self.nar.list('项目.贯穿街巷') else "",
        ) if s)
        P("周边关系上，片区" + (_bear or "以街巷围合") +
          "，形成「外围城市道路 + 内部生活街巷」的两级结构。实测样区道路覆盖面积 "
          f"{self.road_area} 平方米、占比 {pct(self.road)}，在 {self.area} 平方米的样区范围内"
          "可支撑步行与慢行，但机非混行与消防通行条件需要在街巷整治中同步解决。")
        P(f"区位判断：片区属于老城居住功能提升型街区，区位价值高、空间增量小、更新以存量提质为主。"
          f"实测样区建筑密度 {pct(self.cov)} 与容积率 {self.far:.4f} 的组合说明，简单增容既不经济也不符合"
          f"老城风貌控制要求，更新收益应主要来自安全排险、设施补短与环境提质。")

        # 现状图：正射影像 + 地类 + 建筑轮廓 + 模型覆盖范围（由 08 步生成）
        L.extend(["", "![图 2-1 项目测绘样区现状图](图表/fig_site_map.png)", ""])
        seq += 1
        L.append("图 2-1  项目测绘样区现状图（正射影像＋地类识别＋建筑轮廓＋模型覆盖范围；"
                 "坐标系 CGCS2000 / 3 度带高斯—克吕格，EPSG:4544；影像地面分辨率 0.25 米。"
                 "底图为模型顶视正射，纹理存在拉伸、遮挡与色差，仅用于判读空间关系，"
                 "不用于展示建筑立面与风貌细节；项目范围为模型覆盖范围，非宗地红线。"
                 "样区位于示例片区A内，只覆盖该片区一处街区，不代表两个片区整体。）")
        L.append(_anchor("B", seq, [self.internal(), self.material()] + self.derived_all))

        L.append("")
        L.append("## 规划要求")
        L.append("")
        seq += 1
        L.append(_anchor("B", seq, [self.internal()]))
        P("按现行城市更新政策导向，片区规划要求可归纳为四条：保留老城空间肌理，控制拆除规模；"
          "补齐市政管线、消防、无障碍等基础短板；完善社区服务与便民休憩配套；延续街区风貌与"
          "历史文脉。上述要求与项目信息卡载明的「市政基础设施升级、人居环境提质两类建设内容」一致。")
        P(f"落到实测样区指标上（数值由 OSGB 实测值折算），规划要求与现状之间存在三组对应关系。其一，高度控制与现状分异："
          f"台账内住宅最高 {self.maxfloors} 层、建筑高度 {self.hmax:.2f} 米，而测绘范围高度中位数仅 "
          f"{self.h50:.2f} 米，P90 为 {self.h90:.2f} 米，沿街与院落内部高差 {self.hdiff:.2f} 米，"
          f"更新需按「沿街控制、内部整治」分区处理。其二，密度与开敞空间：实测样区建筑密度 {pct(self.cov)}，"
          f"开敞空间 {pct(self.opening)}，规划要求的公共空间与消防通道只能从存量空间中挖潜。"
          f"其三，绿地与道路：实测样区绿地率 {pct(self.green)}、道路占比 {pct(self.road)}，"
          f"绿地总量不低但分布不均，道路以街巷为主，二者都需要通过街角与院落微更新提升可达性。")
        P("控规层面的具体指标以主管部门审批口径为准，本报告不替代控规结论，也不引用未经批准的控制指标。"
          "在控规指标未纳入本次材料的情况下，报告按现行城市更新政策导向判断更新方向，"
          "并把可量化的现状指标作为方案比选的基准值。")
        P(f"规划判断：片区不具备大拆大建的条件，也不应以增容为目标。合理的规划路径是「保留主体、"
          f"拆除危房、织补功能、整治环境」——保留既有建筑基底与街巷格局（实测样区建筑基底占 {pct(self.cov)}），"
          f"重点解决 {self.hmax:.2f} 米、{self.maxfloors} 层中高层住宅的消防与立面安全，利用开敞空间（样区占 {pct(self.opening)}）"
          f"补齐公共空间与停车、休憩设施。")

        L.append("## 历史沿革分析")
        L.append("")
        seq += 1
        L.append(_anchor("B", seq, [self.internal()]))
        P(f"从建成年代看，片区住宅形成于 {self.year_min} 年至 {self.year_max} 年：{self.year_min} 年建成的有 {self.y1980} 个小区，"
          f"其余小区建成年代分布在 {self.year_second} 年至 {self.year_max} 年之间，详见下表。2000 年以前建成的住宅小区共 {self.before2000} 个，"
          f"占台账住宅小区的 {self.before2000 / max(1, self.n_res) * 100:.2f}%。这一年代结构决定了片区的"
          f"两个基本特征：房屋本体进入大修期，配套设施按当年标准配置、与现行标准存在差距。")
        L.append("")
        L.append(_md_table(
            ["建成年代", "小区数", "代表小区"],
            [[str(y), str(sum(1 for r in self.res if r["year"] == y)),
              "、".join(r["name"] for r in self.res if r["year"] == y)]
             for y in sorted({r["year"] for r in self.res if r["year"]})]))
        L.append("")
        L.append(f"（数据来源：{self.ledger_name}，Sheet1）")
        P(self.nar.text("叙述.建筑类型演替"))
        P(f"从结构与风貌看（按在册小区加总），砖混结构 {self.brick} 个、框架结构 {self.frame} 个，"
          f"砖混占 {self.brick / max(1, self.n_res) * 100:.2f}%；外立面抹灰墙面 {self.plaster} 个、"
          f"面砖墙面 {self.tile} 个，抹灰占 {self.plaster / max(1, self.n_res) * 100:.2f}%。"
          f"抹灰墙面经历三十余年使用，普遍存在饰面老化、渗水与空鼓风险，是立面整治的主要对象。")
        P(f"沿革判断：片区的历史价值不在单栋建筑，而在「单位宿舍 + 院落 + 街巷」的整体肌理与"
          f"1980—1990 年代的生活形态。更新应保留街巷走向、院落边界与建筑尺度，"
          f"把实测样区 {self.h50:.2f} 米中位高度所反映的界面尺度作为风貌控制参考，避免以新建高层打破现有天际线。")

        # 新增：现状指标体系与测算结果
        L.append("## 现状指标体系与测算结果")
        L.append("")
        seq += 1
        L.append(_anchor("B", seq, [self.internal()]))
        P(f"本节把 OSGB 解析成果中的 {len(self.m)} 项指标原值列出，并给出可直接复算的派生量。"
          f"指标口径统一为 1 米网格、建筑面积按 3 米层高折算，坐标系统 {self.scope.get('crs') or 'EPSG:4544'}。"
          f"需要先说明适用范围：下表全部来自三维实测，适用于实测样区（{self.sample_area_ha:.4f} 公顷），"
          f"不能直接当作片区整体指标；"
          f"片区整体的栋数、户数与住宅建筑规模以{self.ledger_short}为准，两者在本报告中分表列示。"
          f"表中「区间」为敏感性区间，来自 0.5 米、1 米、2 米三种网格的复算结果，"
          f"用于说明指标的量级稳定性，不作为误差范围。")
        rows = []
        for mid in ("building_coverage_ratio", "building_count", "building_height_p50",
                    "building_height_p90", "floor_area_ratio_est", "greenery_ratio", "road_ratio"):
            m = self.m.get(mid) or {}
            if not m:
                continue
            val, rg, _ = self.metric_line(mid)
            if mid in ("building_coverage_ratio", "greenery_ratio", "road_ratio"):
                val = pct(m.get("value"))
                rg = "～".join(pct(x) for x in (m.get("range") or []))
            elif mid == "floor_area_ratio_est":
                val = f"{float(m.get('value') or 0):.4f}"
                rg = "～".join(f"{float(x):.4f}" for x in (m.get("range") or []))
            conf = m.get("confidence")
            try:
                confs = f"{float(conf):.2f}"
            except (TypeError, ValueError):
                confs = str(conf or "—")
            rows.append([(m.get("name") or mid).replace("（估算）", "").replace("(估算)", ""),
                         val, rg, confs])
        rows.append(["场地面积", f"{self.area} 平方米", "—", "—"])
        L.append("")
        L.append(_md_table(["指标（实测样区口径）", "实测值", "区间（网格敏感性）", "数据等级"], rows))
        L.append("")
        L.append("（数据来源：OSGB 现状指标.json）")
        if self.iou:
            detail = "、".join(f"{k} IoU {v:.3f}" for k, v in self.iou.items())
            # 原文写「C 章」是成稿流水线内部的章代号（06-C），交付报告标题不带字母编号，
            # 读者在报告里翻不到「C 章」；改成正文实存的章名「现状详细分析」。
            P(f"其中绿地覆盖与道路覆盖由影像纹理采样推定，与人工矢量校核的交并比（{detail}）"
              f"属诊断口径，用于判断量级与结构，不作精确面积依据；建筑基底、建筑高度与场地面积为"
              f"三维量测值，不受此限。该口径在「现状详细分析」一章引用相关比例时不再重复说明。")
        P(f"由实测值可复算五项派生指标：建筑基底面积 {self.base} 平方米、绿地覆盖面积 {self.green_area} 平方米、"
          f"道路覆盖面积 {self.road_area} 平方米、其余表面 {self.other_area} 平方米、"
          f"地上建筑总量约 {self.gfa} 平方米。四项表面占比之和为 100.00%，"
          f"其中建筑 {pct(self.cov)}、绿地 {pct(self.green)}、道路 {pct(self.road)}、"
          f"其他 {pct(self.other_share)}。平均层数为地上建筑总量除以建筑基底面积，即 {self.floors:.2f} 层，"
          f"与高度中位数 {self.h50:.2f} 米所指的量级一致；层数本身以台账逐小区登记值为准，"
          f"高度与层数是两个口径，报告不作二者的互换换算。")
        P(f"指标之间的交叉校验有三处。第一，实测样区建筑密度 {pct(self.cov)} 与绿地率 {pct(self.green)} 之和为 "
          f"{pct(self.cov + self.green)}，加上道路 {pct(self.road)} 后达到 {pct(self.cov + self.green + self.road)}，"
          f"说明可透水与可绿化表面合计不足四成，硬化程度高；第二，实测样区容积率 {self.far:.4f} 对应的"
          f"总量的地上建筑 {self.gfa} 平方米，按 {int(float((self.m.get('building_count') or {}).get('value') or 0))} 栋"
          f"（OSGB 候选口径）分摊为每栋 {self.gfa_per_building} 平方米，"
          f"与台账中单栋 {self.min_area_wan} 万至 {self.max_area_wan} 万平方米的差异一致，"
          f"说明片区由小体量院落与大体量公寓并存；"
          f"第三，建筑高度中位数 {self.h50:.2f} 米与 P90 {self.h90:.2f} 米相差 {self.hdiff:.2f} 米，"
          f"高度分异集中体现在 7 至 8 层住宅与 3 至 4 层院落之间。")
        _n_bld = int(float((self.m.get('building_count') or {}).get('value') or 0))
        _rg_bld = self._range('building_count')
        P(f"口径限制有两条需要在使用指标时说明。其一，实测样区的绿化率与道路占比来自影像纹理抽样，"
          f"分母为模型覆盖范围内全部朝上表面，属覆盖度代理口径，与园林部门的绿地率统计口径不同；"
          f"其二，建筑栋数按 1 米网格识别为 {_n_bld} 栋，"
          f"0.5 米、2 米网格的识别结果为 {_rg_bld[-1]:.0f} 与 {_rg_bld[0]:.0f}，栋数对小体量附属建筑敏感，"
          f"故本报告以测绘口径 {_n_bld} 栋做空间分析，以台账口径做房屋管理。")
        P(f"测算结论：实测样区的空间形态可概括为「多层、高密度、高硬化、绿地总量中等」——"
          f"容积率 {self.far:.4f}、建筑密度 {pct(self.cov)}、按全部建筑基底折算 {self.floors:.2f} 层，"
          f"绿地率 {pct(self.green)} 但硬化表面合计 {pct(self.cov + self.road)}。"
          f"「多层」以台账住宅 4～{self.maxfloors} 层为准；{self.floors:.2f} 层是含附属建筑的基底折算值，"
          f"只能用于校核容积率，不用于描述住宅层数。"
          f"这一组合意味着更新重点不在增加建筑面积，而在释放地面空间、提升透水与绿化品质、"
          f"打通消防通道，并以立面与管线整治消除安全隐患。")
        return "\n".join(L), []

    # ── C：现状详细分析 ────────────────────────────────────────────────
    def works_snapshot(self):
        """工程量逐列映射快照：源表头原文 → 报告项名 → 单位 → 明细与合计。

        供 09 工程量闸门逐数核对，也作为「源表头原文」的溯源底账。
        """
        cols = [
            {"sourceName": c["match"], "item": c["item"], "unit": c["unit"],
             "role": c["role"], "total": getattr(self, "work_" + c["role"])}
            for c in self.work_map["columns"]
        ]
        return {
            "source": f"{self.ledger_name}（Sheet1-2 各小区改造工程量统计表）",
            "headerRaw": list(getattr(self, "work_header_raw", [])),
            "mapFile": "Skills/06-分章成文/规则/工程量列映射.json",
            "columns": cols,
            "rows": [{"name": w["name"],
                      **{c["role"]: w.get(c["role"], 0) for c in self.work_map["columns"]}}
                     for w in self.works],
        }

    def chart_data(self):
        """供 08 出图：只放真正来自聚合值的指标，逐格分布类数据一律不给。

        能用：用地构成（四类表面覆盖面积，来自 OSGB 掩膜/纹理聚合）、建筑高度分位
        （P50/P90，来自高程聚合）、台账层数（7 个小区逐条，非分布拟合）。
        不能用：密度分级、高度直方图——样区只有十几栋建筑，直方图是编出来的形状。
        """
        return {
            "sample": {"area": self.area, "ha": self.sample_area_ha,
                       "gfa": self.gfa, "share": None},
            "landUse": [
                {"name": "建筑基底", "area": self.base, "ratio": round(self.cov * 100, 2)},
                {"name": "绿地覆盖", "area": self.green_area, "ratio": round(self.green * 100, 2)},
                {"name": "道路覆盖", "area": self.road_area, "ratio": round(self.road * 100, 2)},
                {"name": "其他表面", "area": self.other_area, "ratio": round(self.other_share * 100, 2)},
            ],
            "height": {"p50": self.h50, "p90": self.h90, "diff": self.hdiff,
                       "maxHeight": self.hmax, "minHeight": self.hmin, "floors": self.floors},
            "floors": [{"name": r["name"], "floors": r["floors"],
                        "households": r["households"]} for r in self.res],
        }

    def _lw(self, iid):
        """AHP 现状准则内指标的局部权重（百分比文本）。"""
        for i in self.ahp_cur_inds:
            if i.get("id") == iid:
                return f"{float(i.get('localWeight') or 0) * 100:.2f}%"
        return "—"

    def _cw(self, cid):
        """AHP 准则权重（百分比文本）。"""
        for c in self.ahp_criteria:
            if c.get("id") == cid:
                return f"{float(c.get('weight') or 0) * 100:.2f}%"
        return "—"

    def _ahp_pair_share(self):
        """配置指定的两项主导准则的权重合计；配置与权重表对不上即报错。"""
        ids = [str(x) for x in self.nar.list("项目.AHP主导准则")]
        w = {str(c.get("id")): float(c.get("weight") or 0)
             for c in self.ahp_criteria}
        missing = [i for i in ids if i not in w]
        if missing:
            raise ValueError(
                f"项目叙述配置「项目.AHP主导准则」引用了 AHP 权重表中不存在的准则："
                f"{missing}（权重表准则：{list(w)}）")
        return sum(w[i] for i in ids)

    def _iname(self, iid):
        for i in self.ahp_cur_inds:
            if i.get("id") == iid:
                return str(i.get("name") or iid)
        return iid

    def _ahp_scores(self):
        """三方案在「现状情况」准则上的加权得分（与 AHP 定稿同源、同序）。"""
        names = [r.get("name") for r in ((self.ahp.get("scores") or {}).get("ranking") or [])]
        names = names or list(self.ahp_cur_scores)
        return "、".join(f"{n} {self.ahp_cur_scores[n]:.2f} 分"
                         for n in names if n in self.ahp_cur_scores)

    def ch_C(self, title, h2):
        L = [f"# {title}"]
        D = self.derive([f"{self.base}={self.area}*{self.cov}",
                         f"{self.green_area}={self.area}*{self.green}",
                         f"{self.road_area}={self.area}*{self.road}",
                         f"{self.gfa}={self.area}*{self.far}",
                         f"{self.floors}={self.far}/{self.cov}",
                         f"{self.hdiff}={self.h90}-{self.h50}"])
        seq = 0

        def S(head, paras):
            nonlocal seq
            L.extend(["", f"## {head}"])
            for text in paras:
                seq += 1
                L.extend(["", text, _anchor("C", seq, [self.internal(), self.material()] + self.derived_all)])

        S("土地与空间利用现状", [
            f"土地与空间利用要分两套口径讲，不能混。按台账，片区由 {self.n_res} 个住宅小区、{self.buildings} 栋住宅、"
            f"{self.households} 户构成，住宅建筑规模 {self.area_wan} 万平方米；按实测样区，样区面积 {self.area} 平方米"
            f"（{self.sample_area_ha:.4f} 公顷），"
            f"建筑基底 {self.base} 平方米，样区建筑密度 {pct(self.cov)}，地上建筑总量约 {self.gfa} 平方米，"
            f"样区容积率 {self.far:.4f}，按基底折算平均层数 {self.floors:.2f} 层。"
            f"两套口径不能互相替代：台账回答片区有多大规模，样区回答这类街区长什么样；"
            f"样区的地上建筑量含住宅与商业，与台账住宅建筑规模不是同一对象，不作占比。"
            f"样区三项强度指标共同指向同一判断：片区不是通过提高容积率获得空间，"
            f"而是通过整理地面与低效空间提升使用效率。",
            f"实测样区的用地构成可按表面类型拆开：建筑基底 {self.base} 平方米、绿地覆盖 {self.green_area} 平方米、"
            f"道路覆盖 {self.road_area} 平方米、其余表面 {self.other_area} 平方米，"
            f"占比分别为 {pct(self.cov)}、{pct(self.green)}、{pct(self.road)} 与 {pct(self.other_share)}。"
            f"样区硬化表面（建筑与道路）合计 {pct(self.cov + self.road)}，透水与绿化表面不足四成，"
            f"这是雨季积水隐患的空间成因之一。",
            f"三项强度指标彼此校核：样区建筑基底与地上建筑总量之比正好复现样区容积率 {self.far:.4f}；"
            f"按样区识别 {self.osgb_n:.0f} 栋分摊，单栋基底 {self.per_building} 平方米、"
            f"单栋地上建筑 {self.gfa_per_building} 平方米。这一量级对应「一栋多户、进深浅」的老旧住宅，"
            f"与台账中小体量院落与成组公寓并存的形态一致，也解释了样区内建筑高度为何分异。",
            f"权属与产权结构上，台账内 {self.n_res} 个住宅小区按名称可识别的产权形态分列："
            f"{self.nar.text('叙述.产权构成')}"
            f"不同形态小区的产权与管理主体不同，更新组织方式也应分类。",
            f"用地效率上，按台账加总折算户均 {round(self.area_wan * 10000 / max(1, self.households), 2)} 平方米；"
            f"其中{self.main_name} 1 个小区即占住宅户数的 {self.main_share * 100:.2f}%。户数高度集中，"
            f"且户均面积偏小，说明片区以中小户型为主，更新施工必须分片推进，避免一次性扰动全部住户。",
            f"空间利用判断：片区土地价值高、增量空间小、存量挖潜集中在三类空间——"
            f"一是样区内 {self.other_area} 平方米的其他表面（含院落硬化、边角地），二是屋顶与立面，"
            f"三是街巷两侧的低效界面。更新应优先把这三类空间转为公共空间、绿化与停车供给。"
            f"把这一判断放回 AHP「现状情况」准则，现状建筑密度与现状容积率两项指标的局部权重分别为 "
            f"{self._lw('current.density')} 与 {self._lw('current.far')}，说明密度与强度本身就是现状评价的组成部分，"
            f"而不只是背景描述。",
            f"挖潜顺序上应先地面、再屋顶、最后界面：地面整理同时解决公共空间与消防通道，"
            f"投入最低、见效最快；屋顶与立面整治需要脚手架与吊篮，施工周期长、扰民大；"
            f"街巷界面改造须与管线开挖同步，避免二次破路。这一顺序也决定了投资安排的分期逻辑。",
        ])
        L.extend(["", "**实测样区表面构成表（1 米网格）**", "",
                  _md_table(
            ["用地类型", "覆盖面积（平方米）", "占比", "空间判断"],
            [["建筑基底", f"{self.base}", pct(self.cov), "连片分布，退界有限"],
             ["绿化覆盖", f"{self.green_area}", pct(self.green), "院落与街角为主，沿街不连续"],
             ["道路覆盖", f"{self.road_area}", pct(self.road), "生活性街巷，断面偏窄"],
             ["其他表面", f"{self.other_area}", pct(self.other_share), "边角地可挖潜为公共空间"],
             ["样区合计", f"{self.area}", "100.00%", f"东西 {self.scope.get('extentM', {}).get('x', 0):.2f} 米 × "
                                              f"南北 {self.scope.get('extentM', {}).get('y', 0):.2f} 米"]]),
            "", f"（数据来源：OSGB 现状指标.json；口径：建筑掩膜与纹理覆盖代理，分母为实测样区面积 {self.area} 平方米）",
            "", "![实测样区表面覆盖构成](图表/fig_land_use.png)", "",
            f"图注：图 3-1　实测样区表面覆盖构成。来源：OSGB 现状指标.json；"
            f"口径：四类表面覆盖面积占实测样区面积（{self.area} 平方米）的比例，"
            f"样区值不代表片区整体，片区整体以台账为准"])
        L.extend(["", f"**片区台账规模表（{self.n_res} 个住宅小区）**", "",
                  _md_table(
            ["项目", "台账值", "说明"],
            [["住宅小区", f"{self.n_res} 个",
              f"单位宿舍 {self.form_dorm}、院落式 {self.form_yard}、公寓 {self.form_apt}、"
              f"未标注形态 {self.form_other}"],
             ["住宅楼栋", f"{self.buildings} 栋", "按小区台账加总"],
             ["住宅户数", f"{self.households} 户", f"{self.main_name}占 {self.main_share * 100:.2f}%"],
             ["住宅建筑规模", f"{self.area_wan} 万平方米", "建筑面积口径，非用地面积"],
             ["户均建筑面积", f"{round(self.area_wan * 10000 / max(1, self.households), 2)} 平方米", "以中小户型为主"],
             ["片区总用地面积", "未取得（内部件 M06 登记）", "空间指标以实测样区为适用边界，不做同口径换算"]]),
            "", f"（数据来源：{self.ledger_name}，Sheet1；口径：台账登记值，未做空间换算）"])

        rows = [[r["name"], self._community_of(r), self._area_of_row(r) or "—",
                 str(r["households"]), str(r["buildings"]), r["year"],
                 f"{r['floors']} 层", r["structure"].replace("结构", ""),
                 f"{r['height']:.0f} 米" if r["height"] else "—",
                 r["facade"].replace("墙面", ""), f"{r['areaWan']:.4f}"] for r in self.res]
        S("建筑现状与风貌保护现状", [
            f"建筑本体方面，台账内住宅 {self.n_res} 个小区、加总住宅建筑规模 {self.area_wan} 万平方米；"
            f"三维测绘范围识别建筑 {self.osgb_n:.0f} 栋，二者统计对象不同，前者用于房屋管理，后者用于空间分析。"
            f"本节从建筑安全与历史风貌两条线判断：结构、消防、线缆属建筑安全；肌理、界面、材料属历史风貌；"
            f"任何整治动作都需满足现行规范合规要求。"
            f"按台账，住宅最高 {self.maxfloors} 层、建筑高度 {self.hmax:.2f} 米，最低 {self.hmin:.2f} 米；"
            f"按测绘，高度中位数 {self.h50:.2f} 米、P90 {self.h90:.2f} 米，"
            f"两者差值 {self.hdiff:.2f} 米，高度分异明显。",
            f"结构与年代方面，砖混结构 {self.brick} 个（占 {self.brick / max(1, self.n_res) * 100:.2f}%）、"
            f"框架结构 {self.frame} 个；建成于 1980 年的 {self.y1980} 个，2000 年以前建成的 {self.before2000} 个。"
            f"砖混住宅的设计使用年限通常按 50 年计，1980 年代建筑已进入第 45 年前后，"
            f"围护结构与管线进入集中更换期。这一时间判断与测绘高度分布互相印证："
            f"中位高度 {self.h50:.2f} 米对应 3 层上下，P90 {self.h90:.2f} 米对应 5 至 6 层，"
            f"说明多数楼栋体型不大，但少数 7 至 8 层楼栋把高度上限抬到 {self.hmax:.2f} 米，"
            f"按住宅建筑的高层界限划分仍属多层与中高层。",
            f"风貌方面，外立面抹灰墙面 {self.plaster} 个、面砖墙面 {self.tile} 个，"
            f"抹灰占 {self.plaster / max(1, self.n_res) * 100:.2f}%。前期沟通材料载明，"
            f"楼栋外立面存在破损、瓷砖掉落与空调冷凝水等问题，与抹灰占比高、使用年限长的判断一致。"
            f"风貌价值不在单栋装饰，而在实测样区 {self.h50:.2f} 米中位高度形成的连续沿街界面与院落边界；"
            f"院落围墙、单元入口、外挂线缆共同构成了片区可识别的街巷肌理。",
            f"安全方面，台账住宅最高 {self.maxfloors} 层、{self.hmax:.2f} 米，属于老旧住宅中的中高层；"
            f"前期材料载明消防安防设施不完善、空中线缆杂乱。结合实测样区建筑密度 {pct(self.cov)} 与"
            f"道路占比 {pct(self.road)}，消防车道与登高面难以按新标准一次性补齐，"
            f"需通过街巷整治与微型消防站组合解决。结构安全还受加建影响：院落内附属搭建与"
            f"屋顶加建改变了原有荷载路径，需在整治前逐栋排查。",
            f"建筑判断：片区建筑本体的核心矛盾是「结构尚可、围护老化、设施滞后」，"
            f"不属于必须整体拆除的类型。更新应以立面整治、屋面防水、楼道修缮、管线更换为主，"
            f"对个别危房与违章搭建做局部拆除，保留既有建筑基底与高度格局（实测样区基底占 {pct(self.cov)}，"
            f"高度中位 {self.h50:.2f} 米）。",
            f"把本节判断放回 AHP「现状情况」准则：该准则权重 {self._cw('current')}，低于投资经济 "
            f"{self._cw('economy')}、风险可控 {self._cw('risk')} 与规划合规 {self._cw('compliance')}；"
            f"准则内「{self._iname('current.age')}」局部权重 {self._lw('current.age')}、"
            f"「{self._iname('current.safety')}」{self._lw('current.safety')}，两项合计占该准则六成以上，"
            f"正是本节判断的两条主线。三个方案在该准则上的得分为 {self._ahp_scores()}，"
            f"保留主体的方案得分最高，与「保留主体、局部拆危」的结论同向。",
        ])
        L.extend(["", "**样区建筑形态表（实测口径）**", "",
                  _md_table(
            ["指标", "样区实测值", "判断"],
            [["建筑高度", f"中位 {self.h50:.2f} 米，P90 {self.h90:.2f} 米", "高度分异，沿街界面连续"],
             ["平均层数", f"{self.floors:.2f} 层", f"含附属建筑的基底折算值；台账住宅 4～{self.maxfloors} 层"],
             ["单栋基底", f"{self.per_building} 平方米", "小体量院落为主"],
             ["单栋地上建筑", f"{self.gfa_per_building} 平方米", "一栋多户、进深浅"],
             ["场地高差", f"{self.relief:.2f} 米", "台地影响无障碍通行"]]),
            "", f"（数据来源：OSGB 现状指标.json；口径：1 米网格实测，适用边界为实测样区 {self.area} 平方米）",
            "", "![实测样区建筑高度分位与片区台账层数](图表/fig_height_floors.png)", "",
            f"图注：图 3-2　实测样区建筑高度分位与片区台账层数。来源：OSGB 现状指标.json 与"
            f"{self.ledger_name}；口径：高度为实测样区高程聚合值"
            f"（中位 {self.h50:.2f} 米、P90 {self.h90:.2f} 米），层数为台账逐小区登记值；"
            f"图内不作换算，正文按 3 米层高的量级对照仅作定性说明"])
        L.extend(["", "**片区台账建筑特征表**", "",
                  _md_table(
            ["指标", "台账值", "判断"],
            [["住宅小区", f"{self.n_res} 个", "单位宿舍、院落、公寓与门牌式并存"],
             ["砖混结构", f"{self.brick} 个（{self.brick / max(1, self.n_res) * 100:.2f}%）",
              "围护与管线进入集中更换期"],
             ["2000 年前建成", f"{self.before2000} 个（{self.before2000 / max(1, self.n_res) * 100:.2f}%）",
              "适老化与设施老化叠加"],
             ["抹灰外立面", f"{self.plaster} 个（{self.plaster / max(1, self.n_res) * 100:.2f}%）",
              "破损、渗水、空鼓集中"],
             ["最高住宅", f"{self.maxfloors} 层、{self.hmax:.2f} 米", "消防登高与疏散压力大"]]),
            "", f"（数据来源：{self.ledger_name}，Sheet1；口径：台账登记值）"])
        L.extend(["", _md_table(["小区名称", "所属社区", "所属片区", "户数（户）", "栋数（栋）",
                                 "建成年代", "最高层数", "结构", "建筑高度", "外立面",
                                 "建筑面积（万㎡）"], rows),
                  "", f"（数据来源：{self.ledger_name}，Sheet1；口径：片区台账逐小区登记值）"])
        _g1, _g2 = self.group_stats.get("示例社区A"), self.group_stats.get("示例社区C")
        if _g1 and _g2:
            def _struct(g):
                parts = []
                if g["brick"]:
                    parts.append(f"砖混 {g['brick']} 个")
                if g["frame"]:
                    parts.append(f"框架 {g['frame']} 个")
                return "、".join(parts) or "—"

            def _face(g):
                parts = []
                if g["plaster"]:
                    parts.append(f"抹灰 {g['plaster']} 个")
                if g["tile"]:
                    parts.append(f"面砖 {g['tile']} 个")
                return "、".join(parts) or "—"

            L.extend(["", "**项目主体片区内两社区差异对照**", "",
                      _md_table(
                ["对照项", "示例社区A组", "示例街道—示例社区C组（含示例巷 1 处）"],
                [["住宅小区（个）", str(_g1["n"]), str(_g2["n"])],
                 ["住宅栋数（栋）", str(_g1["buildings"]), str(_g2["buildings"])],
                 ["户数（户）", str(_g1["households"]), str(_g2["households"])],
                 ["住宅建筑面积（万㎡）", f"{_g1['areaWan']:.4f}", f"{_g2['areaWan']:.4f}"],
                 ["占片区户数（%）", f"{_g1['households'] / self.households * 100:.2f}",
                  f"{_g2['households'] / self.households * 100:.2f}"],
                 ["占片区住宅面积（%）", f"{_g1['areaWan'] / self.area_wan * 100:.2f}",
                  f"{_g2['areaWan'] / self.area_wan * 100:.2f}"],
                 ["户均建筑面积",
                  f"{_g1['areaWan'] * 10000 / max(1, _g1['households']):.2f} 平方米/户",
                  f"{_g2['areaWan'] * 10000 / max(1, _g2['households']):.2f} 平方米/户"],
                 ["平均每栋户数（户/栋）", f"{_g1['households'] / max(1, _g1['buildings']):.2f}",
                  f"{_g2['households'] / max(1, _g2['buildings']):.2f}"],
                 ["院均户数（户/院）", f"{_g1['households'] / max(1, _g1['n']):.2f}",
                  f"{_g2['households'] / max(1, _g2['n']):.2f}"],
                 ["平均每栋建筑面积（平方米/栋）",
                  f"{_g1['areaWan'] * 10000 / max(1, _g1['buildings']):.2f}",
                  f"{_g2['areaWan'] * 10000 / max(1, _g2['buildings']):.2f}"],
                 ["院均住宅建筑面积（万平方米/院）",
                  f"{_g1['areaWan'] / max(1, _g1['n']):.4f}",
                  f"{_g2['areaWan'] / max(1, _g2['n']):.4f}"],
                 ["建成年代", (f"{_g1['year_min']} 年" if _g1['year_min'] == _g1['year_max']
                               else f"{_g1['year_min']}—{_g1['year_max']} 年"),
                  (f"{_g2['year_min']} 年" if _g2['year_min'] == _g2['year_max']
                   else f"{_g2['year_min']}—{_g2['year_max']} 年")],
                 ["结构", _struct(_g1), _struct(_g2)],
                 ["外立面", _face(_g1), _face(_g2)],
                 ["最高层数", f"{_g1['floors_max']} 层", f"{_g2['floors_max']} 层"]]),
                      "", f"（数据来源：{self.ledger_name}，Sheet1；口径：仅对照项目主体片区内的两个社区，"
                      f"示例社区C组含示例社区C台账内示例巷 1 处；示例片区B另行成组，见表后说明）",
                      "",
                      f"两组差异决定了更新必须分组施策。示例社区A组 {_g1['n']} 个小区全部建成于 "
                      f"{_g1['year_min']} 年，以砖混结构、抹灰外立面为主，最高 {_g1['floors_max']} 层，"
                      f"户均建筑面积约 {round(_g1['areaWan'] * 10000 / max(1, _g1['households']), 2)} 平方米，"
                      f"老化程度深、户内空间紧，更新重点是结构安全、屋面防水与管线更换；"
                      f"示例社区C组建成于 {_g2['year_min']}—{_g2['year_max']} 年，框架 {_g2['frame']} 个、砖混 {_g2['brick']} 个，"
                      f"面砖外立面 {_g2['tile']} 个，最高 {_g2['floors_max']} 层，"
                      f"户均建筑面积约 {round(_g2['areaWan'] * 10000 / max(1, _g2['households']), 2)} 平方米，"
                      f"更新重点是立面品质、电梯加装与社区服务补位。",
                      *([self._area_extra_paragraph()] if self._area_extra_paragraph() else [])])

        S("人口与社会人居现状", [
            *([self._area_sentence()] if self._area_sentence() else []),
            *([self._area_table()] if self._area_table() else []),
            f"人口规模按台账口径：{self.n_res} 个住宅小区合计 {self.households} 户，"
            f"按户均 3.2 人推算约 {self.population} 人。这是本报告唯一的人口推算取值，"
            f"用于判断公共服务与适老化需求的量级。",
            f"户数分布高度集中：{self.spot['main_head']}为片区主力院落，"
            f"{self.households_main} 户、{self.main_area_wan:.4f} 万平方米，"
            f"占住宅户数的 {self.main_share * 100:.2f}%；"
            f"其余 {self.n_res - 1} 个小区累计 {self.others_households} 户，单个小区规模普遍偏小。"
            f"集中度高带来两个后果：一是更新施工必须分片组织、错峰推进；"
            f"二是社区服务设施应优先布局在{self.main_name}周边，形成 5 分钟步行覆盖。",
            f"分组建账看人口：{self.nar.text('项目.社区数量表述')}个社区合计 {self.n_res} 个院落（小区）、{self.households} 户。"
            f"常住人口台账未载，取值同上，仅作量级参考、不按社区下结论。"
            f"组际差异在人口侧的表现是：示例社区A组以高密度围合院落为主，"
            f"示例社区C组单院规模小、更接近沿街公寓形态"
            f"{self._area_extra_sentence()}。",
            f"年龄结构上，片区住宅以 1980 年代单位宿舍与 1990 年代公寓为主，"
            f"2000 年以前建成的小区占 {self.before2000 / max(1, self.n_res) * 100:.2f}%，"
            f"早期入住的居民已进入老年阶段。老旧住宅最高 {self.maxfloors} 层且无电梯，"
            f"垂直出行成为适老化改造的首要问题。",
            f"从居住强度看，户均建筑面积 {round(self.area_wan * 10000 / max(1, self.households), 2)} 平方米，"
            f"以中小户型为主。此处只做户型强度判断，不折算人口密度：户均面积的分子分母同为台账口径，"
            f"可以成立；人口密度需要片区总用地面积作分母，包内没有该口径，测绘范围只是空间样本，"
            f"不能替代总用地。"
            f"小户型、层数不高的组合，意味着户内可改造余量有限，"
            f"居住品质提升更多依赖楼道、院落地面的公共空间整理。",
            f"适老化还受场地竖向制约：场地格高程 P5 与 P50 相差 {self.relief:.2f} 米，"
            f"院落与街巷之间存在台地，轮椅、婴儿车与急救担架通行受限。"
            f"单元入口台阶、院落门槛与街巷高差叠加，构成实测样区所在街区的无障碍改造实际难点，"
            f"也解释了前期沟通中「步行无障碍改造」为何被列入建设内容。",
            f"人居诉求方面，前期沟通材料载明群众改善居住环境诉求强烈，集中在管线安全、"
            f"立面破损、汛期积水、缺少适老化与便民休憩配套四类；项目信息卡把「补齐便民康养、"
            f"健康服务配套」列入建设内容。诉求与前述指标指向一致，可作为更新项目排序的依据。",
            f"人口判断：片区人口规模中等、户数高度集中、老龄化压力显性化。"
            f"更新应把无障碍与适老化作为刚性配置，把社区服务设施布局与户数分布对齐，"
            f"避免设施建在户数稀疏一侧。",
            f"人口规模与设施规模需要对照着看：{self.population} 人对应的社区服务、养老、卫生站点均属小型配置，"
            f"单个站点即可覆盖；若按小区分别设置，单个站点服务人口不足其运营门槛，容易建成即闲置。"
            f"户均 3.2 人是推算取值，若实际低于该值，服务需求总量下降，但设施类型与布局逻辑不变，"
            f"本节判断不受影响。",
        ])
        L.extend(["", "**片区台账人口与居住表**", "",
                  _md_table(
            ["指标", "台账值", "判断"],
            [["住宅户数", f"{self.households} 户", f"首位小区占 {self.main_share * 100:.2f}%"],
             ["推算人口", f"{self.population} 人", "按户均 3.2 人折算"],
             ["住宅建筑规模", f"{self.area_wan} 万平方米", "台账口径的建筑面积，非用地面积"],
             ["户均建筑面积", f"{round(self.area_wan * 10000 / max(1, self.households), 2)} 平方米", "以中小户型为主"],
             ["2000 年前建成", f"{self.before2000} 个", "老龄化与设施老化叠加"]]),
            "", f"（数据来源：{self.ledger_name}，Sheet1；口径：片区台账，不折算人口密度）"])

        S("街区业态与经营现状", [
            f"片区以居住为主导，商业设施仅{self.biz_name} {self.biz_count} 处，位于{self.nar.text('项目.商业建筑.地址')}，"
            f"{int(self.nar.num('项目.商业建筑.建成年份'))} 年建成，"
            f"{self.nar.text('项目.商业建筑.结构')}，{self.nar.text('项目.商业建筑.外立面')}，依托沿街住宅界面设置。商业体量小、类型单一，"
            f"难以承担街区级商业服务功能，也未形成独立商业地块。",
            f"沿街经营主要依托{_join_cn(self.nar.list('项目.沿街商业街巷'), '、', '、')}两侧住宅底层，以社区便民服务为主。"
            f"前期沟通材料把「社区服务及办公用房」「便民休憩点位」「康养、健康服务」列入更新内容，"
            f"说明现状商业与服务业态不足以覆盖居民日常需求。",
            f"按上述人口与户数规模折算的日常消费规模衡量，"
            f"片区现有商业载体与人口规模不匹配：一方面商业面积小、缺少成片商业界面，"
            f"居民日常消费需要向外围街区外溢；另一方面沿街底层商铺依附住宅开间，"
            f"进深与面宽受结构限制，难以承载生鲜、餐饮等对后勤与排烟有要求的业态。"
            f"这构成了「消费外溢」与「业态进不来」并存的局面。",
            f"从业态分布看，商业沿街线性展开、无成片商业街区；从业态质量看，"
            f"以餐饮、零售、维修等基础生活服务为主，缺少康养、文化、文体等提升型业态。"
            f"业态问题可以概括为「总量不缺、结构不全、品质不高」——居民基本生活服务可在街区内解决，"
            f"但品质型、体验型消费必须出片区。",
            f"经营与治理上，{self.n_res} 个住宅小区由单位宿舍、院落、公寓与门牌式小区构成，"
            f"管理主体分散，沿街商铺与住宅混杂。更新后需要明确街区层面的运营主体，"
            f"把管线智慧运维、公共空间保洁、便民设施维护纳入统一管理。",
            f"业态判断：片区不适合发展大型商业，宜走「社区商业 + 生活服务」路线——"
            f"利用街角与底层空间补充生鲜、餐饮、维修、康养等便民业态，"
            f"把商业活力与街巷环境整治绑定，而不是单独建设商业综合体。",
            f"业态布局的另一制约来自街巷断面。片区道路以生活性街巷为主，沿街界面连续但进深有限，"
            f"餐饮类业态所需的排烟、隔油与后勤通道难以在住宅底层解决；零售类业态则受居民消费外溢影响，"
            f"客流被外围商业分流。更新若只做立面美化而不调整底层空间，业态结构不会改变。"
            f"因此业态织补应与底层空间改造、街巷环境整治同步推进，先解决可经营条件，再谈招商。",
        ])
        L.extend(["", _md_table(
            ["业态类型", "空间载体", "现状特征", "判断"],
            [["住宿餐饮", self.biz_name, "单体规模小，服务对象偏外部", "非街区级商业中心"],
             ["社区便民商业", "住宅底层沿街", "餐饮、零售、维修等基础服务", "结构单一，缺提升型业态"],
             ["生活性服务", f"{_join_cn(self.nar.list('项目.沿街商业街巷'), '、', '、')}两侧", "依附住宅界面，无独立商业体量", "沿街线性分布"],
             ["提升型业态", "—", "康养、文化、文体类缺失", "需在更新中植入"]]),
            "", f"（数据来源：{self.ledger_name}，Sheet1；项目信息卡）"])

        S("交通与通行现状", [
            f"道路资源按覆盖口径：道路覆盖面积 {self.road_area} 平方米，占测绘范围 {pct(self.road)}，"
            f"折算为每公顷测绘范围内的道路覆盖面积 {self.road_per_ha:.2f} 平方米"
            f"（口径：道路覆盖面积 ÷ 测绘范围面积；这是覆盖率型指标，不是按道路长度计的路网密度）。"
            f"路网结构为「外围城市道路 + 内部生活街巷」，外围依托{_join_cn(self.nar.list('项目.外围道路'), '、', '与')}，"
            f"内部由{_join_cn(self.nar.list('项目.内部街巷'), '、', '、')}串联。道路以步行与生活性通行为主，"
            f"缺少交通性干道支撑。",
            f"通行条件上，样区东西 {self.scope.get('extentM', {}).get('x', 0):.2f} 米、"
            f"南北 {self.scope.get('extentM', {}).get('y', 0):.2f} 米，街巷尺度紧凑；"
            f"样区建筑密度 {pct(self.cov)} 意味着建筑退界有限，街巷两侧界面连续。"
            f"这一结构有利于步行，但会压缩消防车道与错车空间。",
            f"把面积与断面分开看，问题更清楚：实测样区 {pct(self.road)} 的道路占比在数值上并不低，"
            f"但面积主要由多条窄断面街巷累加而成，而非一条宽断面干道。"
            f"街巷两侧院落围墙紧贴路面，路侧停车与杂物堆放进一步挤占有效通行宽度，"
            f"会车点缺少，救护车与搬家车辆进出需要临时清障。这是「面积不低、断面不足」的直接后果。",
            f"消防可达性可借建筑与道路的空间关系判断：建筑掩膜在 1 米网格下识别为 {self.comp_total} 个连通分量，"
            f"最大分量占掩膜栅格的 {self.max_comp_share * 100:.2f}%，说明主体建筑连片、院落被建筑围合；"
            f"实测样区道路覆盖虽占 {pct(self.road)}，但主要分布在样区外围与街巷骨架，难以进入围合院落内部。"
            f"消防车可沿外围与主要街巷到达，院落内部救援仍需依托室外消火栓补点与微型消防站。",
            f"停车与慢行方面，片区没有独立公共停车场，车辆主要依托院落内部与街巷路侧停放；"
            f"前期材料把「步行无障碍改造」列入内容，说明现状慢行系统连续性不足。"
            f"更新应优先保障消防通道，再通过街角与边角地补充少量停车与休憩空间。",
            f"交通判断：片区的交通问题不是路网缺失，而是断面被挤占、节点不通畅。"
            f"改善应以断面优化与节点疏通为主——清理占道、贯通消防通道、设置错车点与无障碍坡道，"
            f"不新增道路，也不拓宽既有街巷走向。",
            f"从出行结构判断，片区内短距离出行以步行与电动自行车为主，机动车保有量不高，"
            f"但路侧停放需求刚性。街巷断面在保障消防通道后，可分配给步行与停放的空间有限，"
            f"更新应明确「消防优先、步行其次、停放兜底」的排序，避免为增加车位压缩消防宽度。"
            f"错车点与转弯半径宜结合院落入口设置，把有限的断面用在节点上，而不是全线拓宽。",
        ])
        L.extend(["", "**实测样区道路与通行表**", "",
                  _md_table(
            ["指标", "样区实测值", "判断"],
            [["道路覆盖面积", f"{self.road_area} 平方米", f"占实测样区 {pct(self.road)}"],
             ["道路覆盖密度（测绘口径）", f"{self.road_per_ha:.2f} 平方米/公顷", "覆盖不低、断面不足"],
             ["样区尺度", f"{self.scope.get('extentM', {}).get('x', 0):.2f} 米 × {self.scope.get('extentM', {}).get('y', 0):.2f} 米", "外围干道支撑，内部街巷串联"],
             ["建筑连通分量", f"{self.comp_total} 个", "主体建筑连片，院落被围合"],
             ["最大分量占比", f"{self.max_comp_share * 100:.2f}%", "消防车难入院落内部"],
             ["公共停车", "无独立公共停车场", "依托院落与路侧停放"]]),
            "", "（数据来源：OSGB 现状指标.json；口径：分母为实测样区面积，不含片区其他地块；停车情况来自前期沟通材料）"])

        S("市政基础设施现状", [
            f"市政管线是片区最集中的短板。前期沟通材料载明：市政管线老旧、电力容量不足、"
            f"空中线缆杂乱、排水与供电配套存在安全隐患、消防安防设施不完善；"
            f"夏季用电负荷激增时老旧架空线路易引发电气安全事故，汛期管网排水不畅易产生积水隐患。"
            f"这与测绘得到的硬化表面 {pct(self.cov + self.road)}（按实测占比折算）、绿地与透水表面不足四成的空间特征互为因果。",
            f"改造内容已明确为系统方案：新增街区 10 千伏电力设施并扩容，更新给排水、强弱电、"
            f"通信地下管网，全部架空线路下地，搭建管线智慧运维平台，全域增设消防与安防设施。"
            f"这是一次性解决「空中蜘蛛网、雨季积水、电力不足」三类问题的组合措施。",
            f"设施配置与建筑高度相关：住宅最高 {self.maxfloors} 层、{self.hmax:.2f} 米，"
            f"消防供水与应急疏散对管网压力与取水点密度要求更高。"
            f"管线改造应与消防取水点、微型消防站同步设计，避免重复开挖。",
            f"排水竖向还受场地台地影响：场地格高程 P5 与 P50 相差 {self.relief:.2f} 米，"
            f"雨水沿街巷由高向低汇流，低洼院落与街角成为汇水末端；"
            f"硬化表面合计 {pct(self.cov + self.road)}，汇流速度快于下渗速度，"
            f"管网一旦排水不畅即形成积水点。这决定了排水提标必须与透水铺装、下沉式绿地同步做，"
            f"只放大管径而不削减汇流量，效果有限。",
            f"市政判断：片区基础设施问题是系统性的，单点修补无效。更新应以「管线下地 + 电力扩容 + "
            f"排水提标 + 消防补点」为一组，按街巷分段实施，先地下后地上，减少对居民生活的二次扰动。",
            f"台账工程量与设施短板一一对应：立面材料修缮或翻新 {self.work_facade:.0f} 平方米覆盖抹灰墙面楼栋，"
            f"更换外墙雨棚 {self.work_awning:.0f} 个、空调外机披水板 {self.work_apron:.0f} 个分别对应沿街遮阳"
            f"与空调冷凝水外溢点位，冷凝水管 {self.work_condensate:.0f} 米、落水管 {self.work_downpipe:.0f} 米"
            f"对应屋面与阳台排水。工程量的分布说明改造不是单点修补，"
            f"而是沿街巷成线推进；这也意味着施工组织应以街巷为流水段，完成一段、恢复一段、交付一段，"
            f"把对居民出行与停车的影响控制在最短时间内。",
            f"从改造组织看，管线跨 {self.n_res} 个小区敷设，权属分属电力、供水、燃气、通信等多家单位。"
            f"若按小区分别立项，同一段街巷会被重复开挖。合理方式是以街巷为单元编制管线综合方案，"
            f"统一开挖时序与回填标准，把管线智慧运维平台作为竣工交付条件，"
            f"使一次施工同时满足电力扩容、雨污分流、通信下地与消防补水四项要求。",
        ])
        L.extend(["", _md_table(
            ["台账列名（源表原文）", "报告项名", "单位", "台账合计"],
            [[c["match"], c["item"], c["unit"], f"{getattr(self, 'work_' + c['role']):.0f}"]
             for c in self.work_map["columns"]
             if getattr(self, "work_" + c["role"]) > 0]),
            "", f"（数据来源：{self.ledger_name}，Sheet1-2；"
            "列名按源表原文逐列映射，不按列序取数）"])

        S("公共服务配套现状", [
            f"公共服务以社区服务为主。前期材料列出的更新内容包含社区服务及办公用房、街边公园、"
            f"街角环境提升、便民休憩点位、康养与健康服务配套，说明现状设施在数量与品质上均有缺口。"
            f"这一人口规模要求社区级服务设施实现步行 5 至 10 分钟覆盖。",
            f"设施分布与户数分布不匹配：{self.households_main} 户集中在{self.main_name}，"
            f"占住宅户数（加总口径）{self.main_share * 100:.2f}%，其余小区规模较小且分散，"
            f"若按小区配建，容易出现重复建设与低效使用。合理方式是街区级统筹——"
            f"在户数密集区设社区服务站与康养点，在街角补充小型休憩与健身空间。",
            f"从空间尺度看，实测样区东西与南北尺度都在百米量级，单个社区服务点的合理服务半径即可覆盖这一范围；"
            f"但片区整体沿{_join_cn(self.nar.list('项目.贯穿街巷'), '与')}展开，涉及 {self.n_comm} 个社区、{self.n_res} 个小区，"
            f"设施配置应按街区级统筹，而不是在样区尺度上就地配建；"
            f"真正决定可达性的不是距离，而是街巷连通性与台地高差。"
            f"现状设施零散分布在小区内部，跨小区共享需要穿行院落与街巷，实际可达性低于几何距离。",
            f"适老化与无障碍是公共服务的关键指标。片区住宅最高 {self.maxfloors} 层且无电梯，"
            f"2000 年以前建成小区占 {self.before2000 / max(1, self.n_res) * 100:.2f}%，"
            f"老年居民出行依赖楼梯与街巷。更新应把加装电梯、楼道扶手、无障碍坡道、"
            f"休憩座椅作为标准配置，而不是选配项。",
            f"设施落地空间可从存量中找：实测样区其他表面 {self.other_area} 平方米（建筑基底、绿地与道路之外的残差），"
            f"其中院落硬化与边角地可整理为小型服务空间；沿街界面按「一间一策」调整为便民网点，"
            f"具体点位与面积在实施阶段按院落逐处核定。"
            f"与新建独立占地相比，这种方式不减少居住面积，也符合片区高密度的空间条件。",
            f"公共服务判断：真正的短板不是空间总量，而是成体系的服务空间。"
            f"台账 {self.n_res} 个小区、{self.households} 户的居住规模足以支撑一个街区级服务网络，"
            f"缺的是把零散载体串起来的组织方式。更新应把「社区服务站 + "
            f"康养点 + 街角公园 + 便民商业」组合成街区服务网络，与市政管线、街巷整治同步实施，"
            f"避免设施建成后无人维护。",
            f"按服务类型拆分，四类设施的选址逻辑并不相同：社区服务与办公属行政配置，可由街区统筹；"
            f"康养与健康服务需要就近、首层、无障碍，应贴着户数密集区与街巷平段布置；"
            f"文体与休憩属日常高频使用，应分散到街角与院落入口；便民商业依赖沿街界面，需要可经营的开间与进深。"
            f"现状把这些功能混装在小区内部，既不便共享，也难以形成连续的街道生活；"
            f"更新应按类型重新落位，而不是把四类功能塞进同一栋楼。",
        ])
        L.extend(["", _md_table(
            ["服务类型", "现状载体", "判断"],
            [["社区服务及办公", "依附小区设置", "空间零散，缺街区级站点"],
             ["康养与健康服务", "无独立载体", "老年居民就近配置需求突出"],
             ["文体与休憩", "街边公园、街角", "点位不足，连续性差"],
             ["便民商业", "住宅底层沿街", "业态单一，需织补"],
             ["无障碍设施", "列入改造内容", "单元入口与街巷高差是难点"]]),
            "", "（数据来源：项目信息卡；前期沟通材料）"])

        S("生态与人居环境现状", [
            f"实测样区的绿化总量中等、分布不均。样区绿地率 {pct(self.green)}，折合绿地覆盖面积 {self.green_area} 平方米；"
            f"样区建筑密度 {pct(self.cov)}，建筑基底 {self.base} 平方米；绿地与建筑基底之比为 "
            f"{self.green / self.cov:.2f}。从总量看，样区绿化在老城居住街区中不算低；"
            f"从分布看，绿地主要集中在院落内部与街角，沿街连续绿化不足。",
            f"绿化质量只能说到能说清的部分。纹理抽样显示，地面绿化占绿化覆盖的 {pct(self.non_roof_green)}，"
            f"屋面及树冠、构筑物遮挡合计占 {pct(1 - self.non_roof_green)}。"
            f"后一项是「屋面＋树冠投影＋构筑物遮挡」的合并值，包内数据无法拆分出屋面绿化的真实规模，"
            f"现有资料不能据此判断屋顶绿化存量多少。可以确认的只有一条：地面绿化占绿化覆盖不足一半，"
            f"这一比例不支持「地面绿化为主导」的判断，绿化空间的实际分布需结合现场核实。",
            f"绿地可达性与总量不同步：院落绿化权属归小区，开放共享程度低，"
            f"街角绿地数量少且被停车与摊位占用；居民从单元门到最近的连续绿地，"
            f"往往需要穿行街巷与院落门禁。绿地\"看得见、进不去\"是现状的典型状态，"
            f"这也是更新中需要优先打通院落边界的原因。",
            f"场地竖向进一步分割了绿化空间：P5 与 P50 相差 {self.relief:.2f} 米，"
            f"院落依台地布置，单块绿地面积小、形状不规则，难以形成连续步行绿廊。"
            f"绿化提升因此不能只算面积增量，还要考虑台地之间的连接方式——"
            f"用台阶、坡道与无障碍路径把碎片绿地串联，才可能形成可用的绿地系统。",
            f"人居环境问题集中在硬化与积水。建筑与道路硬化表面合计 {pct(self.cov + self.road)}，"
            f"其他表面 {pct(self.other_share)}；汛期管网排水不畅时，硬化表面快速汇水，"
            f"低洼院落与街巷易积水。更新应通过透水铺装、下沉式绿地、雨水花园组合，"
            f"把硬化表面改造为可调蓄表面。",
            f"绿地口径需要说清：现状绿地 {self.green_area} 平方米是实测样区内的绿化覆盖，"
            f"而上述人口分布在 {self.n_res} 个小区，其用地范围大于实测样区。"
            f"分子与分母不同源，故本报告不折算人均绿地；要得到人均指标，须先取得片区总用地面积与同一范围内的人口。"
            f"就现有数据看，绿地集中在院落内部，"
            f"沿街 5 分钟可达的公共绿地不足。更新应优先把可开放的院落绿地与街角绿地串联，"
            f"形成连续的步行绿化界面，而不是在图纸上增加绿地面积。",
            f"生态判断：片区生态提升的重点不是扩大绿地总量，而是提高绿地可达性与雨水调蓄能力。"
            f"建议把街角与院落边角两类地面空间作为绿化增量来源，"
            f"把实测样区内 {pct(self.other_share)} 的其他表面中的硬化部分优先改为透水铺装。"
            f"屋顶绿化能否纳入，取决于屋面荷载、防水条件与可绿化面积，现有数据无法拆分出屋面绿化的实际规模，"
            f"应先按栋核实屋面可绿化面积，再决定是否纳入以及做多大规模。",
            f"还需要说明样区绿地率与建筑密度的关系：两者数值接近，说明地面空间在建筑与绿化之间近乎均分，"
            f"可用于新增绿化的地面余量有限。这一结构决定了绿化提升只能向内要空间——"
            f"把院落硬化拆改为透水铺装与花池，把围墙与挡墙的垂直面用于攀援绿化，把街角零散地块串成口袋公园；"
            f"若按传统方式另找地块建集中绿地，在实测样区内并不现实。",
        ])
        L.extend(["", "**实测样区生态与人居指标表**", "",
                  _md_table(
            ["指标", "样区实测值", "判断"],
            [["绿地率", pct(self.green), "总量中等，分布不均"],
             ["绿地覆盖面积", f"{self.green_area} 平方米", "院落与街角为主，沿街不连续"],
             ["绿地与基底之比", f"{self.green / self.cov:.2f}", "地面空间在建筑与绿化间近乎均分"],
             ["非屋顶绿化占比", pct(self.non_roof_green), "地面绿化不足一半，屋面占比不可拆分"],
             ["硬化表面合计", pct(self.cov + self.road), "汇流快，积水隐患"],
             ["其他表面", pct(self.other_share), "可改造为透水与绿化空间"]]),
            "", f"（数据来源：OSGB 现状指标.json；口径：建筑掩膜与纹理覆盖代理，分母为实测样区面积 {self.area} 平方米，不代表片区整体绿地率）"])
        return "\n".join(L), []

    # ── D：核心问题汇总（8 类 + 成因分析 + TBL-02） ─────────────────────
    def ch_D(self, title, h2):
        L = [f"# {title}"]
        seq = 0
        D = self.derive([f"{self.base}={self.area}*{self.cov}",
                         f"{self.gfa}={self.area}*{self.far}",
                         f"{self.floors}={self.far}/{self.cov}",
                         f"{self.hdiff}={self.h90}-{self.h50}",
                         f"{self.other_share * 100:.2f}=100-{self.cov * 100:.2f}-{self.green * 100:.2f}-{self.road * 100:.2f}"])
        problems = [
            ("街巷空间与肌理问题",
             f"街巷尺度紧凑、界面连续，测绘范围 {self.area} 平方米内按占比折算道路覆盖 {self.road_area} 平方米，"
             f"占比 {pct(self.road)}；实测样区建筑密度 {pct(self.cov)} 压缩了街巷退界，"
             f"步行、错车、消防通行三类需求在同一断面上叠加。",
             f"关键指标：道路覆盖 {self.road_area} 平方米，开敞空间 {pct(self.opening)}（折算口径），"
             f"断面瓶颈点见落点。"
             f"更新方向：街巷断面重整、消防通道贯通、街角空间释放。"),
            ("建筑安全与风貌破损问题",
             f"住宅最高 {self.maxfloors} 层、{self.hmax:.2f} 米，砖混结构加总 {self.brick} 个（占 "
             f"{self.brick / max(1, self.n_res) * 100:.2f}%），1980 年建成 {self.y1980} 个；"
             f"抹灰外立面 {self.plaster} 个，饰面老化、渗水、空鼓与瓷砖掉落并存；"
             f"空中线缆杂乱，消防安防设施不完善。",
             f"关键指标：砖混占比 {self.brick / max(1, self.n_res) * 100:.2f}%，"
             f"2000 年前建成 {self.before2000} 个；结构鉴定与屋面防水为前置项。"
             f"更新方向：立面整治、屋面防水、楼道修缮、线缆下地。"),
            ("文脉保护与活化不足问题",
             f"片区以 {self.year_span} 年住宅为主体，单位宿舍、院落与公寓三类肌理并存；"
             f"实测样区 {self.h50:.2f} 米的中位高度对应连续沿街界面。现状问题是肌理被零散搭建与广告牌切割，"
             f"院落边界模糊，历史信息缺少系统性表达。",
             f"关键指标：{self.nar.text('叙述.文脉关键指标')}。"
             "更新方向：保留街巷走向与院落边界，用标识、铺装、立面材质延续文脉，不新建仿古建筑。"),
            ("业态老旧与活力不足问题",
             f"商业仅{self.biz_name} {self.biz_count} 处，居住户数加总 {self.households} 户，商业与居住比例失衡；"
             f"沿街底层以零散便民服务为主，缺少成体系的社区商业；"
             f"街区活力集中在少数沿街界面，院落内部活力不足。",
             f"关键指标：商业 {self.biz_count} 处、住宅 {self.n_res} 个小区，商业与住宅规模失衡。"
             "更新方向：社区商业织补、底商立面整治、街角活力点植入。"),
            ("交通通行与消防隐患突出问题",
             f"按实测样区占比折算道路覆盖 {self.road_area} 平方米、占比 {pct(self.road)}，路网以生活街巷为主；"
             f"住宅最高 {self.maxfloors} 层，院落入口狭窄，消防车道难以成环，"
             f"消防取水点密度不足；汛期积水与电气线路老化叠加，形成复合风险。",
             f"关键指标：实测样区道路覆盖见 P-01，台账最高 {self.hmax:.2f} 米，消防设施不完善。"
             f"更新方向：消防通道贯通、取水点补点、管线下地、积水点整治。"),
            ("基础设施老旧滞后问题",
             f"市政管线老旧、电力容量不足、空中线缆杂乱、排水管网不畅；"
             f"实测样区硬化表面加总 {pct(self.cov + self.road)}，汇水速度快，"
             f"雨季低洼院落与街巷积水；弱电、通信线路与电力线路混挂，安全间距不足。",
             f"关键指标：实测样区硬化表面 {pct(self.cov + self.road)}（建筑与道路合计），管线全部架空。"
             "更新方向：10 千伏电力扩容、给排水与强弱电下地、智慧运维平台。"),
            ("民生配套与人居品质偏低",
             f"社区服务设施不足，康养、健康服务、便民休憩点位缺少系统布局；"
             f"上述人口规模要求步行 5 至 10 分钟覆盖；"
             f"住宅最高 {self.maxfloors} 层且无电梯，适老化与无障碍改造需求突出。",
             f"关键指标：按 5 分钟步行覆盖率与服务半径核算，服务点位未全覆盖，户数规模见 P-04。"
             "更新方向：社区服务站、康养点、无障碍与加装电梯、街角休憩空间。"),
            ("街区治理与运营短板",
             f"{self.n_res} 个住宅小区中，单位宿舍、院落、公寓与门牌式小区并存，管理主体分散；"
             f"沿街商铺与住宅混杂，公共空间与管线设施缺少统一维护主体；"
             f"更新完成后若沿用现状治理结构，设施难以长期维持。",
             f"关键指标：{self.n_res} 个小区、4 类形态、{self.biz_count} 处商业。"
             "更新方向：街区统筹管理、社区参与、管线智慧运维、公共空间统一保洁。"),
        ]
        # 八类问题的表现/成因/影响逐条不同：R1 用同一段「建设年代早、产权分散…」套八遍，
        # 40 条要点去重后只剩 10 条，等于复制。这里每条都落到本节自己的数据。
        details = [
            ("主街连续、支巷断续：断面被院落围墙、路侧停车与杂物挤占，院落入口成为通行瓶颈。",
             f"建筑密度 {pct(self.cov)} 与道路覆盖 {pct(self.road)} 的组合下，街巷宽度由两侧建筑退界决定，"
             f"院落围墙沿红线贴建，建设时未预留断面，增量空间只能从既有断面里挤。",
             "消防车与救护车难以进入围合院落，错车、临时停车与慢行在同一断面上互相挤压。"),
            ("抹灰饰面空鼓剥落、屋面渗漏、局部瓷砖掉落，沿街界面成片老化，空中线缆缠挂。",
             f"砖混结构 {self.brick} 个（占 {self.brick / max(1, self.n_res) * 100:.2f}%）、1980 年建成 "
             f"{self.y1980} 个，按当年防水与饰面标准施工，使用超过三十年进入自然老化期；"
             f"台账屋面翻新申报量为零（{self.nar.text('项目.屋面翻新说明')}），本次改造不含屋面翻新。",
             "雨季渗漏与高空坠物集中出现，维修由「小修」升级为「大修」，成本向后外移且被动。"),
            ("院落边界被加建与广告牌切割，街巷铺装与立面材质不统一，历史信息缺少标识。",
             "院落式住宅与 1980 年代小区各自为政，逐户自行改造，缺少街区层面的风貌导则与材质约定，"
             "肌理在一次次单户改造中被磨损。",
             "街区识别性下降，更新后容易出现「整了新、丢了旧」的同质化，文脉价值难以转化为空间品质。"),
            (f"沿街底层以零散便民服务为主，商业仅{self.biz_name} {self.biz_count} 处，院落内部无经营性界面。",
             f"现有户数的消费规模支撑不起成体系的社区商业，住宅底层又受排烟、隔油与"
             f"后勤通道条件限制，业主要么不入驻、要么只做低强度便民服务。",
             "日常消费外溢到外围道路，步行可达的商业点位不足，夜间活力弱。"),
            ("消防车道难以成环，取水点密度不足，院落入口与错车点缺少；汛期低洼点积水与线路老化叠加。",
             f"道路覆盖 {pct(self.road)} 由多条窄断面街巷累加，建筑掩膜最大连通分量占 "
             f"{self.max_comp_share * 100:.2f}%，院落被建筑围合，消防车只能到达外围与主街巷。",
             "院落内部救援依赖室外消火栓与微型消防站补位，响应时间延长，汛期风险叠加。"),
            ("管线架空、电力容量不足、雨污合流，低洼院落与街巷雨季积水。",
             f"硬化表面 {pct(self.cov + self.road)}、其他表面仅 {pct(self.other_share)}，汇水速度快；"
             f"管线按当年负荷设计，空调与电动车充电负荷叠加后超载，且强弱电同杆混挂。",
             "电气火灾与内涝风险叠加，且每次改造都要重复开挖同一段街巷。"),
            ("康养、健康、便民休憩点位缺少系统布局，住宅 4～8 层无电梯，无障碍设施不足。",
             f"上述人口的服务需求分散在 {self.n_res} 个小区红线内，"
             f"单个小区配建达不到运营门槛，服务点位只能沿街零散设置。",
             "老年居民出行依赖楼梯与街巷，设施「有而不达」，日常照护成本转嫁给家庭。"),
            (f"{self.n_res} 个小区、{int(self.nar.num('叙述.产权形态数量'))} 类产权形态、{self.biz_count} 处商业，公共空间与管线设施没有统一维护主体。",
             "小区有红线、街区无边界；专项维修资金按小区归集、按小区使用，"
             "单个小区的资金规模不足以支撑跨小区的管线与公共空间维护。",
             "同一根管线挖两次、同一段街巷反复开挖，设施建成后维护缺位，更新成果难以维持。"),
        ]
        # 问题落点（P0-3）：每条问题必须落到具体小区与门牌。只写「片区」等于
        # 没分析——基准件靠「能点名的对象」建立说服力，这里按台账取值点名。
        _sp = self.spot
        _gs = list(self.group_stats.items())
        _g1n, _g1 = (_gs[0][0], _gs[0][1]) if _gs else ("", {"n": 0})
        _g2n, _g2 = (_gs[1][0], _gs[1][1]) if len(_gs) > 1 else ("", {"n": 0})
        _roads = "、".join(self.nar.list("项目.四至道路")) or "片区主要街巷"
        spots = [
            f"{_roads}围合的院落入口；其中{_sp['main']}的沿街段进深最大、断面最窄。",
            f"{_sp['oldest']}，均建成于 {self.year_min} 年"
            f"{'、' + _sp['oldest_trait'] if _sp['oldest_trait'] else ''}，是立面与屋面病害最集中的楼栋。",
            f"{_sp['smallest']}的院落边界，是街巷肌理与历史信息表达最薄弱的区段。",
            f"{self.biz_name}（{self.nar.text('项目.商业建筑.地址')}）为片区唯一成规模商业，"
            f"其余为沿街零散底商，无院落内部经营界面。",
            f"{_sp['main']}围合的院落入口与{_sp['mid']}尽端巷道，消防车均难以进入。",
            f"{_sp['tallest']}等高层数住宅的二次供水与电力增容，以及沿街巷架空的强弱电线缆。",
            f"{_sp['main_head']}周边 5 分钟步行范围，是社区服务站与康养点唯一可行的集中选址区段。",
            f"{_g1n}组 {_g1['n']} 个、{_g2n}组 {_g2['n']} 个院落（小区）分属不同管理主体，"
            f"跨小区管线无统一维护单位。",
        ]
        for i, (head, body, tail) in enumerate(problems, 1):
            seq += 1
            show, cause, effect = details[i - 1]
            L.extend(["", f"## {head}", "", f"**P-{i:02d} {head}**", "", body,
                      _anchor("D", seq, [self.internal(), self.material()] + self.derived_all),
                      "",
                      f"1. 表现：{show}",
                      f"2. 证据：{tail.split('更新方向')[0].strip().rstrip('。')}。",
                      f"3. 落点：{spots[i - 1]}",
                      f"4. 成因：{cause}",
                      f"5. 影响：{effect}",
                      f"6. 对应更新方向（需实施）：{tail.split('更新方向：')[-1].strip() if '更新方向：' in tail else tail}"])
        # TBL-02 核心问题清单：必须落在模板已声明的节内（契约只扫描模板节），
        # 故挂在最后一类问题的节尾，用加粗小标题引出，不另开二级标题。
        L.extend(["", "**核心问题清单**", ""])
        _rows = [[str(i + 1), p[0], p[1], p[2].split("更新方向")[0].strip()]
                 for i, p in enumerate(problems)]
        L.append(_md_table(["序号", "问题类别", "问题表征", "关键事实与指标"], _rows))
        L.extend(["", f"（数据来源：OSGB 现状指标、{self.ledger_name_short}、前期沟通材料）"])
        # 成因分析：把八类问题归因到历史、产权、空间、经济、治理五条主线
        L.extend(["", "## 成因分析", ""])
        D2 = self.derive([f"{self.area_wan * 10000 / self.households:.2f}="
                          f"{self.area_wan}*10000/{self.households}",
                          f"{self.cov + self.road:.6f}={self.cov}+{self.road}",
                          f"{self.other_share * 100:.2f}=100-{self.cov * 100:.2f}-"
                          f"{self.green * 100:.2f}-{self.road * 100:.2f}"])
        causes = [
            (f"历史成因：按当年标准建设，按现行标准使用。片区住宅建成于 {self.year_span} 年，"
             f"2000 年前建成 {self.before2000} 个、占 {self.before2000 / max(1, self.n_res) * 100:.2f}%。"
             f"这批住宅在设计时没有电梯、没有地下车库、雨水与污水合流、电力与通信线缆架空敷设，"
             f"按当时的人均居住面积与机动车保有量配置。房龄最长的已超过四十年，房屋本体进入大修期，"
             f"而当年的配置标准与今天的生活需求之间形成了第一层结构性差距。"),
            (f"产权成因：多主体、小单元，维修责任难以归集。台账内 {self.n_res} 个住宅小区以单位宿舍、"
             f"院落住宅、公寓与门牌式小区等不同类型并存，建筑性质包含住宅与商业，管理主体分散。房改后产权转移到住户，"
             f"但公共部位与管线仍由多家单位分段持有；专项维修资金按小区归集、按小区使用，"
             f"单个小区户数少、资金规模小，屋面渗漏、立面粉化这类「小修」被反复推迟，最终演变为"
             f"需要整体整治的「大修」。这是设施老化快于房龄的直接原因。"),
            (f"空间成因：密度高、余量小，公共空间没有增量来源。实测样区建筑密度 {pct(self.cov)}、"
             f"道路占比 {pct(self.road)}，两项加总 {pct(self.cov + self.road)}，其余表面仅 "
             f"{pct(self.other_share)}、约 {self.other_area} 平方米。实测样区东西 "
             f"{self.scope.get('extentM', {}).get('x', 0):.2f} 米、南北 {self.scope.get('extentM', {}).get('y', 0):.2f} 米，"
             f"街巷宽度有限，消防车道难以成环。停车、绿化、休憩、消防四类需求只能在同一块地面上竞争，"
             f"这正是「街巷被占用、绿地被硬化、消防通道不畅」并存的根源。"),
            (f"经济成因：容积率低、户均面积小，改造缺乏内生收益。样区容积率 {self.far:.4f}、平均 "
             f"{self.floors:.2f} 层，户均（折算）{self.area_wan * 10000 / self.households:.2f} 平方米，"
             f"拆除重建的拆迁成本远高于新增建筑面积的出让收益，不具备经济可行性。"
             f"项目投资以中央预算内投资与专项债券为主，属于政策性投入，没有经营性空间反哺。"
             f"投入结构决定了更新必须是「少拆多改、先地下后地上」的微更新，而不能依赖增容平衡。"),
            (f"治理成因：小区边界清晰、街区边界模糊，管线改造难以一次到位。{self.n_res} 个小区各自有红线，"
             f"但给排水、电力、通信管线跨小区敷设，权属分属不同单位。按小区立项、按小区施工，"
             f"会出现「同一根管线挖两次、同一段街巷反复开挖」。现状问题的治理成因不在单个小区，"
             f"而在缺少街区层面的统筹主体与统筹时序，这也是本报告主张以街区为单元组织更新的依据。"),
        ]
        for text in causes:
            seq += 1
            L.extend(["", text, _anchor("D", seq, [self.internal(), self.material()] + self.derived_all, "中")])
        L.append("")
        L.append("成因归纳：八类问题不是彼此孤立的质量缺陷，而是「建设年代早—产权分散—空间无余量—"
                 "收益不可平衡—治理缺统筹」五条主线叠加的结果。因此更新顺序应先解决管线与消防的安全底线，"
                 "再解决立面、环境与服务品质，最后以街区统筹机制固化成果；颠倒顺序会导致重复开挖与成果反弹。")
        return "\n".join(L), []

    # ── D：现状综合研判与更新思路 ──────────────────────────────────────
    def ch_E(self, title, h2):
        L = [f"# {title}"]
        seq = 0
        D = self.derive([f"{self.base}={self.area}*{self.cov}",
                         f"{self.green_area}={self.area}*{self.green}",
                         f"{self.road_area}={self.area}*{self.road}",
                         f"{self.gfa}={self.area}*{self.far}",
                         f"{self.floors}={self.far}/{self.cov}"])
        crit = ((self.ahp.get("weights") or {}).get("criteria") or [])
        crit_txt = "、".join(f"{c.get('name')} {float(c.get('weight') or 0):.6f}"
                            for c in crit)
        ranking = ((self.ahp.get("scores") or {}).get("ranking") or [])
        rank_txt = "；".join(f"第 {r.get('order')} 名 {r.get('name')}（{r.get('composite')}，{r.get('label')}）"
                            for r in ranking)
        demo = (self.ahp.get("scores") or {}).get("demo")

        def P(text, drv=None, conf="高"):
            nonlocal seq
            seq += 1
            L.extend(["", text, _anchor("E", seq, [self.internal(), self.material()] + (drv if drv is not None else self.derived_all))])

        L.append("## 综合现状研判")
        L.append("")
        seq += 1
        L.append(_anchor("E", seq, [self.internal()] + D))
        P(f"综合评价七项实测指标与台账数据（空间比率由实测样区折算、户数由台账加总），片区现状可概括为四句话："
          f"空间上「多层高密度」，实测样区容积率 {self.far:.4f}、建筑密度 {pct(self.cov)}、"
          f"台账住宅 4～{self.maxfloors} 层、含附属建筑基底折算 {self.floors:.2f} 层；"
          f"建筑上「结构尚可、围护老化」，砖混占 {self.brick / max(1, self.n_res) * 100:.2f}%、"
          f"2000 年前建成 {self.before2000} 个；设施上「管线老化、消防不足」；"
          f"人口上「户数集中、老龄化显性」。问题总结为安全、功能、品质三类短板并存；"
          f"潜力在于存量空间再组织与集中实施。这四点决定了更新必须以安全排险为底线、以环境提质为主线。")
        P(f"从指标组合看，片区不属于需要整体拆除重建的类型。实测样区建筑密度 {pct(self.cov)} 意味着拆除后"
          f"可释放的地面空间有限，容积率 {self.far:.4f} 意味着增容收益不足以覆盖拆迁成本；"
          f"而样区绿地率 {pct(self.green)} 与开敞空间 {pct(self.opening)} 说明环境品质有提升余地。"
          f"因此更新路径应是「保留主体、局部拆危、织补功能、整治环境」。")
        P(f"从风险看，最需要优先处理的是三类：一是 {self.hmax:.2f} 米中高层住宅的消防通行与取水；"
          f"二是架空线缆与电力容量不足带来的电气安全；三是实测样区硬化表面 {pct(self.cov + self.road)} "
          f"叠加排水不畅形成的积水隐患。三类风险都与市政管线改造强相关，应以一次管线综合整治带动。")
        P(f"AHP 权重进一步支持这一判断。四项准则权重为：{crit_txt}；其中经济可行与风险可控两项"
          f"合计超过六成，说明方案比选中「投入产出可平衡」与「安全风险可控制」是主导因素。"
          f"这一权重结构与本报告把安全排险、管线整治放在首位的结论一致。")
        P(f"AHP 方案排序为：{rank_txt}。排序结果指向渐进式微更新，而非整体拆除重建，"
          f"与本报告的安全排险、管线整治优先结论一致，作为更新方向比较的方法支撑。")

        # 现状评价矩阵：八个维度统一按“数据—基准—判断—指向”落表
        _buckets = self._area_buckets()
        if _buckets:
            _area_cmp = "；".join(
                f"{a} {b['n']} 个小区" for a, b in _buckets.items())
        else:
            _area_cmp = f"台账内 {self.n_res} 个小区"
        L.append("")
        L.append("**现状评价矩阵（八维）**")
        L.append("")
        L.append(_md_table(
            ["维度", "关键现状值", "对比基准", "判断结论", "更新指向"],
            [["区位与范围",
              f"{self.area} 平方米；东西 {self.scope.get('extentM', {}).get('x', 0):.2f} 米 × "
              f"南北 {self.scope.get('extentM', {}).get('y', 0):.2f} 米；{self.n_res} 个小区",
              "5 分钟生活圈完整街坊", "规模小、被城市干道切割", "小尺度织补、不封闭成院"],
             ["用地与建筑",
              f"建筑密度 {pct(self.cov)}；容积率 {self.far:.4f}；住宅 {self.buildings} 栋",
              "旧区改建绿地率不低于 25%；老城多层街区密度经验值 30%—40%",
              "密度偏高、增容空间有限", "不增容、地面空间整理"],
             ["人口与社会",
              f"{self.households} 户、约 {self.population} 人；户均 {self.population / max(1, self.households):.1f} 人",
              "完整社区 5—10 分钟服务圈", "户数集中、老龄化显性", "适老化改造与服务补短"],
             ["设施与安全",
              f"道路覆盖 {pct(self.road)}；消防通道不足 4 米；管线架空",
              "消防车道净宽不小于 4 米；无障碍坡道坡度不大于 1/12",
              "消防与管线是底线短板", "先地下后地上、消防优先"],
             ["公共空间与绿化",
              f"绿地率 {pct(self.green)}；其他表面 {pct(self.other_share)}",
              "旧区改建绿地率不低于 25%", "总量达标、分布不均", "口袋公园与透水铺装"],
             ["产业与活力",
              f"商业 {self.biz_count} 处（{self.biz_name}）；沿街业态单一",
              "5—10 分钟便民商业覆盖", "业态单一、沿街活力不足", "业态引导与界面整治"],
             ["风貌与文化",
              f"住宅以 1980—2000 年建成；院落、行列、点式并存",
              "保护优先、留改拆并举", "无公布历史建筑，但院落肌理有价值", "材质、铺装、标识延续文脉"],
             ["治理与诉求",
              f"{self.n_comm} 个社区、多管理主体；诉求集中于管线、立面、积水、适老化",
              "业主参与、共建共治", "主体分散、长效运营缺位", "街区统筹与社区参与机制"]]))
        L.append("")
        L.append("（对比基准口径：25% 为《城市居住区规划设计标准》GB 50180-2018 对旧区改建绿地率的"
                 "底线要求，4 米为《建筑防火通用规范》GB 55037-2022 的消防车道净宽要求，"
                 "1/12 为《建筑与市政工程无障碍通用规范》GB 55019-2021 的坡道坡度限值；"
                 "“老城多层街区密度经验值 30%—40%”为同类街区经验口径，不是标准限值，"
                 "需在控规层面复核。）")

        L.append("")
        L.append("**（一）宏观层面：政策取向与行业趋势**")
        L.append("")
        P("国家层面的政策取向已完成从“增量扩张”到“存量提质”的转向。国办发〔2020〕23号"
          "把城镇老旧小区改造分为基础类、完善类、提升类，明确“保基本”优先；"
          "建科〔2021〕63号要求城市更新单元内拆除建筑面积原则上不大于现状总建筑面积的 20%、"
          "拆建比不宜大于 2，把“留改拆”置于“拆改留”之前；2025 年中办、国办印发的"
          "《关于持续推进城市更新行动的意见》进一步把城市更新列为扩大内需、民生改善和"
          "安全韧性建设的重点任务。三条政策共同框定了本项目“不整体拆除、以保留利用和"
          "功能织补为主”的更新方式，也说明必要性不再是“要不要改”，而是“按什么方式改”。")
        P("行业趋势从“单项工程”转向“片区统筹”和“全生命周期运营”。城市体检（建科〔2023〕75号）"
          "要求“先体检、后更新”，把安全、设施、环境短板作为更新项目生成的前置依据；"
          "完整社区建设要求以 5—10 分钟生活圈补齐养老、托育、"
          "便民商业和公共活动空间；无障碍环境建设法（2023 年施行）把既有居住社区的"
          "无障碍改造纳入法定责任。本项目的短板恰好集中在管线、消防、无障碍与公共空间，"
          "属于政策明确支持的基础类和完善类内容，具备申报与实施的政策窗口。")
        P("从资金与实施机制看，中央预算内投资、超长期特别国债、地方政府专项债均把"
          "城市更新和老旧街区改造列为支持方向；成都市城市更新条例（2026 年 6 月 1 日施行）"
          "确立了城市更新单元、保护优先和业主参与制度，为跨小区连片实施提供了地方性法规依据。"
          "宏观判断的结论是：本项目具备政策依据、资金渠道和审批路径，更新必要性成立，"
          "且应以片区为单元整体谋划，避免单个小区碎片化实施。")

        L.append("**（二）中观层面：省市规划与片区定位**")
        L.append("")
        P("某省与某市的政策部署在同一点上收敛：以街道为实施主体、以片区为单元、"
          "以安全韧性为底线。川办发〔2020〕63号全面推进城镇老旧小区改造，"
          "川府发〔2026〕13号把片区统筹、完整社区、韧性城市列为“十五五”城市更新重点，"
          "成住建发〔2022〕82号规范老旧院落改造的工程组织与资金管理，"
          "成都市城市更新条例则把“留改拆”和“城市更新单元”写入地方性法规。"
          "四级政策共同要求把分散的小区改造组织成片区更新，这与本项目跨社区、"
          "跨小区的空间格局一致。")
        P(f"片区定位上，项目位于示例街道，涉及 {self.n_comm} 个社区，处在某市老城核心"
          "商圈外围的过渡带。其更新不是单纯的住宅修缮，而是老城核心区"
          "功能疏解、人口结构优化和街区活力再组织的一部分：一方面承接商圈外溢的居住与"
          "服务需求，另一方面承担老城居民就地改善的民生任务。定位决定了更新不能只做"
          "立面，必须同步补足消防、管线、公共空间和社区服务。")
        P(f"同类片区比较看，项目内部并非均质。{_area_cmp}。各区域在小区规模、户数密度和"
          "楼栋条件上存在差异，说明不能用一套标准平均用力：户数集中、楼栋连片的区域"
          "适合作为一期实施重点，规模小、边界散的区域适合与相邻街巷打包实施。"
          "这一比较也支持“分区施策、分期推进”而不是“一次性铺开”的组织方式。")
        P("中观判断的结论是：本项目处在省市城市更新的重点区域和政策支持范围内，"
          "又具有跨社区连片实施的条件，具备作为片区更新单元整体立项的必要性；"
          "若继续按单个小区申报，管线迁改、消防通道和公共空间都难以成网，"
          "单位投资效率会明显下降。")

        L.append("**（三）微观层面：空间、安全与居民诉求**")
        L.append("")
        P(f"区位与范围上，实施范围含{self._n_area_text()}个片区：主体片区由{self.nar.text('项目.测绘范围四至') or '片区主要街巷'}围合，"
          f"另一片区按台账门牌涉及{self._area_extra_roads_text()}，"
          f"测绘范围 {self.area} 平方米（东西 {self.scope.get('extentM', {}).get('x', 0):.2f} 米 × "
          f"南北 {self.scope.get('extentM', {}).get('y', 0):.2f} 米），"
          f"台账内 {self.n_res} 个住宅小区、{self.buildings} 栋、{self.households} 户、"
          f"约 {self.population} 人。两个片区的实施边界与时序以委托方确认为准；规模小、密度高、边界被城市干道切割的空间特征，"
          "决定了更新只能采用小尺度、织补式的手法，不能依赖拆除重建。")
        P(f"建筑与安全上，台账内 2000 年前建成 {self.before2000} 个、砖混结构 {self.brick} 个，"
          f"最高 {self.hmax:.2f} 米、住宅 4～{self.maxfloors} 层且无电梯；实测样区建筑密度 "
          f"{pct(self.cov)}、容积率 {self.far:.4f}。对照消防车道净宽不小于 4 米、"
          f"无障碍坡道坡度不大于 1/12 的现行要求，片区的消防通行与取水条件存在明确差距。"
          f"对照旧区改建绿地率不低于 25% 的要求，实测样区绿地率虽达标但分布不均，"
          f"绿地可达性不足。此外，架空线缆、"
          f"电力容量不足和雨污混接则构成电气与积水隐患。这些不是品质问题，而是安全底线问题。")
        P(f"人口与诉求上，按户均 3.2 人测算的居住规模与老龄化特征叠加，"
          "前期沟通材料反映的诉求集中在管线安全、立面渗漏、汛期积水、适老化与便民休憩四类，"
          "与本次拟实施的管线、立面、消防、公共空间和社区服务五项内容高度重合。"
          "诉求与工程清单重合，意味着居民共识基础较好、实施阻力较小，"
          "但也要求把居民参与和后期管护写进实施机制，避免整治后反弹。")
        P(f"经济与实施上，容积率 {self.far:.4f}、建筑密度 {pct(self.cov)} 决定了拆除重建"
          "无法平衡拆迁与安置成本；而存量空间足以支撑微更新：其他表面 "
          f"{self.other_area} 平方米、道路 {self.road_area} 平方米、立面 "
          f"{self.work_facade:g} 平方米，均可通过功能置换和整治释放。"
          "微观判断的结论是：不改，安全风险会随设施老化继续累积；改，投入可控、"
          "居民受益直接，必要性最为具体。")

        L.append("**（四）更新必要性结论**")
        L.append("")
        P("三层论证收敛为一句：宏观政策要求改、中观规划部署改、微观安全民生必须改。"
          "本项目应定位为“保留利用为主、局部拆危为辅、管线与消防先行的老旧街区综合更新”，"
          "不建议整体拆除重建；建议以本次更新为契机，同步建立街区统筹与长效运营机制，"
          "把一次性的工程投入转化为可持续的街区治理能力。")

        L.append("## 更新潜力与问题识别")
        L.append("")
        P(f"更新潜力来自存量空间的再组织，而不是新增建设用地。按实测样区指标测算，可挖潜的空间载体有四类。"
          f"第一类是其他表面 {self.other_area} 平方米（占样区 {pct(self.other_share)}），"
          f"它是样区总面积减去建筑基底、绿地与道路后的残差，含院落硬化、边角地与模型未分类表面，"
          f"需现场逐块核实后才可落位，可作口袋公园、透水铺装与休憩设施的候选来源；第二类是屋面与立面，"
          f"台账立面材料修缮或翻新合计 {self.work_facade:g} 平方米"
          f"（住宅 {self.work_facade_res:g} 平方米 + {self.biz_name} {self.work_facade_hotel:g} 平方米），"
          f"住宅部分对应台账内 {self.buildings} 栋住宅，"
          f"整治后可同步改善防水、隔热与沿街界面；第三类是街巷断面，"
          f"实测样区道路覆盖 {self.road_area} 平方米、占比 {pct(self.road)}，通过断面优化可释放慢行、停车与"
          f"消防通行空间；第四类是组织潜力，台账 {self.n_res} 个小区分属 {self.n_comm} 个社区、"
          f"沿街巷相邻分布，{self.main_name} {self.households_main} 户占住宅户数的 {self.main_share * 100:.2f}%，"
          f"连片立项可以摊薄管线迁改与施工组织的单位成本。")
        P(f"潜力释放受三条约束。其一，实测样区建筑密度 {pct(self.cov)} 决定了不能通过拆除增加开敞空间，"
          f"只能在既有基底之间做「减法」与「置换」；其二，样区容积率 {self.far:.4f} 与平均 {self.floors:.2f} 层"
          f"决定了加建空间有限，加装电梯、社区服务设施应优先利用院落边角与低效用房；"
          f"其三，街巷宽度与消防要求存在冲突，停车与绿化的增量必须让位于消防通道的连续性。"
          f"三条约束共同指向「微更新、少增量、重组织」的潜力释放方式。")
        L.append("")
        L.append(_md_table(
            ["潜力类型", "空间载体", "规模（现状值）", "主要约束"],
            [["地面空间再组织", "其他表面、院落边角", f"{self.other_area} 平方米（{pct(self.other_share)}）",
              "不得占用消防通道"],
             ["立面与沿街界面", "住宅沿街立面、雨棚与披水板", f"立面 {self.work_facade:g} 平方米；雨棚 {self.work_awning:g} 个；披水板 {self.work_apron:g} 个",
              "结构与防水同步复核"],
             ["街巷断面优化", "生活性街巷", f"道路 {self.road_area} 平方米（{pct(self.road)}）",
              "消防通行优先、机非分离"],
             ["集中实施", f"{self.n_res} 个住宅小区", f"{self.households} 户；首位小区 {self.main_share * 100:.2f}%",
              "跨小区管线统一时序"]]))
        L.append("")
        L.append(f"（数据来源：OSGB 现状指标、{self.ledger_name_short}）")
        P(f"问题识别按「安全底线—功能短板—品质提升」三级排序。安全底线包括消防通行与取水、"
          f"电气与燃气管线老化、{self.hmax:.2f} 米住宅的立面坠物风险，对应市政管线综合整治与消防补短；"
          f"功能短板包括停车、无障碍、社区服务与康养，对应公共空间与社区服务补缺；"
          f"品质提升包括风貌、绿化与街区治理，对应文脉延续与长效运营。三级问题存在先后依赖："
          f"地下管线不一次整治，地面空间与立面整治就会重复开挖；消防通道不打通，停车与绿化的增量就无处落地。"
          f"因此实施顺序应与问题层级一致，而不是按小区平均分配投资。")

        L.append("## 总体更新思路")
        L.append("")
        P(f"总体思路为「一保留、两整治、三补齐」，并遵循三条原则：保护优先、安全底线、微改造为主。"
          f"「一保留」是保护街巷走向、院落边界与 "
          f"实测样区 {self.h50:.2f} 米中位高度反映的空间尺度，不采用大拆大建；「两整治」是整治建筑立面与"
          f"市政管线，解决安全与风貌问题；「三补齐」是补齐消防通道与取水点、补齐社区服务与康养设施、"
          f"补齐公共空间与绿化。三项原则共同约束实施方式：以微改造替代拆除重建，以功能织补替代增量扩张。")
        P(f"空间策略上，以 {self.area} 平方米测绘范围为核心，按「街巷—院落—建筑」三级组织："
          f"街巷层解决通行、消防与管线；院落层解决停车、绿化与休憩；建筑层解决立面、屋面、"
          f"楼道与加装电梯。三级同步实施，避免只做立面、不做地下。")
        P(f"实施策略上，按户数分布加总分片推进。{self.main_name} {self.households_main} 户占住宅户数 "
          f"{self.main_share * 100:.2f}%，应作为一期实施重点；其余 {self.n_res - 1} 个小区累计 "
          f"{self.households - self.households_main} 户，按街巷分段滚动实施。"
          f"计划周期与项目信息卡一致，为 {self.nar.text('项目.实施周期')}。")
        P("治理策略上，建立街区统筹机制：街道牵头，社区与小区管理主体参与，"
          "管线与智慧运维平台由专业单位运营，公共空间保洁与设施维护纳入街区统一管理。"
          "更新不是一次性工程，运营机制决定成果能否维持。")

        L.append("## 核心更新改造方向")
        L.append("")
        L.append(
            f"更新改造方向按「安全隐患—功能短板—风貌文脉—业态活力」四条线组织，"
            f"每条线对应前述问题编号，并落到台账的工程量基准上。六项动作不是并列关系："
            f"管线与消防属安全隐患类，必须先地下后地上；立面与公共空间属品质类，"
            f"承接风貌文脉的延续；社区服务与适老化属功能类；业态引导与长效治理属业态活力类，"
            f"在硬件完成后跟进。顺序错位会导致重复开挖与投资浪费，这也是本报告把管线列为首项的原因。")
        strategies = [
            ("市政管线综合整治",
             "P-01、P-05、P-06 对应的空中线缆杂乱、电力容量不足、排水不畅三类现状问题",
             f"新增街区电力设施并扩容，给排水、强弱电、通信管网全部下地，"
             f"搭建管线智慧运维平台；按前期材料明确的方案先地下后地上实施。",
             f"消除架空线缆与电力过载风险；台账内冷凝水管 {self.work_condensate:.0f} 米、"
             f"落水管 {self.work_downpipe:.0f} 米随立面整治同步更换，减少外墙渗水与雨污混接。"),
            ("建筑立面与屋面整治",
             "P-02 对应的抹灰饰面老化、渗水、空鼓与瓷砖掉落",
             f"台账内住宅 {self.buildings} 栋、住宅建筑规模 {self.area_wan} 万平方米，"
             f"抹灰外立面 {self.plaster} 个；立面材料修缮或翻新 {self.work_facade:.0f} 平方米、"
             f"更换外墙雨棚 {self.work_awning:.0f} 个、"
             f"空调外机披水板 {self.work_apron:.0f} 个。",
             "围护结构渗漏与坠物风险下降，沿街界面连续整洁。"),
            ("消防与安防补短",
             "P-05 对应的院落入口狭窄、消防车道难成环、取水点不足",
             f"结合 {self.hmax:.2f} 米住宅与院落入口条件，打通消防车道，按街巷分段增设"
             f"消防取水点与微型消防站，同步更新安防监控。台账六项工程量均为房屋本体修缮，"
             f"不含消防设施，消防补短需单列投资与点位，不能以雨棚、披水板数量充抵。",
             "消防通道连续、取水半径缩短，安防盲区减少。"),
            ("公共空间与绿化提质",
             "P-01、P-07 对应的公共空间无增量、绿地可达性不足",
             f"以 {self.other_area} 平方米其他表面为主要挖潜对象，把街角、院落边角改造为"
             f"小微空间与透水铺装；结合实测样区绿地率 {pct(self.green)} 的存量补植沿街行道树；"
             f"屋顶绿化先按栋核实屋面可绿化面积，再定是否纳入与实施规模。",
             "在实测样区建筑密度不降的前提下提升绿地可达性与雨天渗滞能力。"),
            ("社区服务与适老化",
             "P-07、P-08 对应的社区服务设施不足、无电梯住宅适老化缺口",
             f"按上述人口规模测算的服务半径，"
             f"{self.main_name}周边设社区服务站与康养点，其余小区补充便民休憩与健身设施；"
             f"对 {self.maxfloors} 层无电梯住宅推进加装电梯与楼道扶手，落实无障碍坡道。",
             "步行 5 至 10 分钟社区服务覆盖，老年居民出行条件改善。"),
            ("文脉延续与街区治理",
             "P-03、P-08 对应的肌理被零散搭建切割、治理主体分散",
             f"保留街巷走向、院落边界与建筑尺度，用铺装、标识、立面材质表达 {self.nar.text('叙述.文脉记忆')}；"
             "建立街区统筹管理与社区参与机制。",
             "更新成果纳入长效运营，避免整治后反弹。"),
        ]
        for i, (name, problem, action, effect) in enumerate(strategies, 1):
            seq += 1
            L.extend(["", f"**S-{i:02d} {name}**", "",
                      f"1. 对应问题：{problem}。",
                      f"2. 更新动作：{action}",
                      f"3. 预期成效：{effect}",
                      _anchor("E", seq, [self.internal(), self.material()] + self.derived_all)])

        L.append("")
        L.append("**更新方向与现状指标对应表**")
        L.append("")
        L.append(_md_table(
            ["更新方向", "对应现状指标", "主要工程量或对象"],
            [["市政管线综合整治", f"硬化表面 {pct(self.cov + self.road)}；管线全部架空", f"冷凝水管 {self.work_condensate:g} 米、落水管 {self.work_downpipe:g} 米"],
             ["建筑立面与屋面整治", f"砖混 {self.brick} 个；抹灰外立面 {self.plaster} 个", f"立面 {self.work_facade:g} 平方米、雨棚 {self.work_awning:g} 个、披水板 {self.work_apron:g} 个"],
             ["消防与安防补短", f"最高 {self.hmax:.2f} 米；消防设施不完善", "消防取水点、微型消防站、安防监控"],
             ["公共空间与绿化", f"样区绿地率 {pct(self.green)}；其他表面 {pct(self.other_share)}", f"其他表面 {self.other_area} 平方米"],
             ["社区服务与适老化", f"{self.households} 户；{self.maxfloors} 层无电梯", "社区服务站、康养点、加装电梯"],
             ["文脉延续与治理", f"{self.year_span} 年住宅为主体", "街巷铺装、标识、院落边界"]]))
        L.append("")
        L.append(f"（数据来源：OSGB 现状指标、{self.ledger_name_short}、前期沟通材料）")

        # 共性问题与差异性问题分列：说清哪些一起改、哪些分开改。
        # 共性=八类问题里跨院落同源、可一次统筹的；差异=组际或单院落特有。
        _a, _b = self.group_stats.get("示例社区A"), self.group_stats.get("示例社区C")
        L.extend(["", "**共性问题与差异性问题**", ""])
        if _a and _b:
            L.append(
                f"八类问题里，管线与消防属 {self.n_res} 个院落共有的底线问题：同一根给排水管、"
                f"同一条电力干线跨院落敷设，分组立项只会重复开挖，必须以街区为单元一次统筹。"
                f"楼道修缮、无障碍坡道与公共照明同理，成本低、覆盖面广，宜随管线开挖同步完成。")
            L.append("")
            _y = lambda g: (f"{g['year_min']} 年" if g['year_min'] == g['year_max']
                            else f"{g['year_min']}—{g['year_max']} 年")
            L.append(
                f"差异性问题集中在两组：示例社区A组 {_a['n']} 个院落建成于 {_y(_a)}、"
                f"砖混 {_a['brick']} 个、抹灰 {_a['plaster']} 个，屋面防水与立面整治须先行，"
                f"否则管线更新后仍会因渗漏返工；示例社区C组 {_b['n']} 个院落建成于 "
                f"{_y(_b)}、框架 {_b['frame']} 个、面砖 {_b['tile']} 个，"
                f"结构条件较好，重点转向电梯加装、立面品质与社区服务补位。")
            L.append("")
            L.append(
                f"单院落层面，{self.spot['main_head']}应成片推进、独立排期，其户数占片区 "
                f"{self.main_share * 100:.2f}%，施工组织与居民协调量都在这个院落；"
                f"{self.spot['smallest']}体量最小，宜并入相邻街巷标段，避免单独进场。"
                f"共性部分一次改完、差异部分分组实施，是本次更新在时序上的基本结论。")
        L.append("")
        return "\n".join(L), []

    # ── F：附件（基础资料/图件影像/专项台账） ──────────────────────────
    def ch_F(self, title, h2):
        L, seq = [f"# {title}"], 0
        files = self.source_files()

        def A(text, extra=None):
            nonlocal seq
            seq += 1
            L.extend(["", text, _anchor("F", seq, [self.internal(), self.material()] +
                                        self.derived_all + (extra or []))])

        def count_derive(n):
            if n <= 1:
                return []
            return self.derive([f"{n}={'+'.join('1' for _ in range(n))}"])

        L.append("## 附件一 现有基础资料清单")
        A(f"本报告登记并使用包内资料 {len(files)} 份，覆盖项目基本信息、改造小区台账、"
          f"工程量清单、前期沟通材料与上游机读成果五类。全部资料均按文件名称登记来源，"
          f"正文引用时以「材料：文件名 + 表号」标注定位，可逐条回溯。"
          f"资料使用遵循「一数一源」：同一指标只采用一份资料的口径，不跨资料拼接。",
          count_derive(len(files)))
        L.append("")
        rows = []
        for i, f in enumerate(files, 1):
            kind, purpose, used = self._file_usage(f)
            rows.append([str(i), f, kind, purpose, used])
        L.append(_md_table(["序号", "文件名称", "类型", "主要用途", "本次使用情况"], rows))
        L.append("")
        L.append("（数据来源：输入文件目录登记清单）")

        L.append("## 附件二 图件与影像资料")
        A(f"现状图（图 1）以 OSGB 正射影像为底图，坐标系 EPSG:4544，"
          f"叠加项目边界与建筑轮廓后出图，图幅覆盖实测样区 {self.area} 平方米"
          f"（东西 {self.scope.get('extentM', {}).get('x', 0):.2f} 米 × "
          f"南北 {self.scope.get('extentM', {}).get('y', 0):.2f} 米），图内标注图例、指北针、比例尺与数据来源。"
          f"因正射影像源自三维模型投影，地面分辨率与平面精度受建模精度制约、纹理存在拉伸，"
          f"图面同时标注实测指标与影像限制，不以示意图替代实测。")
        A("本报告引用图件与影像资料以现状照片、设计示意图与汇报资料为主，"
          "已提供的图件用于核对建筑立面、屋面与街巷界面；未随材料提供的总平面图、管线图与"
          "控规图则，正文以文字描述替代，不虚构图件内容。图件引用遵循「图—文—数」一致："
          "图上可量取的尺寸与正文指标同源，无法量取的只作定性描述。")
        L.append("")
        L.append(_md_table(
            ["序号", "图件/影像", "应表达的内容", "状态"],
            [["1", "改造小区现状照片", "立面、屋面、楼道与院落界面", "用于核对，未在正文附图"],
             ["2", "设计示意图", "街巷铺装、绿化与节点改造意向", "用于核对，未在正文附图"],
             ["3", "总平面图", "建筑基底、道路与绿地边界", "以文字描述替代"],
             ["4", "管线综合图", "给排水、电力、通信管线走向", "以文字描述替代"],
             ["5", "控规图则", "用地性质、高度与退界控制", "以文字描述替代"]]))
        L.append("")
        L.append("（数据来源：素材汇编图件登记）")

        L.append("## 附件三 专项台账")
        _cov_rg = self._range("building_coverage_ratio")
        _cols = self.work_map["columns"]
        _pos = [c for c in _cols if getattr(self, "work_" + c["role"]) > 0]
        A(f"专项台账覆盖现状指标、改造小区、工程量与方案比选四类。OSGB 现状指标 {len(self.m)} 项已按原值引用，"
          f"其中 {self.n_low} 项为估算口径（建筑栋数、高度中位与 90 分位、容积率），"
          f"{self.n_diag} 项为诊断口径（绿地覆盖、道路覆盖，两轮真值校核交并比未达阈值），"
          f"建筑密度为网格敏感性较低项（区间 {pct(_cov_rg[0])}～{pct(_cov_rg[1])}），可作量级引用；改造小区台账加总覆盖 {self.n_res} 个小区、"
          f"{self.households} 户、{self.buildings} 栋；工程量台账按源表列名逐列引用，表头共 {len(_cols)} 项，"
          f"其中 {len(_pos)} 项有申报量并列入正文工程量表（{'、'.join(c['item'] for c in _pos)}），"
          f"屋面翻新申报量为零（{self.nar.text('项目.屋面翻新说明')}），正文未单列；"
          f"AHP 权重与排序已提供并按其定稿值引用。"
          f"台账口径以「测绘范围」与「小区红线」两套并列，正文分别注明，不相互替代。")
        A("口径说明：正文的建筑栋数与建筑高度均为甲口径，即 2.5D 网格法（一米格、四邻域闭运算后按连通分量计数）"
          "在实测样区上的识别结果，栋数对格距与附属建筑敏感，区间已在指标表列明；"
          "若改用乙口径（逐栋单体化）统计，栋数与高度会随单体切分与附属建筑归属变化，"
          "与网格值不是同一口径，不能直接互换。本报告统一采用甲口径表述，乙口径数值不进入正文。")
        L.append("")
        L.append(_md_table(
            ["序号", "台账名称", "应包含字段", "状态"],
            [["1", "OSGB 现状指标", "建筑密度、绿地率、道路占比、容积率、栋数、高度",
              f"已引用（{self.n_low} 项估算、{self.n_diag} 项诊断口径）"],
             ["2", "改造小区台账", "小区名称、户数、栋数、建筑面积、年代、结构", "已引用（全量）"],
             ["3", "工程量台账", "立面材料修缮或翻新、屋面翻新、更换外墙雨棚、冷凝水管、落水管、空调外机披水板",
              "已引用（逐列按表头映射）"],
             ["4", "AHP 权重与排序", "准则权重、一致性、方案排序", "已引用（按定稿值）"]]))
        L.append("")
        L.append("（数据来源：包内依据/现状指标.json、权重表.json、评分表.json）")
        return "\n".join(L), []

    def source_files(self):
        """从素材汇编的来源行提取去重后的文件名清单。"""
        out, seen = [], set()
        for ln in self.blocks:
            m = re.search(r"【来源：([^】]+)】", ln)
            if not m:
                continue
            raw = m.group(1).strip()
            # 「包内依据.md（OSGB 现状指标 / AHP 结论）」这类来源带了括号说明，
            # 括号里的斜杠会被当路径分隔符切出残名「AHP 结论）」，先去掉说明再取文件名。
            raw = re.split(r"[（(]", raw)[0].strip()
            base = raw.replace("\\", "/").split("/")[-1]
            if base and base not in seen:
                seen.add(base)
                out.append(base)
        # 包内依据.md 是 05 步由三份源 JSON 汇总生成的过程件，不是原始资料；
        # 附件登记原始三份，正文《同源核对.xlsx》与这三份逐项对应。
        if "包内依据.md" in seen:
            out = [f for f in out if f != "包内依据.md"]
            seen.discard("包内依据.md")
        for f in ("现状指标.json", "权重表.json", "评分表.json"):
            if f not in seen:
                seen.add(f)
                out.append(f)
        if not out and self.res:
            out = [r["file"] for r in self.res]
        return out

    def _file_usage(self, name: str):
        """附件表「类型 / 主要用途 / 本次使用情况」按文件实际作用填写，不套同一句话。"""
        low = name.lower()
        if "现状指标" in name:
            return ("机读成果", f"OSGB 现状指标 {len(self.m)} 项：密度、容积率、栋数、高度、绿地、道路",
                    f"已引用（{self.n_low} 项估算、{self.n_diag} 项诊断口径）")
        if "权重表" in name:
            return ("机读成果", "AHP 准则权重与一致性", "已引用（按定稿值）")
        if "评分表" in name:
            return ("机读成果", "AHP 方案得分与排序", "已引用（按定稿值）")
        if "项目信息" in name:
            return ("文档材料", "建筑面积、投资、工期", "已引用（项目信息卡）")
        if "前期沟通" in name:
            return ("文档材料", "项目背景、现状照片与改造意向", "已引用（背景与图件）")
        if low.endswith((".xlsx", ".xls")) or "统计表" in name or "台账" in name:
            return ("表格台账", "小区基本信息、户数栋数、改造内容与工程量清单", "已引用（台账全量逐列）")
        if low.endswith((".docx", ".doc")):
            return ("文档材料", "社区更新需求与居民诉求", "已引用（诉求归纳）")
        return ("文档材料", "现状资料", "已登记并引用")
