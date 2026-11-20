# -*- coding: utf-8 -*-
"""工程量闸门：报告里的每个台账工程量，都必须能在源表对应列找到。

R1 挑刺的 P0-1 是这么错的：台账 Sheet1-2 六列按列序套成
「楼道 765 个 / 管线 330m / 道路 294m / 其他工程 834 个」，五行错四行，
834 还被当消防设施引用。所以这里不认列序，只认显式映射：
  * 源列名（台账原文）→ 报告项名 → 单位，登记在 规则/工程量列映射.json；
  * 合计必须等于逐小区明细相加（改了正文没改明细会被抓）；
  * 正文里「报告项名 + 数字」必须等于该列合计；
  * 映射表里没登记的老项名（楼道/管线改造/其他工程/其他设施/道路整治）
    一旦后面跟着数量，直接判违规——改名必须回映射表登记。
"""
from __future__ import annotations

import json
import re

NUM = r"([0-9][0-9,]*\.?[0-9]*)"
LEGACY = ("楼道", "管线改造", "其他工程", "其他设施", "道路整治")
# 项名与数字之间不许跨句读、括号、表格竖线：PDF 抽文本会把相邻单元格/相邻句子糊在一起，
# 12 字窗口原样放开时，「……空调外机披水板已引用（逐列按表头映射）4 AHP……」的 4 会被当工程量。
GAP = r"[^0-9。；，、：！？（）()|｜]{0,12}"
# 旧项名要判违规，还必须是「数量」而不只是「碰巧后面有数字」：
# PDF 抽文本会把表格行糊成「……楼道与院落界面已引用2 设计示意图」，那个 2 是序号。
# 所以要求数字后紧跟计量单位。
UNIT_AFTER = r"\s*(?:个|米|m|平方米|㎡|处|盏|套|台|项|根|块|组|座)"


def load_snapshot(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def scan(text: str, snap: dict):
    """返回 (违规清单, 核对明细)。"""
    bad, rows = [], []
    flat = re.sub(r"\s+", "", text or "")
    for col in snap.get("columns", []):
        item = col.get("item") or ""
        role = col.get("role") or ""
        total = round(float(col.get("total") or 0), 2)
        subtotal = round(sum(float(r.get(role) or 0) for r in snap.get("rows", [])), 2)
        rows.append({"sourceName": col.get("sourceName"), "item": item,
                     "unit": col.get("unit"), "total": total,
                     "subtotal": subtotal, "balanced": abs(subtotal - total) < 0.01})
        if abs(subtotal - total) >= 0.01:
            bad.append(f"{item}：合计 {total:g} 与逐小区加总 {subtotal:g} 不一致")
        if total <= 0:
            # 台账该列为 0（如屋面翻新）：正文不得出现「项名 0 单位」的凑数写法
            if re.search(re.escape(item) + GAP + r"0(?!\d)", flat):
                bad.append(f"{item}：台账合计为 0，正文不应以「0」列示")
            continue
        if item not in flat:
            bad.append(f"{item}：正文未出现（台账合计 {total:g}）")
        for m in re.finditer(re.escape(item) + GAP + NUM, flat):
            got = float(m.group(1).replace(",", ""))
            if abs(got - total) > 0.01:
                bad.append(f"{item}：正文写 {got:g}，台账合计 {total:g}")
    for word in LEGACY:
        for m in re.finditer(re.escape(word) + GAP + NUM + UNIT_AFTER, flat):
            bad.append(f"未登记工程项名「{word} {m.group(1)}」，"
                       f"须回 规则/工程量列映射.json 登记后再引用")
    return bad, rows
