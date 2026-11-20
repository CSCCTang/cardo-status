#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_num_match.py -- 数值回溯的匹配口径（build_trace 与 check_trace 共用）。

**为什么要有这个模块**：回溯表生成（build_trace）与回溯校验（check_trace）
必须用**同一套**匹配口径，否则会出现「回溯表登记了、校验却不认」的假失败。

**为什么要归一**：事实库存的是 Python 浮点与符号单位
（`744.0` / `6.268万㎡`），正文写的是整数与中文单位名（`744` / `6.268万平方米`），
字符串直接比必然不等——实测 73 条事实只命中 1 条，回溯表近乎为空。

**边界**：只做**书写形式**归一（数值表示 + 单位别名），**不做单位换算**
（换算会引入正文里没有的数，等于凭空造数）。

约定：本模块是纯函数，不读业务文件、不写任何文件。
"""

from __future__ import annotations

import re

# 单位别名：同一物理量的不同书写形式。
# 事实库用符号、正文用中文单位名（或反过来）都在这张表里。
UNIT_ALIAS = {
    "㎡": "平方米", "m²": "平方米", "平米": "平方米",
    "万㎡": "万平方米", "万平方米": "万平方米",
    "m": "米", "米": "米",
    "km": "公里", "千米": "公里", "公里": "公里",
    "ha": "公顷", "公顷": "公顷", "亩": "亩",
    "%": "%", "％": "%", "个百分点": "个百分点",
    "户": "户", "人": "人", "个": "个", "处": "处", "座": "座",
    "栋": "栋", "条": "条", "项": "项", "辆": "辆", "家": "家",
    "元": "元", "万元": "万元", "亿元": "亿元",
    "吨": "吨", "万吨": "万吨", "千瓦": "千瓦", "千伏": "千伏",
    "小时": "小时", "天": "天", "秒": "秒", "度": "度", "℃": "℃",
}

# 派生值标注词：正文里出现这些词，该行的数值是按材料算出的派生量
# （如「按台账加总 22 栋」「测算占 60.0%」），不要求与事实库逐字命中——
# 事实库存的是分项原值，派生值本就不在其中。
DERIVED_RE = re.compile(r"测算|加总|计算|折算|相除|推算|除以|累加|合计")

# 派生值登记用的「值+单位」形态（单位按长优先，避免「万」抢在「万平方米」前面）
# 2026-11 补齐 ㎡ / m² / m 等符号单位：此前单位枚举只有「平方米」，
# 正文写「17420 ㎡」「47333 ㎡」时整条数字都不进回溯表（实测 09 自检零登记），
# 派生值因此既不登记也不校验，是本轮三处数字硬伤漏检的直接原因。
DERIVED_TOKEN_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*(万平方米|万㎡|平方米|㎡|平米|m²|万个|万|个百分点|公里|公顷|"
    r"万元|亿元|元|万户|万人|户|栋|个|处|座|条|项|辆|家|米|m|吨|千瓦|千伏|人|年|月|天|小时|%)")


def norm_text(text: str) -> str:
    """去掉所有空白（正文与事实的空格差异不该导致匹配失败）。"""
    return re.sub(r"\s+", "", str(text or ""))


def numkey(value) -> str:
    """数值归一：744.0 → 744；6.0 → 6；0.252 → 0.252；1.4255000000000001 → 1.4255。"""
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if f == int(f):
        return str(int(f))
    return "%g" % f


def canon_unit(unit: str) -> str:
    """单位别名归一：万㎡ → 万平方米；㎡ → 平方米；m → 米。"""
    u = str(unit or "").strip()
    return UNIT_ALIAS.get(u, u)


def canon_token(value, unit: str) -> str:
    """事实的规范化 token（数值归一 + 单位别名），用于正文匹配。"""
    if value is None:
        return ""
    return norm_text(numkey(value) + canon_unit(unit))


def matches_in_report(value, unit: str, quote: str, report_text: str) -> bool:
    """判断一条事实是否在正文出现。三路匹配，与 build_trace 一致。

    ① 引文前缀命中（最可靠，优先）
    ② 原样 token 命中
    ③ 归一 + 单位别名后命中
    """
    if not report_text:
        return False
    q = norm_text(quote)
    if q and len(q) >= 12 and q[:40] in report_text:
        return True
    if value is None:
        return False
    token = norm_text(f"{value}{unit or ''}")
    if token and token in report_text:
        return True
    ctoken = canon_token(value, unit)
    return bool(ctoken and ctoken in report_text)


def derived_tokens(report_text: str) -> list[tuple[str, str, int]]:
    """扫出正文里标了「加总/测算」的数值，返回 [(值, 单位, 行号)]。"""
    out = []
    for no, ln in enumerate(str(report_text or "").splitlines(), start=1):
        if not DERIVED_RE.search(ln):
            continue
        for m in DERIVED_TOKEN_RE.finditer(ln):
            out.append((m.group(1), m.group(2), no))
    return out


# ───────────────────────── 算术对账（CHK-NUM-2） ─────────────────────────
# 背景：派生值此前只要带了「加总/测算」标注就直接记 confirmed，**不做任何算术校验**，
# 于是「分项 19720 却写合计 17420」「10.2945 万㎡ 占 47333 ㎡ 主要部分（217%）」
# 「按 13% 计提却写 15% 的金额」三处硬伤全部通过 09 自检。
# 这里把三类可判定的关系写成硬校验，任一不成立即计 conflict：
#   ① 分项加总 = 声明合计    ② 占比 = 分子 ÷ 分母    ③ 计提金额 = 比例 × 基数
# 另加 ④ 同一指标在多处出现多个值（面积/用地/立面工程量）。
# 只做**格式可判定**的校验，不做语义理解；判不了的返回 unverified，不冒充通过。

_RNUM = r"\d[\d,]*(?:\.\d+)?"
_UNIT_ALT = ("万平方米|万㎡|平方米|㎡|平米|m²|公顷|亩|"
             "万元|亿元|元|栋|户|个|处|座|条|项|辆|家|人|吨|米|m")
_TOKEN_RE = re.compile(r"(" + _RNUM + r")\s*(" + _UNIT_ALT + r")")
_TOTAL_RE = re.compile(r"(?:加总|合计|共计|总计)(?:为|是|达|约)?\s*(" + _RNUM + r")\s*(" + _UNIT_ALT + r")")
_RATIO_RE = re.compile(r"占([^，。；]{0,20}?)(" + _RNUM + r")\s*(万平方米|万㎡|平方米|㎡)\s*的\s*(\d+(?:\.\d+)?)\s*%")
_ACCRUAL_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%\s*计提")
_RANGE_RE = re.compile(r"(?<![\d.,])(" + _RNUM + r")\s*[～~\-—]\s*(" + _RNUM + r")(?![\d.])")

_COUNT_UNITS = {"栋", "户", "个", "处", "座", "条", "项", "辆", "家", "人"}
_CROSS_METRICS = {
    "建筑面积": re.compile(r"建筑面积[^，。；\n]{0,6}?(" + _RNUM + r")\s*(万平方米|万㎡|平方米|㎡)"),
    "用地面积": re.compile(r"用地面积[^，。；\n]{0,6}?(" + _RNUM + r")\s*(万平方米|万㎡|平方米|㎡)"),
    "立面工程量": re.compile(r"立面(?:修缮|工程量)[^，。；\n]{0,16}?(" + _RNUM + r")\s*(平方米|㎡)"),
    # 栋数 / 户数 / 投资额只认**总量**写法（前有合计/加总/共计/总计/共），
    # 逐栋逐户的分项值不参与，否则「11 栋 744 户」「2 栋 132 户」会互相打架。
    "合计栋数": re.compile(r"(?:合计|加总|共计|总计|共)[^，。；\n]{0,8}?(" + _RNUM + r")\s*(栋)"),
    "合计户数": re.compile(r"(?:合计|加总|共计|总计|共)[^，。；\n]{0,8}?(" + _RNUM + r")\s*(户)"),
    "总投资额": re.compile(r"(?:总投资|投资总额|总投资额)[^，。；\n]{0,8}?(" + _RNUM + r")\s*(万元|亿元)"),
}


def _f(text) -> float:
    try:
        return float(str(text).replace(",", ""))
    except (TypeError, ValueError):
        return 0.0


def _strip_parens(line: str) -> str:
    """去掉括号内容：括注多为另一口径的补充值，参与加总会重复计数。"""
    out, depth = [], 0
    for ch in line:
        if ch in "（(【":
            depth += 1
            continue
        if ch in "）)】":
            depth = max(0, depth - 1)
            continue
        if depth == 0:
            out.append(ch)
    return "".join(out)


def _family(unit: str) -> str:
    """量纲族：万平方米与平方米同族（只在书写差一个万，不代表不同指标）。"""
    u = canon_unit(unit)
    return "平方米" if u == "万平方米" else u


def _factor(unit: str) -> float:
    """到族内基准单位的倍数（万平方米 → ×10000）。"""
    return 10000.0 if canon_unit(unit) == "万平方米" else 1.0


def reconcile(report_text: str) -> dict:
    """对正文做算术对账。

    返回 {"issues": [{kind, line, message}], "verified": set[(值, 归一单位, 行号)],
          "checkedLines": set[int]}。
    issues 非空即应判 fail；verified 是真正算通了的派生值；
    未出现在 verified / issues 里的派生值属「无法用格式校验」，调用方应记 unverified。
    """
    issues: list[dict] = []
    verified: set[tuple[str, str, int]] = set()
    checked_lines: set[int] = set()
    rows = str(report_text or "").splitlines()

    # ① 分项加总 = 声明合计（按 ；。 切句，避免同一行多个合计互相干扰）
    for no, raw in enumerate(rows, start=1):
        line = _strip_parens(raw)
        for seg in re.split(r"[；;。]", line):
            totals = list(_TOTAL_RE.finditer(seg))
            if len(totals) != 1:
                continue
            tm = totals[0]
            tval, tun = _f(tm.group(1)), canon_unit(tm.group(2))
            if tun not in _COUNT_UNITS and tun not in {"平方米", "万平方米", "米", "万元", "亿元"}:
                continue
            parts = []
            for m in _TOKEN_RE.finditer(seg):
                if tm.start(1) <= m.start(1) < tm.end():
                    continue
                if canon_unit(m.group(2)) != tun:
                    continue
                parts.append(_f(m.group(1)))
            if len(parts) < 2:
                continue
            checked_lines.add(no)
            part_sum = sum(parts)
            tol = 0.0 if tun in _COUNT_UNITS else max(1.0, abs(tval) * 0.01)
            if abs(part_sum - tval) > tol:
                issues.append({
                    "kind": "sum", "line": no, "lines": [no],
                    "message": (f"分项加总 {numkey(part_sum)}{tun} ≠ 声明合计 {numkey(tval)}{tun}（差 {numkey(round(abs(part_sum - tval), 4))}{tun}，行 {no}）")})
            else:
                verified.add((numkey(tval), tun, no))

    # ② 占比 = 分子 ÷ 分母
    #    分子优先级：同句最近的「合计/加总」→ 同句最近的同族数值。
    #    只认「合计」会让「A、B、C 三处合计 15350 ㎡，占…的 77.4%」以外的写法
    #    （如「立面修缮累计 19820 ㎡，占…的 88.1%」）整条规则失效——实测漏判。
    for no, raw in enumerate(rows, start=1):
        line = _strip_parens(raw)
        for m in _RATIO_RE.finditer(line):
            den = _f(m.group(2)) * _factor(m.group(3))
            pct = _f(m.group(4))
            if den <= 0:
                continue
            head_txt = line[:m.start()]
            heads = list(_TOTAL_RE.finditer(head_txt))
            if heads:
                num = _f(heads[-1].group(1)) * _factor(heads[-1].group(2))
            else:
                toks = [t for t in _TOKEN_RE.finditer(head_txt)
                        if _family(t.group(2)) == "平方米"]
                if not toks:
                    continue
                num = _f(toks[-1].group(1)) * _factor(toks[-1].group(2))
            checked_lines.add(no)
            exp = num / den * 100
            # 百分比写成整数时四舍五入误差可达 0.5pp；写一位小数时 ≤0.05pp
            tol = 0.15 if "." in m.group(4) else 0.55
            if abs(exp - pct) > tol:
                issues.append({
                    "kind": "ratio", "line": no, "lines": [no],
                    "message": (f"占比 {numkey(pct)}% ≠ {numkey(num)} ÷ {numkey(den)} = {exp:.1f}%（行 {no}）")})
            else:
                verified.add((numkey(num), "平方米", no))

    # ③ 计提金额 = 比例 × 基数（表格行；基数取「前 N 项」的金额区间之和）
    idx = 0
    while idx < len(rows):
        if not rows[idx].lstrip().startswith("|"):
            idx += 1
            continue
        start = idx
        while idx < len(rows) and rows[idx].lstrip().startswith("|"):
            idx += 1
        data: list[tuple[int, list[str]]] = []
        for i in range(start, idx):
            cells = [c.strip() for c in rows[i].strip().strip("|").split("|")]
            if set("".join(cells)) <= set("-: "):
                continue
            data.append((i + 1, cells))
        for pos, (no, cells) in enumerate(data):
            joined = "|".join(cells)
            am = _ACCRUAL_RE.search(joined)
            if not am:
                continue
            pct = _f(am.group(1))
            nm = re.search(r"前(\d+)项", joined)
            need = int(nm.group(1)) if nm else 3
            amount = None
            for cell in cells:
                rm = _RANGE_RE.search(cell)
                if rm:
                    amount = (_f(rm.group(1)), _f(rm.group(2)))
                    break
            if not amount:
                continue
            base = [c for _, c in data[:pos] if "合计" not in "|".join(c)][-need:]
            if len(base) < need:
                # 「前三项」写在表首（基数行在其后）时，改用后续行；遇到合计行即止
                base = []
                for _, c in data[pos + 1:]:
                    if "合计" in "|".join(c):
                        break
                    base.append(c)
                base = base[:need]
            if len(base) < need:
                continue
            lows = highs = 0.0
            ok = True
            for cell in base:
                rm = _RANGE_RE.search("|".join(cell))
                if not rm:
                    ok = False
                    break
                lows += _f(rm.group(1))
                highs += _f(rm.group(2))
            if not ok:
                continue
            checked_lines.add(no)
            exp_lo, exp_hi = lows * pct / 100, highs * pct / 100
            tol = max(10.0, 0.02 * max(exp_hi, 1.0))
            if abs(amount[0] - exp_lo) > tol or abs(amount[1] - exp_hi) > tol:
                issues.append({
                    "kind": "accrual", "line": no, "lines": [no],
                    "message": (f"「按前{need}项合计的 {numkey(pct)}% 计提」与金额 "
                                f"{numkey(amount[0])}～{numkey(amount[1])} 不符："
                                f"{numkey(lows)}～{numkey(highs)} × {numkey(pct)}% = "
                                f"{exp_lo:.1f}～{exp_hi:.1f}（行 {no}）")})
            else:
                verified.add((numkey(amount[0]), "万元", no))

    # ④ 同一指标多个值
    seen: dict[tuple[str, str], dict[str, list[int]]] = {}
    for no, raw in enumerate(rows, start=1):
        line = _strip_parens(raw)
        for name, rx in _CROSS_METRICS.items():
            for m in rx.finditer(line):
                val, unit = _f(m.group(1)), canon_unit(m.group(2))
                if name == "建筑面积":
                    val = val / 10000 if unit == "平方米" else val
                    if val < 1:          # 单个对象面积不是「总量」
                        continue
                    unit = "万平方米"
                elif name == "立面工程量":
                    if val < 1000:       # 单个对象工程量不是「总量」
                        continue
                    unit = "平方米"
                seen.setdefault((name, unit), {}).setdefault(numkey(val), []).append(no)
                checked_lines.add(no)
    for (name, unit), vals in seen.items():
        if len(vals) > 1:
            detail = "；".join(
                f"{k}{unit}（行 {','.join(map(str, sorted(set(v))))}）" for k, v in sorted(vals.items()))
            all_lines = sorted({n for v in vals.values() for n in v})
            issues.append({
                "kind": "cross", "line": all_lines[0], "lines": all_lines,
                "message": f"同一指标「{name}」出现多个值：{detail}"})

    return {"issues": issues, "verified": verified, "checkedLines": checked_lines}

