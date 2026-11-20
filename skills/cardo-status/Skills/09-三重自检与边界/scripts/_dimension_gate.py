#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_dimension_gate.py -- 派生量量纲闸门：带单位 + 落区间，任一条不过即判违规。

为什么单独一条闸：正文里的「每公顷 / 每人 / 每栋」类派生量，数值本身能算通，
但分母选错（把万平方米当公顷、把建筑面积当用地面积）时仍会算出「像模像样」的
错数——3,142.79 人/公顷 就是这样出的厂。所以派生量除了可复算，还必须过两道硬检查：

  1. 必须带单位（关键词后 40 字内出现数字就必须跟单位），否则判「缺单位」；
  2. 数值必须落在该指标的量级区间内，否则判「越界」。

区间取自老旧街区常态：人口密度 100–2,000 人/公顷（低了不是城区、高了不是居住区），
人均绿地 2–50 平方米/人（低于 2 说明分子只看测绘范围、分母却按小区台账人口，不同源；
高于 50 说明分母漏乘或单位换算错），户均建筑面积 20–200 平方米。
"""
from __future__ import annotations

import re

# keyword 用正则（允许「人均公共绿地」这类变体）；unit 必须紧跟数值出现
RULES = [
    {"metric": "人口密度", "keyword": r"人口密度",
     "unit": r"(?:人\s*[/／]\s*(?:公顷|ha|hm2)|人\s*每\s*公顷)",
     "unit_label": "人/公顷", "low": 100.0, "high": 2000.0},
    {"metric": "人均绿地", "keyword": r"人均(?:公共)?绿地", "unit": r"平方米",
     "unit_label": "平方米/人", "low": 2.0, "high": 50.0},
    {"metric": "户均面积", "keyword": r"户均(?:建筑)?面积", "unit": r"平方米",
     "unit_label": "平方米/户", "low": 20.0, "high": 200.0},
    {"metric": "单栋基底", "keyword": r"单栋基底", "unit": r"平方米",
     "unit_label": "平方米/栋", "low": 10.0, "high": 20000.0},
    {"metric": "单栋地上建筑", "keyword": r"单栋地上建筑", "unit": r"平方米",
     "unit_label": "平方米/栋", "low": 10.0, "high": 20000.0},
    {"metric": "道路覆盖密度", "keyword": r"(?:道路覆盖密度|路网密度)",
     "unit": r"平方米\s*[/／]\s*公顷",
     "unit_label": "平方米/公顷", "low": 0.1, "high": 10000.0},
    {"metric": "每公顷道路面积", "keyword": r"每公顷", "unit": r"平方米",
     "unit_label": "平方米/公顷", "low": 0.1, "high": 10000.0},
]

NUM_RE = re.compile(r"\d[\d,]*(?:\.\d+)?")
WINDOW = 40          # 关键词之后取多少字找数值
# 句读边界：数值只在本分句内认领，避免把下一句的数字算到本条指标头上
# （「本报告不计算人口密度：……台账 10.0425 万平方米」这类口径说明不算指标值）
SENT_END = "。；：，\n"


def _num_after(chunk: str):
    m = NUM_RE.search(chunk)
    if not m:
        return None, ""
    raw = m.group(0)
    try:
        return float(raw.replace(",", "")), raw
    except ValueError:
        return None, raw


def scan(text: str) -> list[dict]:
    """返回违规列表；空列表表示通过。"""
    # PDF 抽取会把「平方米」按行折断成「平方\n米」，先去掉全部空白再认领，
    # 否则单位被换行截断会误判成「缺单位」。
    text = re.sub(r"\s+", "", text)
    bad: list[dict] = []
    for rule in RULES:
        for hit in re.finditer(rule["keyword"], text):
            tail = text[hit.end(): hit.end() + WINDOW]
            # 只在本句内找数，避免把下一句的数字算进来
            cut = min([tail.find(c) for c in SENT_END if tail.find(c) >= 0] or [len(tail)])
            chunk = tail[:cut]
            # 关键词在数值之后（如「2,000 人/公顷的人口密度」）不在本闸覆盖范围
            val, raw = _num_after(chunk)
            if val is None:
                continue
            if not re.search(rule["unit"], chunk):
                bad.append({"metric": rule["metric"], "value": raw, "unit": "",
                            "detail": f"「{rule['metric']}」的数值 {raw} 缺单位"
                                      f"（应为 {rule['unit_label']}）"})
                continue
            if not (rule["low"] <= val <= rule["high"]):
                bad.append({"metric": rule["metric"], "value": raw,
                            "unit": rule["unit_label"],
                            "detail": f"「{rule['metric']}」= {raw} {rule['unit_label']} "
                                      f"越界（合理区间 {rule['low']:g}–{rule['high']:g}），"
                                      f"请检查分母口径与单位换算"})
    return bad


def summarize(bad: list[dict]) -> str:
    if not bad:
        return "派生量量纲自检通过"
    return "；".join(b["detail"] for b in bad[:5])
