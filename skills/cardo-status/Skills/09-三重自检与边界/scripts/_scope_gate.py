#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_scope_gate.py -- 口径范围闸门：样本值不得冒充片区整体值。

背景：OSGB 测绘范围只有 1.22 公顷（bbox），是样本，不是片区总用地；台账给出的是
7 个小区的户数与住宅建筑面积（10.0425 万平方米），两者对象范围不同。报告里凡是
把建筑密度、绿地率、容积率、道路占比这类**面积型指标**说成「片区……」的句子，
读者会默认它是全片区值——这是「看起来专业、其实范围错位」的错，比算错数更难发现。

规则：同一句里同时出现「片区」与上述面积型指标词，必须同时出现「实测样区」「样区」
或「样本」（如「实测样区建筑密度 36.18%」「片区整体以台账为准」），否则 FAIL。
"""
from __future__ import annotations

import re

METRICS = ("建筑密度", "绿地率", "绿化率", "容积率", "道路占比")
SCOPE_WORDS = ("实测样区", "样区", "样本")
SPLIT_RE = re.compile(r"(?<=[。；;！？\n])")


def scan(text: str) -> list[dict]:
    bad: list[dict] = []
    for sent in SPLIT_RE.split(text or ""):
        s = sent.strip().replace("\n", "")
        if not s or "片区" not in s:
            continue
        hit = [m for m in METRICS if m in s]
        if not hit:
            continue
        if any(w in s for w in SCOPE_WORDS):
            continue
        bad.append({"metrics": hit, "sentence": s[:80]})
    return bad


def summarize(bad: list[dict]) -> str:
    if not bad:
        return "口径范围自检通过"
    return "；".join(f"「{b['sentence']}」" for b in bad[:3])
