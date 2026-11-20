#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_skill_paths.py -- 资源定位底座（只用标准库）：自包含副本优先，兄弟 skill 仅作回退。

解决的问题：本 skill 被整体剪切到独立目录后，脚本里写死的「兄弟 skill 路径」
（如 `<skill上级>/cardo-status-spec/规则/章节契约.yaml`、`<skill上级>/cardo-project-material/规则/...`）
不可达，导致：
  * assemble_chapters.contract_level1_units() 恒返回空表 → 缺项字母归一（remap）永不生效
    → 成稿缺项字母与章节契约 writing_unit 不一致 → CHK-GAP-1 必然 FAIL；
  * check_contracts 的「必备材料清单 chapters 含非模板 H1 名」子项被静默跳过。

定位优先级（自包含优先，符合「规则唯一真相在 spec、调优副本以自带为准」的约定）：
  1. 调用方显式给出的候选（prefer）；
  2. 本 skill 自带副本：<skill>/规则齐套/<name> → <skill>/规则/<name>；
  3. 兄弟 skill 回退（插件原布局与独立调优布局都支持）：
     <root>/<peer>/skill/规则/<name> → <root>/<peer>/规则/<name>
     其中 peer ∈ {cardo-project-material, cardo-status-spec}，
     root ∈ {$CARDO_STATUS_PEERS_ROOT, <skill上级>, <skill上上级>}。

全部落空返回 ""（调用方按「缺失」处理，不静默当成功）。
"""

from __future__ import annotations

import os

__version__ = "1.0.0"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)

PEERS = ("cardo-project-material", "cardo-status-spec")


def peer_roots() -> list[str]:
    """兄弟 skill 根候选：环境变量 → skill 上级 → skill 上上级。"""
    roots: list[str] = []
    env = os.environ.get("CARDO_STATUS_PEERS_ROOT", "").strip()
    if env:
        roots.append(os.path.abspath(env))
    roots.append(os.path.dirname(SKILL_DIR))
    roots.append(os.path.dirname(os.path.dirname(SKILL_DIR)))
    out: list[str] = []
    for r in roots:
        if r not in out:
            out.append(r)
    return out


def own_rule_candidates(name: str) -> list[str]:
    """本 skill 自带规则副本候选（齐套规则在前，其次是契约规则）。"""
    return [os.path.join(SKILL_DIR, "规则齐套", name),
            os.path.join(SKILL_DIR, "规则", name)]


def peer_rule_candidates(name: str) -> list[str]:
    """兄弟 skill 的规则文件候选（插件原布局 / 独立调优布局）。"""
    out: list[str] = []
    for root in peer_roots():
        for peer in PEERS:
            for sub in (os.path.join(peer, "skill", "规则"), os.path.join(peer, "规则")):
                cand = os.path.join(root, sub, name)
                if cand not in out:
                    out.append(cand)
    return out


def resolve_rule(name: str, prefer: list[str] | None = None) -> str:
    """返回第一个存在的规则文件路径；都不存在返回 ""。"""
    for cand in list(prefer or []) + own_rule_candidates(name) + peer_rule_candidates(name):
        if cand and os.path.isfile(cand):
            return cand
    return ""


def resolve_rule_or(name: str, fallback: str, prefer: list[str] | None = None) -> str:
    """同 resolve_rule，但落空时返回 fallback（原路径，用于保持既有报错语义）。"""
    return resolve_rule(name, prefer) or fallback
