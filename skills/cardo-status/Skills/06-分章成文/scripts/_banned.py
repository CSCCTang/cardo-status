#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""禁句单一来源的读取器（唯一真相：`06-分章成文/规则/禁句表.yaml`）。

设计要点：
- 只解析本表的固定行格式，不引第三方 YAML 依赖（避免与 gen_chapters 循环导入）。
- **文件缺失/解析失败时回退内置常量**（与引入本模块前的字面量逐字一致）：
  自检不能因为找不到规则文件就放行禁句。
- 导出三组：TODO_PHRASES（gap）/ EMPTY_PHRASES（empty）/ CLEAN_PHRASES（clean）。
"""
from __future__ import annotations

import re
from pathlib import Path

RULES_FILE = Path(__file__).resolve().parents[1] / "规则" / "禁句表.yaml"
LINE_RE = re.compile(r"-\s*\{\s*pattern:\s*([^,}]+?)\s*,\s*scopes:\s*\[([^\]]*)\]\s*\}")

# 回退常量：与 2026-11-11 前的硬编码完全一致，保证「读不到规则」时行为不降级。
_FALLBACK = [
    ("资料缺失", ("gap", "clean")),
    ("缺项", ("gap",)),
    ("待补", ("gap", "clean")),
    ("未提供", ("gap", "clean")),
    ("资料不足", ("gap", "clean")),
    ("本节不作推断", ("gap",)),
    ("不作推断", ("gap",)),
    ("缺项-", ("clean",)),
    ("找谁补", ("clean",)),
    ("缺什么", ("clean",)),
    ("高度重视", ("empty",)),
    ("进一步加强", ("empty",)),
    ("大力推进", ("empty",)),
    ("不断提升", ("empty",)),
    ("切实抓好", ("empty",)),
]


def load() -> list[tuple[str, tuple[str, ...]]]:
    """返回 [(pattern, scopes)]；读不到或解析为空时回退内置常量。"""
    try:
        text = RULES_FILE.read_text(encoding="utf-8")
    except OSError:
        return list(_FALLBACK)
    rows: list[tuple[str, tuple[str, ...]]] = []
    for line in text.splitlines():
        match = LINE_RE.search(line)
        if not match:
            continue
        scopes = tuple(s.strip() for s in match.group(2).split(",") if s.strip())
        if scopes:
            rows.append((match.group(1).strip(), scopes))
    return rows or list(_FALLBACK)


def by_scope(scope: str) -> list[str]:
    return [pattern for pattern, scopes in load() if scope in scopes]


TODO_PHRASES = by_scope("gap")
EMPTY_PHRASES = by_scope("empty")
CLEAN_PHRASES = by_scope("clean")
