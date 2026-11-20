#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_ledger.py -- 检索台账（cardo/research-ledger/v1）读取与校验的唯一实现。

为什么单独一个模块：台账要在三个地方被校验——开工.py（开工前）、
06 分章成文（写作前）、09 三重自检（成文后）。三处各写一套必然漂移，
所以规则只在这里定义，其余地方只调用。

台账不是"证据本体"，落盘文件才是：`file` 字段指向清洗后的正文文本，
正文里 `[S-012]` 等于"引用 S-012 那份文件"。缺 URL / 缺检索日期 /
缺原文摘录 / 文件不存在或为空 —— 任一成立，该条即无效（等同不存在）。
"""

from __future__ import annotations

import json
import os
import re

SCHEMA = "cardo/research-ledger/v1"
# 兼容两种工作区布局：独立工作区用「过程文件/」，标准项目库用「05_过程/」
REL_CANDIDATES = (
    "过程文件/00-总控/网络补料/检索台账.json",
    "05_过程/网络补料/检索台账.json",
)
CITE_RE = re.compile(r"\[(S-\d{2,})\]")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
URL_RE = re.compile(r"^https?://\S+$", re.IGNORECASE)
ID_RE = re.compile(r"^S-\d{2,}$")
MIN_QUOTE = 20

# 必填字段（notFound 条目另外要求 reason；value 与 notFound 互斥）
REQUIRED = ("id", "metric", "publisher", "title", "date", "quote", "scope")

# CITE-2「空头依据」触发词：只收"声称外部来源"的措辞。
# 「同类项目经验区间」属内部推断，不在此列（由 CITE-3 管置信度与缺口登记）。
CITE2_PHRASES = (
    "公开区间", "据公开资料", "据公开数据", "公开资料显示", "据行业报告",
    "行业报告显示", "据网络资料", "网络公开数据", "据网上资料", "网上公开",
)
# CITE-3：低置信推断若**声称外部核实过**，必须能在台账找到「已查无果」记录
CITE3_EXTERNAL = ("公开", "网络", "检索", "行业报告", "造价信息", "市场价")


def _s(value) -> str:
    return str(value or "").strip()


def find_ledger(project_dir: str) -> str:
    for rel in REL_CANDIDATES:
        path = os.path.join(project_dir, *rel.split("/"))
        if os.path.isfile(path):
            return path
    return ""


def values_of(entry: dict) -> list[float]:
    """value 支持『410～560』『约 12』『10.2 万』：拆出全部数值，供数字来源池比对。"""
    out: list[float] = []
    for m in re.finditer(r"\d+(?:\.\d+)?", _s(entry.get("value"))):
        try:
            out.append(float(m.group(0)))
        except ValueError:
            pass
    return out


def validate_entry(entry: dict, project_dir: str) -> list[str]:
    """返回该条目的问题清单；空列表＝有效。"""
    if not isinstance(entry, dict):
        return ["条目不是对象"]
    bad: list[str] = []
    eid = _s(entry.get("id"))
    if not ID_RE.match(eid):
        bad.append(f"id 非法：{eid!r}（应为 S-001 形式）")
    for field in REQUIRED:
        if not _s(entry.get(field)):
            bad.append(f"缺 {field}")
    date = _s(entry.get("date"))
    if date and not DATE_RE.match(date):
        bad.append(f"date 须为 YYYY-MM-DD：{date!r}")
    not_found = bool(entry.get("notFound"))
    if not_found:
        if not _s(entry.get("reason")):
            bad.append("已查无果条目缺 reason（付费墙/需登录/无公开数据）")
    else:
        url = _s(entry.get("url"))
        if not url:
            bad.append("缺 URL")
        elif not URL_RE.match(url):
            bad.append(f"URL 非法：{url!r}")
        if not _s(entry.get("value")):
            bad.append("缺 value（或标 notFound）")
    quote = _s(entry.get("quote"))
    if len(quote) < MIN_QUOTE:
        bad.append(f"原文摘录不足 {MIN_QUOTE} 字（现 {len(quote)} 字）")
    # 落盘文件：正文引用编号＝引用该文件，必须真实存在且非空
    rel = _s(entry.get("file"))
    if not rel:
        if not not_found:
            bad.append("缺 file（清洗后正文文本的相对路径）")
        return bad
    parts = rel.replace("\\", "/").split("/")
    if os.path.isabs(rel) or ".." in parts:
        bad.append(f"file 必须是相对项目根的路径：{rel!r}")
        return bad
    path = os.path.join(project_dir, *parts)
    if not os.path.isfile(path):
        bad.append(f"file 不存在：{rel}")
    elif os.path.getsize(path) == 0:
        bad.append(f"file 为空：{rel}")
    else:
        base = os.path.basename(path)
        if not base.startswith(eid + "_"):
            bad.append(f"文件名须以「{eid}_」开头：{base}")
        if not base.endswith(".txt"):
            bad.append(f"落盘文件须为 .txt：{base}")
        if date and date not in base:
            bad.append(f"文件名须含检索日期 {date}：{base}")
    return bad


def load(project_dir: str, path: str = "") -> dict:
    """读取并校验台账。

    返回 {path, entries, valid{id:entry}, invalid{id:[问题]}, problems:[...], dup:[id]}。
    无台账时 path 为空、其余为空集合——调用方自行决定是跳过还是判失败。
    """
    out = {"path": path or find_ledger(project_dir), "entries": [],
           "valid": {}, "invalid": {}, "problems": [], "dup": []}
    if not out["path"]:
        return out
    try:
        with open(out["path"], encoding="utf-8-sig") as fh:
            data = json.load(fh)
    except (OSError, ValueError) as exc:
        out["problems"].append(f"检索台账解析失败：{type(exc).__name__}")
        return out
    if not isinstance(data, dict) or data.get("schema") != SCHEMA:
        out["problems"].append(f"检索台账 schema 必须是 {SCHEMA}")
        return out
    seen: set[str] = set()
    for entry in data.get("entries") or []:
        out["entries"].append(entry)
        eid = _s(entry.get("id")) if isinstance(entry, dict) else ""
        if eid and eid in seen:
            out["dup"].append(eid)
            out["problems"].append(f"条目编号重复：{eid}（编号只增不改，不得复用）")
            continue
        seen.add(eid)
        bad = validate_entry(entry, project_dir)
        if bad:
            out["invalid"][eid or "?"] = bad
            out["problems"].append(f"条目 {eid or '?'} 无效：" + "；".join(bad))
        else:
            out["valid"][eid] = entry
    return out


def cited(text: str) -> list[tuple[str, int]]:
    """正文里全部 [S-xxx] 引用 → [(编号, 行号)]（行号从 1 起）。"""
    hits: list[tuple[str, int]] = []
    for no, line in enumerate(text.splitlines(), start=1):
        for m in CITE_RE.finditer(line):
            hits.append((m.group(1), no))
    return hits


def cite_refs_in_line(line: str) -> list[str]:
    return [m.group(1) for m in CITE_RE.finditer(line)]
