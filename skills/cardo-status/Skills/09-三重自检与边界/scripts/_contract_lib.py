#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_contract_lib.py -- 契约工具共享底座（只用标准库，给 check_contracts / extract_contract 用）。

这里只放"两种脚本都要用、且与领域无关"的东西：

  * YAML 子集解析：块映射 / 块序列 / 行内列表 / 引号标量 / 行尾注释。
    契约与规则文件都控制在这个子集里；不依赖 pyyaml（运行环境不保证有）。
  * YAML 发射：把标量/行内列表写成稳定、可 diff 的文本（抽取器用，保证逐字节可复现）。
  * 比对口径：空白、全半角括号、大小写归一。
  * markdown 结构：标题序列、标题作用域、表格（表头 + 数据行）。
  * 路径：读取与"相对化"输出（产物里不出现盘符绝对路径）。

约定：本模块不读业务文件、不写任何文件，纯函数。
"""

from __future__ import annotations

import hashlib
import os
import re

# 读取
def read_text(path: str) -> str:
    """UTF-8（容忍 BOM）读取；解码失败用 replace，保证只读校验器不因坏字节崩。"""
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        return f.read()


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def rel_posix(path: str, base: str) -> str:
    """相对 base 的 posix 路径；跨盘等无法相对化时退化为文件名（绝不输出盘符）。"""
    try:
        r = os.path.relpath(os.path.abspath(path), os.path.abspath(base))
    except ValueError:
        return os.path.basename(path)
    return r.replace(os.sep, "/")


# YAML 子集：解析
def _strip_comment(s: str) -> str:
    """去掉行尾注释（引号内的 # 保留；仅当 # 前是空白时视为注释起点）。"""
    quote = ""
    for i, ch in enumerate(s):
        if quote:
            if ch == quote:
                quote = ""
        elif ch in "\"'":
            quote = ch
        elif ch == "#" and (i == 0 or s[i - 1] in " \t"):
            return s[:i].rstrip()
    return s.rstrip()


def _split_key(s: str) -> tuple[str, str]:
    """按第一个引号外的 ':' 切 key / value。"""
    quote = ""
    for i, ch in enumerate(s):
        if quote:
            if ch == quote:
                quote = ""
        elif ch in "\"'":
            quote = ch
        elif ch == ":":
            return s[:i].strip(), s[i + 1:].strip()
    return s.strip(), ""


def _split_flow(s: str) -> list[str]:
    """按顶层逗号切分行内集合（尊重引号与 []{} 嵌套）。"""
    items, buf, depth, quote = [], "", 0, ""
    for ch in s:
        if quote:
            buf += ch
            if ch == quote:
                quote = ""
        elif ch in "\"'":
            quote = ch
            buf += ch
        elif ch in "[{":
            depth += 1
            buf += ch
        elif ch in "]}":
            depth -= 1
            buf += ch
        elif ch == "," and depth == 0:
            items.append(buf)
            buf = ""
        else:
            buf += ch
    if buf.strip():
        items.append(buf)
    return items


def _parse_flow(s: str):
    """行内集合：[a, b] 或 {k: v, ...}（可嵌套）；其它按标量。"""
    s = s.strip()
    if s.startswith("[") and s.endswith("]"):
        return [_parse_flow(x) for x in _split_flow(s[1:-1])]
    if s.startswith("{") and s.endswith("}"):
        out = {}
        for part in _split_flow(s[1:-1]):
            key, val = _split_key(part)
            out[key.strip()] = _parse_flow(val)
        return out
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        return s[1:-1].replace('\\"', '"').replace("\\\\", "\\")
    if s == "":
        return ""
    if s in ("true", "True"):
        return True
    if s in ("false", "False"):
        return False
    if s in ("null", "~"):
        return None
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d+\.\d+", s):
        return float(s)
    return s


def _scalar(s: str):
    s = s.strip()
    if s.startswith(("[", "{")):
        return _parse_flow(s)
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        inner = s[1:-1]
        return inner.replace('\\"', '"').replace("\\\\", "\\")
    if s == "":
        return ""
    if s in ("true", "True"):
        return True
    if s in ("false", "False"):
        return False
    if s in ("null", "~"):
        return None
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d+\.\d+", s):
        return float(s)
    return s


def parse_yaml(text: str):
    """解析契约/规则使用的 YAML 子集。结构错误抛 ValueError（调用方按参数错处理）。"""
    lines: list[tuple[int, str]] = []
    for raw in text.splitlines():
        if not raw.strip():
            continue
        indent = len(raw) - len(raw.lstrip(" "))
        body = _strip_comment(raw.strip())
        if not body:
            continue
        lines.append((indent, body))
    if not lines:
        return {}
    val, _ = _parse_block(lines, 0, lines[0][0])
    return val


def _parse_block(lines: list, i: int, indent: int):
    if i >= len(lines):
        return None, i
    ind, s = lines[i]
    if ind < indent:
        return None, i
    if s == "-" or s.startswith("- "):
        return _parse_seq(lines, i, ind)
    return _parse_map(lines, i, ind)


def _parse_seq(lines: list, i: int, indent: int):
    out: list = []
    while i < len(lines):
        ind, s = lines[i]
        if ind != indent or not (s == "-" or s.startswith("- ")):
            break
        rest = s[1:].strip()
        i += 1
        if not rest:
            if i < len(lines) and lines[i][0] > indent:
                val, i = _parse_block(lines, i, lines[i][0])
            else:
                val = ""
            out.append(val)
            continue
        if rest.startswith("{"):
            out.append(_parse_flow(rest))
            continue
        key, val = _split_key(rest)
        if ":" in rest and not rest.startswith(("'", '"')):
            item = {}
            if val == "":
                if i < len(lines) and lines[i][0] > indent + 1:
                    sub, i = _parse_block(lines, i, lines[i][0])
                else:
                    sub = ""
                item[key] = sub
            else:
                item[key] = _scalar(val)
            if i < len(lines) and lines[i][0] > indent:
                more, i = _parse_map(lines, i, lines[i][0], item)
                item = more
            out.append(item)
        else:
            out.append(_scalar(rest))
    return out, i


def _parse_map(lines: list, i: int, indent: int, out: dict | None = None):
    out = {} if out is None else out
    while i < len(lines):
        ind, s = lines[i]
        if ind < indent or (ind == indent and (s == "-" or s.startswith("- "))):
            break
        if ind != indent:
            break
        key, val = _split_key(s)
        if not key:
            break
        i += 1
        if val == "":
            if i < len(lines) and lines[i][0] > indent:
                sub, i = _parse_block(lines, i, lines[i][0])
            else:
                sub = ""
            out[key] = sub
        else:
            out[key] = _scalar(val)
    return out, i


# YAML 子集：发射（稳定输出，保证抽取器逐字节可复现）
def yaml_scalar(value) -> str:
    if value is None:
        return "null"
    if value is True:
        return "true"
    if value is False:
        return "false"
    if isinstance(value, (int, float)):
        return str(value)
    s = str(value)
    if s == "" or re.search(r"^[\s\-?:,\[\]{}#&*!|>'\"%@`]|[:#]\s|\s$", s):
        return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return s


def yaml_flow(items) -> str:
    """行内列表（契约里 keywords/chapters 等短列表用）。"""
    return "[" + ", ".join(yaml_scalar(x) for x in items) + "]"


# 比对口径
def norm_title(text: str) -> str:
    """标题/键比对口径：去空白，全角括号与冒号转半角，小写。"""
    s = (text or "").strip().replace("（", "(").replace("）", ")").replace("：", ":")
    s = re.sub(r"\s+", "", s)
    return s.lower()


def norm_cell(text: str) -> str:
    """表格单元格比对口径：在标题口径上再去掉 %/％ 差异与首尾标点空白。"""
    s = norm_title(text).replace("％", "%").replace("／", "/")
    return s.strip("| ")


# markdown 结构
HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
SOURCE_ROW_RE = re.compile(r"^（数据来源：")
GAP_INLINE_RE = re.compile(r"【缺项-([A-Za-z]?)(\d{1,3})：")


def iter_headings(lines: list[str]) -> list[dict]:
    out = []
    for no, line in enumerate(lines):
        m = HEADING_RE.match(line)
        if m:
            out.append({"line": no, "level": len(m.group(1)), "title": m.group(2).strip()})
    return out


def heading_scope(lines: list[str], headings: list[dict]) -> list[dict]:
    """给每个标题挂上作用域 [start, end)（到下一个同级或更高级标题为止）。"""
    for k, h in enumerate(headings):
        end = len(lines)
        for nxt in headings[k + 1:]:
            if nxt["level"] <= h["level"]:
                end = nxt["line"]
                break
        h["start"] = h["line"]
        h["end"] = end
    return headings


def table_sep(line: str) -> bool:
    s = line.strip()
    if not s.startswith("|") or "|" not in s[1:]:
        return False
    return bool(re.fullmatch(r"\|?[\s:|\-]+\|?[\s:|\-]*", s))


def split_table_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    parts, buf, esc = [], "", False
    for ch in s:
        if esc:
            buf += ch
            esc = False
        elif ch == "\\":
            buf += ch
            esc = True
        elif ch == "|":
            parts.append(buf.strip())
            buf = ""
        else:
            buf += ch
    parts.append(buf.strip())
    return parts


def iter_tables(lines: list[str]) -> list[dict]:
    """markdown 表格：表头行 + 分隔行 + 数据行；返回 line/header/rows/end_line。"""
    out, i = [], 0
    while i < len(lines) - 1:
        if lines[i].lstrip().startswith("|") and table_sep(lines[i + 1]):
            header = split_table_row(lines[i])
            rows, j = [], i + 2
            while j < len(lines) and lines[j].lstrip().startswith("|"):
                if table_sep(lines[j]):
                    j += 1
                    continue
                rows.append(split_table_row(lines[j]))
                j += 1
            out.append({"line": i, "end_line": j, "header": header, "rows": rows})
            i = j
        else:
            i += 1
    return out


def header_match(header: list[str], columns: list[str], mode: str) -> bool:
    """表头是否匹配契约列：exact=列名顺序全等；subset=契约列必须都在。"""
    hdr = [norm_cell(c) for c in header if str(c).strip()]
    cols = [norm_cell(c) for c in columns if str(c).strip()]
    if mode == "subset":
        return all(c in hdr for c in cols)
    return hdr == cols


# 数值：单位归一 + 范围拆分（CHK-NUM-1 动态溯源用，不硬编码任何项目数字）
UNIT_ALT = ("平方公里|平方米|立方米|千伏安|万元|亿元|万人|个月|千伏|千瓦|公里|公顷|"
            "毫米|厘米|小时|分钟|分贝|米|万|亿|亩|个|户|人|处|栋|座|条|项|辆|吨|"
            "年|月|日|天|秒|%|℃|度")
NUM_RE = re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)(?:\s*[%％～~至—\-]\s*(\d+(?:\.\d+)?))?"
                    r"\s*(" + UNIT_ALT + r")")
UNIT_DROP = {"年", "月"}
UNIT_MAP = {"％": "%", "平方千米": "平方公里", "千米": "公里", "个月": "个月"}


def normalize_units(text: str) -> str:
    """把 ASCII/符号写法归一到中文单位（只换写法，不动数值）。"""
    reps = [
        (r"km²|km2|KM²|KM2", "平方公里"), (r"㎡|m²|m2|M²|M2", "平方米"),
        (r"m³|m3|M³|M3", "立方米"), (r"kVA|KVA|kva", "千伏安"),
        (r"kV|KV|kv", "千伏"), (r"kW|KW|kw", "千瓦"), (r"km|KM", "公里"),
        (r"(?<![A-Za-z])m(?![A-Za-z])", "米"), (r"％", "%"),
    ]
    for pat, rep in reps:
        text = re.sub(pat, rep, text)
    return text


# 裸数字（不带单位）——用于溯源校验的「数值命中、单位由表头/列名补足」判据。
# 表格里的数值在转储里通常只有数字本身（单位在表头或列名里），
# 如附表「现状基本情况」的主楼层面积列，源里只有 512.45。
BARE_NUM_RE = re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)(?![\d.])")


def bare_numbers(text: str) -> set[str]:
    """抽取文本中所有裸数值（不要求带单位），返回数值字符串集合。"""
    return set(BARE_NUM_RE.findall(text or ""))


def numeric_pairs(text: str) -> list[tuple[str, str]]:
    """抽取（数值, 归一单位）；区间拆两端各一条（470～770个 -> 470个/770个）。"""
    out = []
    for m in NUM_RE.finditer(text):
        unit = UNIT_MAP.get(m.group(3), m.group(3))
        if unit in UNIT_DROP:
            continue
        out.append((m.group(1), unit))
        if m.group(2):
            out.append((m.group(2), unit))
    return out
