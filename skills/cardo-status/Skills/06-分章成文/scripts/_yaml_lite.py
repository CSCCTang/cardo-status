"""极简 YAML 子集解析器（不依赖 PyYAML）。

只支持本项目规则文件用到的写法：
- 映射（缩进嵌套）
- 列表：`- 值` 与 `- key: value` 映射项
- 行内列表：[a, b, "c"]
- 注释：整行 # 与「空格 + #」起的行尾注释
- 标量：字符串、整数、浮点、true/false、空值
不支持多行字符串（| >）、锚点、引用、复杂转义。

解析失败抛 YamlError，调用方应转成可读错误（含行号更佳）。
"""
from __future__ import annotations

import re


class YamlError(ValueError):
    """规则文件格式错误。"""


def _strip_comment(line: str) -> str:
    in_single = in_double = False
    for i, ch in enumerate(line):
        if ch == "'" and not in_double:
            in_single = not in_single
        elif ch == '"' and not in_single:
            in_double = not in_double
        elif ch == "#" and not in_single and not in_double:
            if i == 0 or line[i - 1] in " \t":
                return line[:i]
    return line


def _split_inline(text: str) -> list[str]:
    parts: list[str] = []
    buf: list[str] = []
    quote: str | None = None
    for ch in text:
        if quote:
            buf.append(ch)
            if ch == quote:
                quote = None
        elif ch in "\"'":
            quote = ch
            buf.append(ch)
        elif ch == ",":
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    if buf:
        parts.append("".join(buf))
    return [p.strip() for p in parts if p.strip() != ""]


def _scalar(text: str):
    s = text.strip()
    if s == "":
        return None
    if s.startswith("[") and s.endswith("]"):
        inner = s[1:-1].strip()
        return [] if not inner else [_scalar(x) for x in _split_inline(inner)]
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        return s[1:-1]
    low = s.lower()
    if low in ("true", "yes", "on"):
        return True
    if low in ("false", "no", "off"):
        return False
    if low in ("null", "~"):
        return None
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d+\.\d+", s):
        return float(s)
    return s


def _prepare(text: str) -> list[tuple[int, str]]:
    rows: list[tuple[int, str]] = []
    for raw in text.splitlines():
        line = _strip_comment(raw).rstrip()
        if not line.strip():
            continue
        stripped = line.lstrip(" ")
        indent = len(line) - len(stripped)
        if "\t" in line[:indent]:
            raise YamlError("不支持 Tab 缩进，请改用空格")
        rows.append((indent, stripped))
    return rows


def parse(text: str):
    rows = _prepare(text)
    if not rows:
        return {}
    value, idx = _parse_block(rows, 0, rows[0][0])
    if idx != len(rows):
        raise YamlError(
            f"第 {idx + 1} 个有效行缩进不一致：{rows[idx][1]!r}（indent={rows[idx][0]}）"
        )
    return value


def _parse_block(rows, idx: int, indent: int):
    if idx >= len(rows):
        return None, idx
    if rows[idx][1].startswith("- ") or rows[idx][1] == "-":
        return _parse_list(rows, idx, indent)
    return _parse_map(rows, idx, indent)


def _parse_map(rows, idx: int, indent: int):
    out: dict = {}
    while idx < len(rows) and rows[idx][0] == indent and not rows[idx][1].startswith("-"):
        content = rows[idx][1]
        if ":" not in content:
            raise YamlError(f"映射行缺少冒号：{content!r}")
        key, rest = content.split(":", 1)
        key = key.strip()
        rest = rest.strip()
        idx += 1
        if rest:
            out[key] = _scalar(rest)
        elif idx < len(rows) and rows[idx][0] > indent:
            out[key], idx = _parse_block(rows, idx, rows[idx][0])
        else:
            out[key] = None
    return out, idx


def _parse_list(rows, idx: int, indent: int):
    out: list = []
    while idx < len(rows) and rows[idx][0] == indent:
        head = rows[idx][1]
        if not (head.startswith("- ") or head == "-"):
            break
        content = head[2:].strip() if head.startswith("- ") else ""
        item_indent = indent + 2
        idx += 1
        if content == "":
            if idx < len(rows) and rows[idx][0] > indent:
                item, idx = _parse_block(rows, idx, rows[idx][0])
            else:
                item = None
            out.append(item)
            continue
        # 形如 `- key: value` 的映射项：把首行按 item_indent 复位后与后续缩进行合并
        looks_like_map = ":" in content and re.match(r"^[^:\s][^:]*:", content)
        if looks_like_map:
            sub = [(item_indent, content)]
            while idx < len(rows) and rows[idx][0] > indent:
                sub.append(rows[idx])
                idx += 1
            item, consumed = _parse_block(sub, 0, item_indent)
            if consumed != len(sub):
                raise YamlError(f"列表项内部缩进异常：{content!r}")
            out.append(item)
        else:
            out.append(_scalar(content))
    return out, idx