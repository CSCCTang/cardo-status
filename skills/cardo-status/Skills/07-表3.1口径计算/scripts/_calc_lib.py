#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_calc_lib.py -- 计算引擎底座（T23 v0；纯函数：口径校验后的指标计算/检查/轨迹元素）。

与 calc_metrics.py 的分工（同 T19 的 lib/CLI 拆分，控制单文件行数）：
  * 本模块：受限公式解析与求值、指标计算与校验、状态聚合、展示格式化；
  * calc_metrics.py：CLI、项目目录解析、口径/输入文件读取、轨迹落盘。

只用标准库；不做文件读写（口径/输入由调用方读好后传入），不写死盘符/项目名。
"""
from __future__ import annotations

import decimal
import os
import re
from pathlib import Path

__version__ = "0.1.0"
ENGINE_VERSION = __version__
SCHEMA_TRACE = "cardo/calc-trace/v1"
SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULT_RULES = SKILL_DIR / "规则" / "计算口径.yaml"

OK, GAP, NEEDS_CONFIRM, FAIL = "ok", "gap", "needs_confirm", "fail"
STATUS_ORDER = (OK, GAP, NEEDS_CONFIRM, FAIL)
CHECK_NAMES = {"unit_required", "divide_nonzero", "sum_row", "ratio_sum_100"}


class BadInput(Exception):
    """参数/输入/口径错误 → exit 2。"""


class CalcRunError(Exception):
    """计算/写入失败 → exit 1。"""


class DivideZero(Exception):
    """公式除零 → 该指标 needs_confirm（exit 3），不产出 0。"""


class Series(list):
    """逐项序列：元素为 (标签, Decimal)；cells 保留原始单元格供 inputs[] 引用。"""

    def __init__(self, pairs, cells):
        super().__init__(pairs)
        self.cells = cells


def disp(path, base=None) -> str:
    """产物/错误里的路径展示：skill 内 → <技能目录>/…；项目内 → <项目目录>/…；其它 → 文件名。"""
    p = Path(path)
    try:
        rp = p.resolve()
    except OSError:
        rp = Path(os.path.abspath(str(p)))
    pairs = [(SKILL_DIR, "<技能目录>")]
    if base:
        try:
            pairs.append((Path(base).resolve(), "<项目目录>"))
        except OSError:
            pass
    for root, prefix in pairs:
        try:
            rel = rp.relative_to(root).as_posix()
        except ValueError:
            continue
        return prefix if rel == "." else f"{prefix}/{rel}"
    return rp.name or "."


def norm_unit(value) -> str:
    return str(value or "").strip().replace("％", "%").replace("㎡", "平方米").replace("m²", "平方米")


def deci(value) -> decimal.Decimal:
    return decimal.Decimal(str(value))


def fmt(value) -> str:
    if value is None:
        return "—"
    if isinstance(value, float):
        text = f"{value:.6f}".rstrip("0").rstrip(".")
        return text or "0"
    return str(value)


def esc(text) -> str:
    return str(text).replace("|", "／").replace("\n", " ")


def cell_label(cell: dict) -> str:
    return str(cell.get("label") or "").strip() or str(cell.get("key") or "")


def cell_input(cell: dict) -> dict:
    return {"key": str(cell.get("key") or ""), "label": cell_label(cell),
            "value": cell.get("value"), "unit": cell.get("unit") or "",
            "ref": cell.get("ref") or {}}


# ---------------------------------------------------------------- 受限公式（白名单，禁任意代码）
TOKEN_RE = re.compile(r"\s*(?:(?P<num>\d+(?:\.\d+)?)|(?P<metric>CALC-\d{2})|"
                      r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)|(?P<op>[()+\-*/,]))")


def tokenize(formula: str) -> list[tuple[str, str]]:
    text = str(formula or "").strip()
    tokens, pos = [], 0
    while pos < len(text):
        match = TOKEN_RE.match(text, pos)
        if not match:
            raise BadInput(f"公式含非法字符：{formula!r}（位置 {pos}）")
        pos = match.end()
        for kind in ("num", "metric", "name", "op"):
            if match.group(kind) is not None:
                tokens.append((kind, match.group(kind)))
                break
    if not tokens:
        raise BadInput("formula 不能为空")
    return tokens


class Parser:
    """递归下降：expr(+ -) → term(* /) → factor(数/键/函数/括号/负号)。"""

    def __init__(self, tokens, formula):
        self.tokens, self.formula, self.i = tokens, formula, 0

    def peek(self):
        return self.tokens[self.i] if self.i < len(self.tokens) else (None, None)

    def take(self):
        token = self.peek()
        if token[0] is None:
            raise BadInput(f"公式不完整：{self.formula}")
        self.i += 1
        return token

    def parse(self):
        node = self.expr()
        if self.i != len(self.tokens):
            raise BadInput(f"公式有多余内容：{self.formula}")
        return node

    def expr(self):
        node = self.term()
        while self.peek() in (("op", "+"), ("op", "-")):
            node = (self.take()[1], node, self.term())
        return node

    def term(self):
        node = self.factor()
        while self.peek() in (("op", "*"), ("op", "/")):
            node = (self.take()[1], node, self.factor())
        return node

    def factor(self):
        kind, value = self.take()
        if kind == "num":
            return ("num", deci(value))
        if kind in ("metric", "name"):
            if self.peek() == ("op", "("):
                self.take()
                args = []
                if self.peek() != ("op", ")"):
                    args.append(self.expr())
                    while self.peek() == ("op", ","):
                        self.take()
                        args.append(self.expr())
                if self.take() != ("op", ")"):
                    raise BadInput(f"公式括号不匹配：{self.formula}")
                return ("call", value, args)
            return ("ref", value)
        if (kind, value) == ("op", "("):
            node = self.expr()
            if self.take() != ("op", ")"):
                raise BadInput(f"公式括号不匹配：{self.formula}")
            return node
        if (kind, value) == ("op", "-"):
            return ("neg", self.factor())
        raise BadInput(f"公式无法解析：{self.formula}")


def collect_refs(node) -> set[str]:
    kind = node[0]
    if kind == "ref":
        return {node[1]}
    if kind in ("num", "neg"):
        return set() if kind == "num" else collect_refs(node[1])
    if kind == "call":
        refs: set[str] = set()
        for arg in node[2]:
            refs |= collect_refs(arg)
        return refs
    return collect_refs(node[1]) | collect_refs(node[2])


class Ctx:
    """公式求值上下文：cell key → 序列/标量；已算指标 id → 结果。"""

    def __init__(self, cells: dict, results: dict):
        self.cells, self.results = cells, results

    def get(self, name: str):
        if name in self.results:
            return self.results[name]
        if name not in self.cells:
            raise BadInput(f"公式引用了无数据的键：{name}")
        cells = self.cells[name]
        if name.endswith("_i"):
            if not cells:
                raise BadInput(f"序列 {name} 没有 cell")
            return Series([(cell_label(c), deci(c["value"])) for c in cells], cells)
        if len(cells) != 1:
            raise BadInput(f"标量键 {name} 只允许一条 cell（当前 {len(cells)} 条）")
        return deci(cells[0]["value"])


def scalar(value) -> decimal.Decimal:
    if isinstance(value, Series):
        raise BadInput("公式此处需要标量，实际得到逐项序列（只有 sum/avg/ratio 接受序列）")
    return value


def call(name: str, args: list):
    if name in ("sum", "avg", "min", "max"):
        if len(args) != 1:
            raise BadInput(f"{name}() 需要 1 个参数")
        numbers = ([value for _label, value in args[0]] if isinstance(args[0], Series)
                   else [scalar(args[0])])
        if not numbers:
            raise BadInput(f"{name}() 序列为空")
        if name == "sum":
            return sum(numbers, decimal.Decimal(0))
        if name == "avg":
            return sum(numbers, decimal.Decimal(0)) / len(numbers)
        return min(numbers) if name == "min" else max(numbers)
    if name == "ratio":
        if len(args) != 2:
            raise BadInput("ratio() 需要 2 个参数")
        left, right = args
        if isinstance(left, Series):
            denominator = scalar(right)
            if denominator == 0:
                raise DivideZero("占比分母为 0")
            return Series([(label, value / denominator * 100) for label, value in left], left.cells)
        denominator = scalar(right)
        if denominator == 0:
            raise DivideZero("比值分母为 0")
        return scalar(left) / denominator * 100
    raise BadInput(f"未知公式函数：{name}（白名单：sum/avg/ratio/min/max）")


def eval_node(node, ctx: Ctx):
    kind = node[0]
    if kind == "num":
        return node[1]
    if kind == "ref":
        return ctx.get(node[1])
    if kind == "neg":
        return -scalar(eval_node(node[1], ctx))
    if kind == "call":
        return call(node[1], [eval_node(a, ctx) for a in node[2]])
    left, right = scalar(eval_node(node[1], ctx)), scalar(eval_node(node[2], ctx))
    if kind == "+":
        return left + right
    if kind == "-":
        return left - right
    if kind == "*":
        return left * right
    if right == 0:
        raise DivideZero("公式除数为 0")
    return left / right
