#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_mapping_skeleton.py -- 生成 / 校验「模板章节 -> 原始报告来源」映射骨架。

用法:
    # 生成骨架
    python make_mapping_skeleton.py --template <模板.docx> --source <原始报告.docx> \
                                    --out <章节映射.yaml> [--max-level 2]
    # 校验映射（fail-closed）
    python make_mapping_skeleton.py --check <章节映射.yaml> \
                                    --template <模板.docx> --source <原始报告.docx>
    # 生成后立刻校验
    python make_mapping_skeleton.py --template ... --source ... --out ... --check ...

生成模式:
    * 从模板抽取 level<=--max-level（默认 1/2）的标题骨架，逐字保留；
    * 每个节点写 sources: [] 与 status: pending，待人/agent 填写；
    * 把原始报告的完整标题大纲附在 YAML 末尾的 source_outline 列表块里。

校验模式:
    * 每个模板节点必须 sources 非空或 status 属于 gap / 不涉及 / na / skip 等；
    * 节点缺失、多出、顺序错位，或 sources 写的标题在 source 大纲里找不到 -> 列问题并非零退出；
    * 通过时打印 `[mapping] ok N nodes`。

不依赖 pyyaml：写出的是扁平 YAML 子集，读取用内置的行解析器（环境若有 pyyaml 则优先用）。
退出码: 0 成功 / 校验通过; 1 校验未通过; 2 参数或输入错误。
"""

from __future__ import annotations

import argparse
import os
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

__version__ = "1.0.0"

HEADING_STYLE_RE = re.compile(r"^(?:heading|标题)\s*([1-9])$", re.IGNORECASE)

# 视为「已定义来源」的 status（其余（尤其 pending）都算未完成）
FILLED_STATUSES = {
    "done",
    "filled",
    "ok",
    "gap",
    "missing",
    "不涉及",
    "不适用",
    "无需",
    "无",
    "na",
    "n/a",
    "none",
    "skip",
    "skipped",
}
PENDING_STATUSES = {"pending", "todo", "待填", "待定", ""}

# 标题归一化时要剥掉的前导编号（"一、" "1." "1.2.3" "(1)" 等）
LEADING_NUMBER_RE = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百0-9]+[章节部分]|"
    r"[（(]\s*[0-9一二三四五六七八九十]+\s*[)）]|"
    r"[0-9]+(?:[.．][0-9]+)*[、.．)）]?|"
    r"[一二三四五六七八九十]+[、.．)）])\s*"
)


class MappingError(Exception):
    """参数 / 输入 / YAML 结构错误（退出码 2）。"""


# --------------------------------------------------------------------------- #
# 基础
# --------------------------------------------------------------------------- #
def _silence_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass


def iter_blocks(doc):
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield "paragraph", Paragraph(child, doc)


def para_style_name(p: Paragraph) -> str:
    try:
        name = p.style.name
    except Exception:
        name = ""
    return (name or "").strip()


def _outline_lvl_from_pPr(pPr) -> int | None:
    if pPr is None:
        return None
    el = pPr.find(qn("w:outlineLvl"))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val is None:
        return None
    try:
        iv = int(val)
    except (TypeError, ValueError):
        return None
    return None if iv >= 9 else iv + 1


def _style_outline_level(style) -> int | None:
    seen: set[int] = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        try:
            pPr = style.element.find(qn("w:pPr"))
        except Exception:
            pPr = None
        lvl = _outline_lvl_from_pPr(pPr)
        if lvl is not None:
            return lvl
        try:
            style = style.base_style
        except Exception:
            return None
    return None


def heading_level(p: Paragraph) -> int:
    m = HEADING_STYLE_RE.match(para_style_name(p))
    if m:
        return int(m.group(1))
    try:
        pPr = p._p.pPr
    except Exception:
        pPr = None
    lvl = _outline_lvl_from_pPr(pPr)
    if lvl is not None:
        return lvl
    try:
        return _style_outline_level(p.style) or 0
    except Exception:
        return 0


def extract_headings(docx_path: str, max_level: int | None = None) -> list[dict]:
    """按文档顺序抽取标题：{level, text}；max_level 为 None 表示不限。"""
    try:
        doc = Document(docx_path)
    except Exception as exc:  # noqa: BLE001
        raise MappingError(f"无法打开 docx {docx_path}: {type(exc).__name__}: {exc}") from exc
    out: list[dict] = []
    for _, p in iter_blocks(doc):
        lvl = heading_level(p)
        if not lvl:
            continue
        if max_level is not None and lvl > max_level:
            continue
        out.append({"level": lvl, "text": p.text.strip()})
    return out


def normalize_title(text: str) -> str:
    """标题比对口径：去前导编号、去空白、全角括号/冒号半角化后比较。"""
    s = (text or "").strip()
    s = LEADING_NUMBER_RE.sub("", s)
    s = s.replace("（", "(").replace("）", ")").replace("：", ":")
    s = re.sub(r"\s+", "", s)
    return s.lower()


# --------------------------------------------------------------------------- #
# 极简 YAML 读写（只覆盖本工具生成的子集，容忍常见手工编辑）
# --------------------------------------------------------------------------- #
def _yaml_quote(text: str) -> str:
    return '"' + str(text).replace("\\", "\\\\").replace('"', '\\"') + '"'


def _split_kv(text: str) -> tuple[str, str]:
    if ":" not in text:
        raise MappingError(f"YAML 行缺少 ':'：{text!r}")
    key, val = text.split(":", 1)
    return key.strip(), val.strip()


def _parse_scalar(val: str):
    val = val.strip()
    if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
        inner = val[1:-1]
        return inner.replace('\\"', '"').replace("\\\\", "\\")
    if val == "":
        return ""
    if val.startswith("[") and val.endswith("]"):
        inner = val[1:-1].strip()
        if not inner:
            return []
        items, buf, quote = [], "", None
        for ch in inner:
            if quote:
                buf += ch
                if ch == quote:
                    quote = None
            elif ch in "\"'":
                quote = ch
                buf += ch
            elif ch == ",":
                items.append(_parse_scalar(buf.strip()))
                buf = ""
            else:
                buf += ch
        if buf.strip():
            items.append(_parse_scalar(buf.strip()))
        return items
    return val


def parse_mapping_yaml(path: str) -> dict:
    """解析本工具生成的映射 YAML（nodes + source_outline + 顶层标量）。

    环境里若有 pyyaml 则优先用它整块解析；否则用下面的内置行解析器（零依赖）。
    """
    if not os.path.isfile(path):
        raise MappingError(f"--check 文件不存在: {path}")
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            text = f.read()
    except OSError as exc:
        raise MappingError(f"读取映射文件失败: {exc}") from exc
    lines = text.splitlines()

    # 有 pyyaml 时优先：整块解析成功且结构可用就采用，否则回退到内置解析器
    try:
        import yaml  # type: ignore

        loaded = yaml.safe_load(text)
        if isinstance(loaded, dict) and isinstance(loaded.get("nodes"), list):
            data = dict(loaded)
            data["nodes"] = [n for n in data["nodes"] if isinstance(n, dict)]
            outline = data.get("source_outline")
            if isinstance(outline, list):
                data["source_outline"] = [n for n in outline if isinstance(n, dict)]
            else:
                data["source_outline"] = []
            if data["nodes"]:
                return data
    except ImportError:
        pass
    except Exception:  # noqa: BLE001 - 任何 pyyaml 解析问题都回退到内置解析器
        pass

    root: dict = {"nodes": [], "source_outline": []}
    section = None
    cur_node: dict | None = None
    cur_item: dict | None = None
    pending_key: str | None = None

    for lineno, raw in enumerate(lines, 1):
        if not raw.strip():
            continue
        stripped = raw.strip()
        if stripped.startswith("#"):
            continue
        indent = len(raw) - len(raw.lstrip(" "))

        try:
            if indent == 0:
                if stripped.endswith(":"):
                    section = stripped[:-1].strip()
                    if section not in ("nodes", "source_outline"):
                        # 未知顶层块：整体跳过（容忍注释性扩展）
                        section = "__skip__"
                    cur_node = cur_item = None
                    pending_key = None
                    continue
                key, val = _split_kv(stripped)
                root[key] = _parse_scalar(val)
                section = "__skip__" if section in (None, "__skip__") else section
                continue

            if section == "nodes":
                if stripped.startswith("- ") and indent <= 2:
                    cur_node = {}
                    root["nodes"].append(cur_node)
                    body = stripped[2:].strip()
                    if body:
                        k, v = _split_kv(body)
                        cur_node[k] = _parse_scalar(v)
                    pending_key = None
                    continue
                if cur_node is None:
                    raise MappingError(f"第 {lineno} 行: nodes 块里出现游离内容")
                if stripped.startswith("- "):
                    val = _parse_scalar(stripped[2:].strip())
                    if pending_key is None:
                        raise MappingError(f"第 {lineno} 行: 列表项没有对应的键")
                    cur = cur_node.get(pending_key)
                    if not isinstance(cur, list):
                        cur = []
                        cur_node[pending_key] = cur
                    cur.append(val)
                    continue
                if ":" in stripped:
                    k, v = _split_kv(stripped)
                    if v == "":
                        cur_node[k] = []
                        pending_key = k
                    else:
                        cur_node[k] = _parse_scalar(v)
                        pending_key = None
                    continue
                raise MappingError(f"第 {lineno} 行: 无法解析节点字段 {stripped!r}")

            if section == "source_outline":
                if stripped.startswith("- ") and indent <= 2:
                    cur_item = {}
                    root["source_outline"].append(cur_item)
                    body = stripped[2:].strip()
                    if body:
                        k, v = _split_kv(body)
                        cur_item[k] = _parse_scalar(v)
                    continue
                if cur_item is None:
                    raise MappingError(f"第 {lineno} 行: source_outline 块里出现游离内容")
                if ":" in stripped:
                    k, v = _split_kv(stripped)
                    cur_item[k] = _parse_scalar(v)
                    continue
                raise MappingError(f"第 {lineno} 行: 无法解析 source_outline 字段 {stripped!r}")
            # __skip__ 或未知块：忽略
        except MappingError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise MappingError(f"第 {lineno} 行解析失败: {exc}") from exc

    return root


# --------------------------------------------------------------------------- #
# 生成
# --------------------------------------------------------------------------- #
def build_skeleton(template: str, source: str, max_level: int) -> tuple[str, dict]:
    tpl_nodes = extract_headings(template, max_level)
    if not tpl_nodes:
        raise MappingError(
            f"模板中未探测到 level<= {max_level} 的标题（样式名 / outlineLvl 均无），无法生成骨架"
        )
    src_nodes = extract_headings(source, None)
    if not src_nodes:
        raise MappingError("原始报告未探测到任何标题，无法提供 source_outline")

    width = max(2, len(str(len(tpl_nodes))))
    out: list[str] = []
    tpl_name = os.path.basename(template)
    src_name = os.path.basename(source)
    out.append(f"# 章节映射骨架 —— 由 make_mapping_skeleton.py {__version__} 生成")
    out.append(f"# 模板: {tpl_name}")
    out.append(f"# 原始报告: {src_name}")
    out.append("#")
    out.append("# 填法：")
    out.append("#   sources: 该模板节点对应的原始报告标题列表（逐字填写，可跨多个标题）")
    out.append("#   status:  pending=待填 / done=已填 / gap=原文缺失 / 不涉及=模板该节无需来源")
    out.append("# 校验（fail-closed）：")
    out.append("#   python make_mapping_skeleton.py --check <本文件> --template <模板.docx> --source <原始报告.docx>")
    out.append("# 合并顺序以 nodes 在本文档中的出现顺序为准。")
    out.append("")
    out.append("version: 1")
    out.append('tool: "make_mapping_skeleton.py"')
    out.append(f"template: {_yaml_quote(tpl_name)}")
    out.append(f"source: {_yaml_quote(src_name)}")
    out.append(f"max_level: {max_level}")
    out.append(f"node_count: {len(tpl_nodes)}")
    out.append("")
    out.append("nodes:")
    for i, node in enumerate(tpl_nodes, start=1):
        out.append(f"  - id: {_yaml_quote(f't{i:0{width}d}')}")
        out.append(f"    level: {node['level']}")
        out.append(f"    title: {_yaml_quote(node['text'])}")
        out.append("    sources: []")
        out.append("    status: pending")
    out.append("")
    out.append("# ---- 原始报告完整标题大纲（供填写 sources 参考，校验时也会使用）----")
    out.append("source_outline:")
    for node in src_nodes:
        out.append(f"  - level: {node['level']}")
        out.append(f"    title: {_yaml_quote(node['text'])}")
    out.append("")

    info = {
        "template_nodes": len(tpl_nodes),
        "template_level1": sum(1 for n in tpl_nodes if n["level"] == 1),
        "template_level2": sum(1 for n in tpl_nodes if n["level"] == 2),
        "source_nodes": len(src_nodes),
        "source_levels": sorted({n["level"] for n in src_nodes}),
    }
    return "\n".join(out), info


def write_skeleton(path: str, content: str) -> None:
    out_dir = os.path.dirname(os.path.abspath(path))
    os.makedirs(out_dir, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)


# --------------------------------------------------------------------------- #
# 校验
# --------------------------------------------------------------------------- #
def normalize_status(value: str) -> str:
    return re.sub(r"\s+", "", str(value or "")).strip().lower()


def check_mapping(mapping_path: str, template: str, source: str,
                  max_level_override: int | None = None) -> tuple[bool, list[str], dict]:
    data = parse_mapping_yaml(mapping_path)
    try:
        max_level = int(data.get("max_level", 2) or 2)
    except (TypeError, ValueError):
        raise MappingError(f"映射文件 max_level 不是整数: {data.get('max_level')!r}")
    if max_level_override is not None:
        max_level = max_level_override

    nodes = data.get("nodes") or []
    if not nodes:
        raise MappingError("映射文件里 nodes 为空（没有可校验的模板节点）")

    tpl_nodes = extract_headings(template, max_level)
    if not tpl_nodes:
        raise MappingError(f"模板中未探测到 level<= {max_level} 的标题，无法校验节点定义")

    src_outline = extract_headings(source, None)
    if not src_outline:
        # 模板/原始报告都读不到标题时退回到嵌入的 source_outline 并给出提示
        embedded = data.get("source_outline") or []
        if not embedded:
            raise MappingError("原始报告未探测到任何标题，且映射文件没有 source_outline，无法校验 sources")
        src_outline = [
            {"level": int(n.get("level", 0) or 0), "text": str(n.get("title", ""))}
            for n in embedded
        ]

    problems: list[str] = []

    # 1) 节点定义：与模板骨架按顺序比对，找缺失 / 多出 / 顺序错位
    yaml_seq: list[tuple[int, str, dict]] = []
    for idx, node in enumerate(nodes, start=1):
        try:
            lvl = int(node.get("level"))
        except (TypeError, ValueError):
            problems.append(f"节点 #{idx}: level 不是整数: {node.get('level')!r}")
            lvl = 0
        title = str(node.get("title", "") or "")
        nid = str(node.get("id", f"#{idx}") or f"#{idx}")
        yaml_seq.append((lvl, title, node))
        if not title.strip():
            problems.append(f"节点 {nid}: title 为空")

    tpl_seq = [(n["level"], n["text"]) for n in tpl_nodes]
    p = 0
    defined_ids: set[int] = set()
    for qi, (qlvl, qtitle, _node) in enumerate(yaml_seq):
        target = (qlvl, qtitle)
        found = None
        for k in range(p, len(tpl_seq)):
            if normalize_title(tpl_seq[k][1]) == normalize_title(qtitle) and tpl_seq[k][0] == qlvl:
                found = k
                break
        if found is None:
            problems.append(
                f"节点 {yaml_seq[qi][2].get('id', f'#{qi + 1}')}: 模板中不存在该节点 "
                f"[H{qlvl}] {qtitle!r}（模板已变更或标题写错）"
            )
        else:
            if found != p:
                for skip in range(p, found):
                    problems.append(
                        f"模板节点未定义: [H{tpl_seq[skip][0]}] {tpl_seq[skip][1]!r}（映射里缺失或顺序错位）"
                    )
            defined_ids.add(found)
            p = found + 1
    for skip in range(p, len(tpl_seq)):
        problems.append(
            f"模板节点未定义: [H{tpl_seq[skip][0]}] {tpl_seq[skip][1]!r}（映射里缺失或顺序错位）"
        )

    # 2) 每个节点的 sources / status
    src_index: dict[str, list[int]] = {}
    for i, n in enumerate(src_outline):
        src_index.setdefault(normalize_title(n["text"]), []).append(i)
    src_titles = [n["text"] for n in src_outline]
    src_norm_set = set(src_index)

    filled = 0
    gap = 0
    total_sources = 0
    for idx, (lvl, title, node) in enumerate(yaml_seq, start=1):
        nid = str(node.get("id", f"#{idx}") or f"#{idx}")
        raw_sources = node.get("sources", [])
        if isinstance(raw_sources, str):
            raw_sources = [raw_sources] if raw_sources.strip() else []
        sources = []
        if isinstance(raw_sources, list):
            for item in raw_sources:
                if isinstance(item, str) and item.strip():
                    sources.append(item.strip())
                elif isinstance(item, dict) and item.get("title"):
                    sources.append(str(item["title"]).strip())
        status = normalize_status(node.get("status", ""))
        total_sources += len(sources)

        if sources:
            if status in PENDING_STATUSES:
                print(
                    f"WARN: 节点 {nid} 已填 sources 但 status 仍为 {node.get('status')!r}",
                    file=sys.stderr,
                )
        if not sources and status not in FILLED_STATUSES:
            problems.append(
                f"节点 {nid} [H{lvl}] {title!r}: sources 为空且 status={node.get('status')!r}"
                f"（需填 sources 或把 status 设为 gap / 不涉及）"
            )
        if sources:
            filled += 1
            for s in sources:
                if normalize_title(s) not in src_norm_set:
                    # 给出最接近的候选，方便修正
                    near = _nearest_titles(normalize_title(s), src_titles, limit=3)
                    hint = f"；相近的 source 标题: {near}" if near else ""
                    problems.append(
                        f"节点 {nid}: 来源标题 {s!r} 在 source 大纲中找不到{hint}"
                    )
        elif status in FILLED_STATUSES:
            gap += 1

    stats = {
        "nodes": len(yaml_seq),
        "template_nodes": len(tpl_nodes),
        "filled": filled,
        "gap": gap,
        "sources": total_sources,
        "levels": {},
    }
    for lvl, _t, _n in yaml_seq:
        stats["levels"][lvl] = stats["levels"].get(lvl, 0) + 1
    return (not problems), problems, {"stats": stats, "max_level": max_level,
                                      "template": template, "source": source}


def _nearest_titles(target: str, titles: list[str], limit: int = 3) -> list[str]:
    if not target:
        return []
    scored = []
    for t in titles:
        nt = normalize_title(t)
        if not nt:
            continue
        if target in nt or nt in target:
            scored.append((abs(len(nt) - len(target)), t))
    scored.sort(key=lambda x: x[0])
    return [t for _, t in scored[:limit]]


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="make_mapping_skeleton.py",
        description="生成 / 校验「模板章节 -> 原始报告来源」映射骨架（不依赖 pyyaml）。",
    )
    ap.add_argument("--template", default=None, help="模板 .docx（生成 / 校验节点定义用）")
    ap.add_argument("--source", default=None, help="原始报告 .docx（source_outline 来源）")
    ap.add_argument("--out", default=None, help="生成模式：输出的映射 YAML 路径")
    ap.add_argument("--check", default=None, help="校验模式：待校验的映射 YAML 路径")
    ap.add_argument("--max-level", type=int, default=None,
                    help="模板骨架收录的标题层级上限（默认 2）")
    ap.add_argument("--version", action="version", version="%(prog)s " + __version__)
    return ap


def main(argv: list[str] | None = None) -> int:
    _silence_console()
    args = build_parser().parse_args(argv)

    if not args.out and not args.check:
        print("ERROR: 需要 --out（生成骨架）或 --check（校验映射）至少一个", file=sys.stderr)
        return 2
    for label, path in (("--template", args.template), ("--source", args.source)):
        if path is None:
            need = "生成与校验" if (args.out and args.check) else ("生成" if args.out else "校验")
            print(f"ERROR: {need}模式必须提供 {label}", file=sys.stderr)
            return 2
        if not os.path.isfile(path):
            print(f"ERROR: {label} 文件不存在: {path}", file=sys.stderr)
            return 2
        if not path.lower().endswith(".docx"):
            print(f"ERROR: {label} 必须是 .docx: {path}", file=sys.stderr)
            return 2

    max_level = args.max_level
    if max_level is not None and not (1 <= max_level <= 9):
        print("ERROR: --max-level 必须在 1..9 之间", file=sys.stderr)
        return 2

    try:
        if args.out:
            content, info = build_skeleton(
                args.template, args.source, 2 if max_level is None else max_level
            )
            write_skeleton(args.out, content)
            print(f"OK mapping skeleton {args.out}")
            print(
                f"  template nodes={info['template_nodes']} "
                f"(H1={info['template_level1']} H2={info['template_level2']})"
            )
            print(
                f"  source headings={info['source_nodes']} "
                f"levels={info['source_levels']}  (ends with source_outline block)"
            )
            print(f"  all {info['template_nodes']} nodes: sources=[] status=pending -> 待人/agent 填写")
            print(
                "  next: python make_mapping_skeleton.py --check <该文件> "
                "--template <模板.docx> --source <原始报告.docx>"
            )

        if args.check:
            ok, problems, meta = check_mapping(args.check, args.template, args.source, max_level)
            stats = meta["stats"]
            if ok:
                print(
                    f"[mapping] ok {stats['nodes']} nodes "
                    f"(filled={stats['filled']} gap={stats['gap']} sources={stats['sources']} "
                    f"max_level={meta['max_level']})"
                )
                return 0
            print(f"[mapping] FAIL {len(problems)} 个问题（fail-closed，未通过）")
            for i, msg in enumerate(problems, start=1):
                print(f"  {i}. {msg}")
            return 1
        return 0
    except MappingError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())