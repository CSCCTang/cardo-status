#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""precheck_drafts.py -- 分章回填稿的**格式预检**（在 06 单章自检之前跑）。

为什么要有：09 步的契约检查（CHK-GAP-1）会判「附件四有行但正文未引用」，
但那已经跑到最后一步了，用户看到的是一条指向附件的失败，
真正的原因（内联标记写成「见缺项-D01」，缺了方括号和冒号）却没被指出。

这里在 06 步就把它挑出来，直接说「哪个文件、哪一行、该写成什么样」。

检查项（只查格式，不查内容、不改文件）：
  ① 缺项内联标记格式：必须是 【缺项-XNN：…】——方括号内、冒号在括号内。
     写成「缺项-D01」「见缺项-D01」「【缺项-D01】」都判不合格。
  ② 缺项编号的章字母：正文里引用他章的编号要显式说明，
     否则 09 步会判「编号与所属章节不符」。
  ③ 标题层级：只允许 # 与 ##（模板最多二级）。
  ④ 禁句：命中即不合格。
  ⑤ 来源行：每节末尾恰有一条「（数据来源：…）」。
  ⑥ 三级及以上标题、缺项编号格式错误。

退出码：0 全部合格；1 有格式不合格项（逐条报出）；2 参数/输入错误；
        3 脚本自身异常（崩溃/环境问题）——**不是**格式判定，总控须分开报。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path

# 规范格式：方括号内、冒号在括号内
GAP_INLINE_OK = re.compile(r"【缺项-([A-F])(\d{2})：[^】]*】")
# 不合格的写法（用于给出针对性提示）
GAP_INLINE_BAD = re.compile(r"(?:见\s*)?缺项-([A-F])(\d{2})(?!：)")
GAP_TABLE_ROW = re.compile(r"^\|\s*缺项-([A-F])(\d{2})\s*\|")
HEAD_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*$")
# 来源行：整行是「（数据来源：…）」。
# 注意不能写成 `[^）]+` —— 来源文件名里可能带中文括号
# （如「…统计表（某社区）.xlsx」），第一个「）」就会截断匹配。
# 故用「最后一个右括号收尾」：非贪婪前缀 + 贪婪到行尾。
SOURCE_RE = re.compile(r"^\s*（数据来源：.+）\s*$")

# 运行时是 python313._pth 模式，脚本目录不进 sys.path：同目录兄弟模块必须显式加进来
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from _banned import TODO_PHRASES  # noqa: E402
except Exception as exc:  # 模块级导入失败发生在 main 之外，必须在这里就以 exit 3 退出
    import traceback
    traceback.print_exc()
    print(f"[预检][内部错误] 无法导入同目录模块 _banned：{type(exc).__name__}: {exc}", file=sys.stderr)
    print("  这是脚本执行错误（打包缺失/环境问题），不是回填稿格式判定。", file=sys.stderr)
    raise SystemExit(3) from None

BANNED = list(TODO_PHRASES)
# 写作任务包模板自带的提示注释：带这行的文件**还没回填**，不是成稿，预检跳过
TEMPLATE_MARK = "<!-- 请由大模型成文后写入本文件"


def chapter_letter(name: str) -> str:
    m = re.match(r"^([A-F])[_\-]", name)
    return m.group(1) if m else ""


def check_file(path: Path) -> list[dict]:
    """返回该文件的问题清单 [{line, id, msg}]。"""
    problems: list[dict] = []
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        return [{"line": 0, "id": "PRE-READ", "msg": f"读不到文件：{exc}"}]
    lines = text.splitlines()
    letter = chapter_letter(path.name)

    in_gap_table = False
    for no, ln in enumerate(lines, start=1):
        s = ln.strip()
        if not s:
            continue

        # 标题层级：只允许 # 与 ##
        m = HEAD_RE.match(s)
        if m and len(m.group(1)) > 2:
            problems.append({"line": no, "id": "PRE-HEAD",
                             "msg": f"出现三级及以上标题「{m.group(2)}」——模板最多二级，"
                                    f"节标题用 ##"})

        # 禁句
        for b in BANNED:
            if b in ln:
                problems.append({"line": no, "id": "PRE-BAN",
                                 "msg": f"命中禁句「{b}」"})

        # 缺项表区间：表格行里的裸编号是**正确**的（表头列就是编号）
        if s.startswith("## ") and "缺项表" in s:
            in_gap_table = True
            continue
        if in_gap_table and s.startswith("## "):
            in_gap_table = False
        if in_gap_table:
            continue

        # 缺项内联标记：正文里的编号必须是 【缺项-XNN：…】
        if GAP_TABLE_ROW.match(s):
            continue
        for mm in GAP_INLINE_BAD.finditer(ln):
            # 排除已经被合规标记覆盖的（合规标记里也有缺项-XNN）
            good = GAP_INLINE_OK.search(ln)
            if good and good.group(0) in ln:
                continue
            frag = mm.group(0)
            problems.append({
                "line": no, "id": "PRE-GAP",
                "msg": (f"缺项内联标记「{frag}」格式不对：必须写成"
                        f"【缺项-{mm.group(1)}{mm.group(2)}：见本节缺项表】"
                        f"（方括号内、冒号在括号内）")})

        # 章字母：引用他章编号要显式说明
        if letter:
            for mm in GAP_INLINE_OK.finditer(ln):
                if mm.group(1) != letter and "见第" not in ln and "详见" not in ln:
                    problems.append({
                        "line": no, "id": "PRE-GAP-X",
                        "msg": (f"引用了别的章的缺项编号 缺项-{mm.group(1)}{mm.group(2)}，"
                                f"本章是 {letter}——09 步会判「编号与所属章节不符」。"
                                f"请改成文字说明（如「详见第三章缺项登记」）"
                                f"或登记本章自己的编号")})

    # 来源行：每节末尾恰有一条。
    # 「本节缺项表」「来源映射」是**过程段**（由 08 组装时剥离），不要求来源行。
    PROCESS_SECTIONS = ("本节缺项表", "来源映射")
    cur_title, cur_start, srcs = "", 0, []
    for no, ln in enumerate(lines, start=1):
        s = ln.strip()
        if s.startswith("## "):
            if cur_title and not any(p in cur_title for p in PROCESS_SECTIONS):
                _check_section(problems, cur_title, cur_start, srcs, path)
            cur_title, cur_start, srcs = s[3:], no, []
        elif SOURCE_RE.match(s):
            srcs.append(no)
    if cur_title and not any(p in cur_title for p in PROCESS_SECTIONS):
        _check_section(problems, cur_title, cur_start, srcs, path)

    return problems


def _check_section(problems, title, start, srcs, path):
    if not srcs:
        problems.append({"line": start, "id": "PRE-SRC",
                         "msg": f"节「{title}」末尾缺来源行"
                                f"（格式：（数据来源：项目资料，来源：<来源文件名>））"})
    elif len(srcs) > 1:
        problems.append({"line": srcs[1], "id": "PRE-SRC",
                         "msg": f"节「{title}」有 {len(srcs)} 条来源行，建议只留 1 条"})


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="分章回填稿格式预检（缺项标记/标题层级/禁句/来源行）")
    ap.add_argument("--draft-dir", required=True, help="分章初稿目录")
    ap.add_argument("--out", default="", help="预检报告 JSON 输出路径")
    args = ap.parse_args(argv)

    draft_dir = Path(args.draft_dir)
    if not draft_dir.is_dir():
        print(f"[预检] 目录不存在：{draft_dir}", file=sys.stderr)
        return 2

    files = sorted(p for p in draft_dir.glob("*.md") if not p.name.startswith("_"))
    if not files:
        print(f"[预检] {draft_dir} 里没有 .md 文件（尚未回填，跳过预检）")
        return 0

    all_problems = []
    checked, skipped = 0, 0
    for p in files:
        try:
            raw = p.read_text(encoding="utf-8")
        except OSError:
            raw = ""
        # 只检「回填后的成稿」：仍是模板的文件（自带提示注释/禁句/占位来源行）不检
        if TEMPLATE_MARK in raw:
            skipped += 1
            continue
        checked += 1
        for pr in check_file(p):
            all_problems.append({"file": p.name, **pr})

    payload = {
        "schema": "cardo/draft-precheck/v1",
        "draftDir": str(draft_dir),
        "checked": checked,
        "skippedTemplates": skipped,
        "problems": all_problems,
        "ok": not all_problems,
    }
    if args.out:
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.out).write_text(json.dumps(payload, ensure_ascii=False, indent=1),
                                  encoding="utf-8")

    if not all_problems:
        print(f"[预检] {checked} 章格式合格（缺项标记/标题层级/禁句/来源行）"
              + (f"；{skipped} 章仍是模板，未检" if skipped else ""))
        return 0

    print(f"[预检] {checked} 章发现 {len(all_problems)} 处格式问题：")
    for pr in all_problems[:20]:
        print(f"   {pr['file']}:{pr['line']}  {pr['id']}  {pr['msg']}")
    if len(all_problems) > 20:
        print(f"   …另有 {len(all_problems) - 20} 处")
    print("   改完再跑 06；这些问题不修，09 步会以「附件四有行但正文未引用」的形式报出来。")
    return 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception as exc:  # 崩溃与业务不合格必须分开出口，不能混成「格式不合格」
        import traceback
        traceback.print_exc()
        print(f"[预检][内部错误] {type(exc).__name__}: {exc}", file=sys.stderr)
        print("  这是脚本执行错误（环境/代码问题），不是回填稿格式判定；"
              "请按上面的堆栈定位，不要把 exit 3 当成格式不合格处理。", file=sys.stderr)
        sys.exit(3)
