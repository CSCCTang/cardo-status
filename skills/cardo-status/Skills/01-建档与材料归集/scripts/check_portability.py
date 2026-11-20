#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_portability.py -- 打包/剪切技能前自检：技能目录里不许出现机器相关的绝对路径。

为什么需要：技能是要**整体剪切到别的项目/别的机器**复用的资产。一旦写死
`D:\\某个工作区\\...`、`C:\\Users\\谁\\python.exe`，换台机器就跑不起来；
写死具体项目名/模板文件名，换个项目就指错。

检查什么：
  1. 盘符绝对路径       `D:\\` `C:/`
  2. 用户/系统目录      `/Users/` `/home/` `AppData` `Program Files`
  3. 解释器绝对路径     `python.exe` `python3.exe`
  4. UNC 路径           `\\\\server\\share`
  5. 附加黑名单         `--deny 某某`（项目名、模板文件名等，按需传，可重复）

不检查什么（有意为之）：
  * URL（http/https 协议头里的 `://` 不算盘符）；
  * ZIP/XML 里的命名空间字符串（形如 `schemas.openxmlformats.org`）；
  * 代码里的相对定位写法 `Path(__file__).resolve().parent`（这正是要求写法）；
  * 业务正文里的中文路径描述（`<项目库根>\\<方向>\\<项目名>\\` 是占位符，不是绝对路径）。

用法：
    python check_portability.py                       # 扫本脚本所属技能目录
    python check_portability.py --siblings            # 再扫同级技能目录（skills 根下其它 skill）
    python check_portability.py --dir <目录> [...]    # 额外指定要扫的目录
    python check_portability.py --deny 某某 --deny 某某项目

退出码：0 干净；1 命中问题（不许分发）；2 参数错误。
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

__version__ = "1.0.0"

TEXT_SUFFIXES = {".md", ".py", ".yaml", ".yml", ".json", ".txt", ".csv", ".html", ".js"}
SKIP_DIRS = {"__pycache__", ".git", ".idea", ".vscode", "node_modules"}

# (规则名, 正则, 是否大小写敏感)
RULES: list[tuple[str, str, bool]] = [
    ("盘符绝对路径", r"[A-Za-z]:[\\/]", False),
    ("UNC 路径", r"\\\\[A-Za-z0-9._-]+[\\/]", False),
    ("用户主目录", r"/Users/|/home/|/root/", True),
    ("Windows 用户目录", r"AppData|Program Files|Users\\\\", False),
    ("解释器绝对路径", r"python\d?\.exe|python3?\s+[A-Za-z]:[\\/]", False),
]

# 明显不是盘符的上下文（协议头、XML 命名空间）
FALSE_POSITIVE = re.compile(
    r"https?://|ftp://|file://|schemas\.openxmlformats\.org|schemas\.microsoft\.com|"
    r"www\.w3\.org|purl\.org|xmlns",
    re.I,
)


def setup_stdout() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def iter_files(root: Path):
    self_path = Path(__file__).resolve()
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for name in filenames:
            path = Path(dirpath) / name
            # 本脚本自己必须包含这些模式（规则定义/示例），跳过自身，避免自证成假红
            if path.resolve() == self_path:
                continue
            if path.suffix.lower() in TEXT_SUFFIXES:
                yield path


def find_artifacts(root: Path) -> list[str]:
    """分发前该清掉的产物（不影响检查结论，只提示）。"""
    found: list[str] = []
    for dirpath, dirnames, filenames in os.walk(root):
        for d in list(dirnames):
            if d in ("__pycache__", ".pytest_cache", ".git"):
                found.append(str((Path(dirpath) / d).relative_to(root).as_posix()))
        for name in filenames:
            if name.endswith((".pyc", ".pyo", ".swp")) or name in (".DS_Store", "Thumbs.db"):
                found.append(str((Path(dirpath) / name).relative_to(root).as_posix()))
    return found


def scan_file(path: Path, deny: list[str]) -> list[tuple[int, str, str]]:
    hits: list[tuple[int, str, str]] = []
    try:
        text = path.read_text(encoding="utf-8-sig", errors="replace")
    except OSError:
        return hits
    for lineno, line in enumerate(text.splitlines(), start=1):
        # 行内忽略标记：文档里需要"举例说明不许写什么"时会引用路径样式，
        # 这类说明行加 `portability-ignore` 即可跳过（只跳过该行，不影响其余检查）。
        if "portability-ignore" in line:
            continue
        if FALSE_POSITIVE.search(line):
            # 去掉协议/命名空间片段后再判断，避免 URL 里的 "://" 被当盘符
            probe = FALSE_POSITIVE.sub("", line)
        else:
            probe = line
        for rule_name, pattern, case_sensitive in RULES:
            flags = 0 if case_sensitive else re.IGNORECASE
            if re.search(pattern, probe, flags):
                hits.append((lineno, rule_name, line.strip()[:160]))
                break
        else:
            for token in deny:
                if token and token in line:
                    hits.append((lineno, f"黑名单词「{token}」", line.strip()[:160]))
                    break
    return hits


def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    ap = argparse.ArgumentParser(description="技能可移植性自检（禁绝对路径/机器相关路径）")
    ap.add_argument("--dir", action="append", default=[], help="额外要扫描的目录，可重复")
    ap.add_argument("--siblings", action="store_true", help="同时扫描同级技能目录")
    ap.add_argument("--deny", action="append", default=[], help="额外黑名单词（项目名/模板名等），可重复")
    ap.add_argument("--quiet", action="store_true", help="只输出结论")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args(argv)

    skill_dir = Path(__file__).resolve().parent.parent
    targets: list[Path] = [skill_dir]
    if args.siblings:
        for sib in sorted(skill_dir.parent.iterdir()):
            if sib.is_dir() and sib != skill_dir and (sib / "SKILL.md").is_file():
                targets.append(sib)
    for extra in args.dir:
        p = Path(extra).expanduser()
        if not p.is_dir():
            print(f"错误：目录不存在：{extra}", file=sys.stderr)
            return 2
        targets.append(p.resolve())

    # 去重（按解析后的路径）
    seen: set[Path] = set()
    uniq: list[Path] = []
    for t in targets:
        rt = t.resolve()
        if rt not in seen:
            seen.add(rt)
            uniq.append(rt)

    total_files = 0
    total_hits = 0
    total_artifacts = 0
    print(f"[可移植性] 扫描 {len(uniq)} 个技能目录")
    for root in uniq:
        hits_by_file: dict[Path, list] = {}
        files = 0
        for path in iter_files(root):
            files += 1
            found = scan_file(path, args.deny)
            if found:
                hits_by_file[path] = found
        total_files += files
        n_hits = sum(len(v) for v in hits_by_file.values())
        total_hits += n_hits
        artifacts = find_artifacts(root)
        total_artifacts += len(artifacts)
        tag = "OK" if not hits_by_file else f"FAIL（{n_hits} 处）"
        print(f"  - {root.name}: 文件 {files} → {tag}")
        if artifacts and not args.quiet:
            for item in artifacts[:8]:
                print(f"      [建议删除] {item}")
            if len(artifacts) > 8:
                print(f"      [建议删除] … 另有 {len(artifacts) - 8} 项")
        if not args.quiet:
            for path, found in hits_by_file.items():
                rel = path.relative_to(root).as_posix()
                for lineno, rule_name, line in found:
                    print(f"      {rel}:{lineno}  [{rule_name}]  {line}")

    print("")
    if total_hits:
        print(f"结论：❌ 有 {total_hits} 处机器相关路径/词，**不要打包分发**。"
              "改成：脚本内用 Path(__file__) 定位、文档里用 <占位符>、路径由参数/配置注入。")
        return 1
    if total_artifacts:
        print(f"结论：✅ 路径干净（{total_files} 个文件，0 处机器相关路径）；"
              f"但发现 {total_artifacts} 个运行产物（__pycache__/.pyc 等），**分发前请删除**。")
        return 0
    print(f"结论：✅ 干净（{total_files} 个文件，0 处机器相关路径）。可以整体剪切到其它项目/机器。")
    return 0


if __name__ == "__main__":
    sys.exit(main())