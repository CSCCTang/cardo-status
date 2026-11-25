# -*- coding: utf-8 -*-
"""资产一致性校验：同一份 assets 在仓库各处必须一致。

插件仓库内有两条资产路径：
- `assets/`（仓库根，便于直接查看）
- `skills/cardo-status/assets/`（运行时实际读取的那份）

默认行为是自检这两处；也可以由外部发布脚本通过 `compare()` 传入更多目标目录
（例如主产品、安装后的插件目录），保证“改一处”时立刻暴露分叉。

规则
- 普通文件逐字节比对（CRLF 归一为 LF，避免换行被当成差异）。
- .docx 以“除 docProps/ 外逐部件哈希”比对；只有元数据或压缩方式不同时只提示、不算分叉。
- 任一处缺文件或真实内容不同，退出码为 1。

用法：
  python scripts/check_assets_sync.py [--quiet]
"""
from __future__ import annotations

import hashlib
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent          # <仓库根>/scripts
ROOT = HERE.parent                              # 仓库根
SOURCE = ROOT / "skills" / "cardo-status" / "assets"
DEFAULT_TARGETS = {"仓库根": ROOT / "assets"}


def _files(root: Path) -> set[str]:
    return {p.relative_to(root).as_posix() for p in root.rglob("*") if p.is_file()}


def _digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes().replace(b"\r\n", b"\n")).hexdigest()


def _docx_digest(path: Path, skip_meta: bool = True) -> str:
    with zipfile.ZipFile(path) as z:
        parts = [n for n in sorted(z.namelist())
                 if not (skip_meta and n.startswith("docProps/"))]
        h = hashlib.sha256()
        for n in parts:
            h.update(n.encode("utf-8"))
            h.update(z.read(n))
    return h.hexdigest()


def _same(source: Path, target: Path) -> tuple[bool, str]:
    if source.suffix.lower() == ".docx" and target.suffix.lower() == ".docx":
        if _docx_digest(source) != _docx_digest(target):
            return False, "正文内容不同"
        if source.read_bytes() != target.read_bytes():
            return True, "仅压缩/元数据差异"
        return True, "一致"
    if _digest(source) != _digest(target):
        return False, "内容不同"
    if source.read_bytes() != target.read_bytes():
        return True, "仅换行差异"
    return True, "一致"


def compare(source: Path, targets: dict[str, Path], quiet: bool = False) -> int:
    """以 source 为基准比对 targets；返回 0 表示一致（允许仅元数据差异）。"""
    if not source.is_dir():
        print("资产校验失败：源目录不存在 %s" % source)
        return 1
    src_files = _files(source)
    problems: list[str] = []
    notes: list[str] = []
    for name, root in targets.items():
        if not root.is_dir():
            notes.append("%s 无 assets 目录（%s），跳过" % (name, root))
            continue
        cur = _files(root)
        problems += ["%s 缺少 %s" % (name, rel) for rel in sorted(src_files - cur)]
        problems += ["%s 多出未登记资产 %s" % (name, rel) for rel in sorted(cur - src_files)]
        for rel in sorted(src_files & cur):
            ok, why = _same(source / rel, root / rel)
            if not ok:
                problems.append("%s 的 %s：%s" % (name, rel, why))
            elif why.startswith("仅"):
                notes.append("%s 的 %s：%s（不计为分叉）" % (name, rel, why))
    if not quiet:
        print("资产源：%s（%d 个文件）" % (source, len(src_files)))
        for n in notes:
            print("  提示：%s" % n)
    if problems:
        print("资产不一致 %d 处：" % len(problems))
        for p in problems:
            print("  [FAIL] %s" % p)
        return 1
    if not quiet:
        print("资产一致：%s 共 %d 个资产文件，0 处分叉。" % (" / ".join(targets), len(src_files)))
    return 0


if __name__ == "__main__":
    sys.exit(compare(SOURCE, DEFAULT_TARGETS, "--quiet" in sys.argv[1:]))
