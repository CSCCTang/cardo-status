#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_evidence.py -- 09 步依据校验（只校验、不写产物）。

校验对象是 08 已导出的标注源与三件套：段落覆盖 / 数值有源 / 依据有效 /
剥离一致 / 正文非表格化 / 对照表回指 / 链条闭合。同一套实现（evidence_pack.py），
--check-only 只读不写，避免 08 与 09 各写一份 docx 互相覆盖。

用法：
    python check_evidence.py --project-dir <项目目录>
退出码：0 通过；1 校验不通过或无标注源（正文已不带来源行，没有〔锚点〕即断链）。
"""

import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import evidence_pack  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", default=os.getcwd())
    ap.add_argument("--workspace", default="")
    ap.add_argument("--md", default="")
    ap.add_argument("--version", default="v1")
    args, _unknown = ap.parse_known_args(argv)
    ws = args.workspace or args.project_dir
    cmd = ["--project-dir", ws, "--version", args.version, "--check-only"]
    if args.md:
        cmd += ["--md", args.md]
    try:
        ret = evidence_pack.main(cmd)
    except Exception as exc:  # noqa: BLE001
        # 脚本自身异常与业务不合格分开：总控据此返回 3，不混同 1。
        print(f"[09] 校验脚本自身异常：{type(exc).__name__}: {exc}")
        return 3
    if ret == 2:
        # 设计好的拒收：正文已不写来源行、又没有〔锚点〕，追溯链不可能建立。
        # 与「业务不合格（1）」「脚本异常（3）」分开，便于总控区分处置。
        print("[09] 未找到标注源（06 分章初稿无〔锚点〕依据块）：正文已不写来源行，"
              "追溯链无法建立；按标注源口径重跑 06 后再交付")
        return 4
    return 0 if ret == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
