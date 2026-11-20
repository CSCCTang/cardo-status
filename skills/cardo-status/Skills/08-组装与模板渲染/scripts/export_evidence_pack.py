#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""export_evidence_pack.py -- 08 步导出《依据三件套》。

三件套 = 干净版 docx / 标注版 docx / 依据对照表 xlsx，同源生成：
  标注源（06 写出的〔锚点 …〕依据块）→ 标注版 → **剥离**出干净版，
  剥离前后数值集合与顺序必须一致（不一致即判不合格）。
唯一实现在 09/scripts/evidence_pack.py，本脚本只负责在流水线里调起来，
不复制任何校验逻辑（避免第二套实现漂移）。

用法：
    python export_evidence_pack.py --project-dir <项目目录> [--out-dir 成果文件/证据包]
退出码：0 导出完成并通过校验；1 已导出但校验不通过（09 步会拦）；2 无标注源（跳过）。
"""

import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "..",
                                               "09-三重自检与边界", "scripts")))

import evidence_pack  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", default=os.getcwd())
    ap.add_argument("--workspace", default="")
    ap.add_argument("--md", default="")
    ap.add_argument("--out-dir", default="")
    ap.add_argument("--version", default="v1")
    args, _unknown = ap.parse_known_args(argv)
    ws = args.workspace or args.project_dir
    out_dir = args.out_dir or os.path.join(ws, "成果文件/证据包")
    cmd = ["--project-dir", ws, "--out-dir", out_dir, "--version", args.version]
    if args.md:
        cmd += ["--md", args.md]
    ret = evidence_pack.main(cmd)
    if ret == 2:
        print("[08] 未发现标注源，三件套未导出（旧格式章节不纳入，不影响报告出稿）")
    elif ret == 1:
        print("[08] 三件套已导出，但依据校验不通过：详见 "
              "过程文件/09-三重自检与边界/依据校验_*.md（09 步会拦）")
    else:
        try:
            where = os.path.relpath(out_dir, ws)
        except ValueError:            # 跨盘符导出
            where = os.path.abspath(out_dir)
        print("[08] 三件套已导出：" + where)
    return ret


if __name__ == "__main__":
    raise SystemExit(main())
