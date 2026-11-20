#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第 09 步边界检查：只读原件 / 模板指纹 / 技能可移植 / 交付物齐套。

任一失败 exit 1；本脚本不修改任何原件。
"""
import argparse
import hashlib
import json
import os
import re
import sys

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None

ABS_PATH_RE = re.compile(r"(?<![A-Za-z0-9+.:/])[A-Za-z]:[\\/]|(?<![\w.])/(?:Users|home|mnt)/")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--fingerprint", required=True, help="第 01 步保存的原件指纹 JSON")
    ap.add_argument("--template", required=True)
    ap.add_argument("--contract", required=True)
    ap.add_argument("--deliverables", nargs="+", default=[])
    ap.add_argument("--inputs", default="", help="输入目录（只允许白名单扩展名）")
    ap.add_argument("--outputs", default="", help="成果目录（只允许 Office 扩展名）")
    ap.add_argument("--report", required=True)
    ap.add_argument("--gap-mode", action="store_true")
    ap.add_argument("--out", required=True)
    args = ap.parse_args(argv)

    # 输入目录允许的扩展名：办公文档 + 图纸及其伴随文件 + 压缩包 + 文本。
    # CAD 伴随文件（shx/fas/dxf/dgn/rvt）、压缩包（rar/7z）、日志（log）等
    # 都是真实投递材料，不应被判成「非白名单文件」——否则图纸类项目 09 必然失败。
    INPUT_EXTS = {".docx", ".doc", ".pdf", ".xlsx", ".xls", ".pptx", ".ppt",
                  ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp",
                  ".dwg", ".dxf", ".shx", ".fas", ".dgn", ".rvt", ".mdb",
                  ".json", ".csv", ".txt", ".md", ".log",
                  ".zip", ".rar", ".7z",
                  # 三维模型瓦片与配套元数据：低空课题的核心交付数据（osgb/ive 等），
                  # 由 01 步解包 zip 后落在 输入文件/解包-*/ 下。
                  # 不放行会让 09 把「流水线自己解出来的材料」判成野文件。
                  ".osgb", ".ive", ".xml",
                  # 三维模型成果的网页浏览包：3D Tiles 瓦片（b3dm/glb/pnts 等）、
                  # 矢量轮廓（geojson）、查看器脚本与样式（js/css/html/wasm/svg）
                  # 及其本地服务（bat/exe）。它们是上游成果自带的查看器，属投递内容，
                  # 不参与成文，也不被执行——只用于判「输入目录有没有野文件」。
                  ".b3dm", ".glb", ".gltf", ".pnts", ".i3dm", ".cmpt",
                  ".geojson", ".js", ".cjs", ".mjs", ".css", ".html", ".htm", ".wasm", ".svg",
                  ".py", ".bat", ".exe",
                  # 网盘未下载完的临时残片（*.xxx.baiduyun.p.downloading）：
                  # 是用户真实投递的内容，不该判成野文件。
                  ".downloading"}
    OUTPUT_EXTS = {".docx", ".xlsx", ".pdf"}

    fails, warns = [], []
    fp = json.load(open(args.fingerprint, encoding="utf-8"))
    for rel, meta in (fp.get("files") or {}).items():
        p = os.path.join(args.project_dir, rel)
        if not os.path.exists(p):
            fails.append(f"原件被删除：{rel}")
        elif sha256(p) != meta["sha256"]:
            fails.append(f"原件被改动：{rel}")

    contract = yaml.safe_load(open(args.contract, encoding="utf-8")) or {}
    exp = (contract.get("template") or {}).get("sha256")
    got = sha256(args.template)
    if exp != got:
        fails.append("模板 sha256 与章节契约不一致（可能用了旧模板）")

    skills = os.path.join(args.project_dir, "Skills")
    hits = []
    for root, _, files in os.walk(skills):
        for fn in files:
            if not fn.endswith(".py"):
                continue
            p = os.path.join(root, fn)
            try:
                txt = open(p, encoding="utf-8").read()
            except UnicodeDecodeError:
                continue
            for ln in txt.splitlines():
                s = ln.strip()
                # 文档示例 / URL / 注释不算硬编码路径
                if s.startswith("#") or "http" in s or "`" in s or 'r"' in s or "r'" in s:
                    continue
                if ABS_PATH_RE.search(s):
                    hits.append(f"{os.path.relpath(p, args.project_dir)}: {s[:80]}")
                    break
    if hits:
        fails.append(f"技能目录存在绝对路径 {len(hits)} 处（可移植性）")
        warns.extend(hits[:5])

    for d in args.deliverables:
        if not os.path.exists(d):
            fails.append(f"交付物缺失：{os.path.relpath(d, args.project_dir)}")
        elif os.path.splitext(d)[1].lower() not in OUTPUT_EXTS:
            fails.append(f"交付物扩展名不在白名单（{'/'.join(sorted(OUTPUT_EXTS))}）："
                         f"{os.path.relpath(d, args.project_dir)}")

    # 输入目录扩展名白名单 / 成果目录只放 Office
    if args.inputs and os.path.isdir(args.inputs):
        for root, _, files in os.walk(args.inputs):
            for fn in files:
                ext = os.path.splitext(fn)[1].lower()
                if ext not in INPUT_EXTS:
                    fails.append(f"输入目录出现非白名单文件：{os.path.relpath(os.path.join(root, fn), args.project_dir)}")
    if args.outputs and os.path.isdir(args.outputs):
        for root, _, files in os.walk(args.outputs):
            for fn in files:
                ext = os.path.splitext(fn)[1].lower()
                if ext not in OUTPUT_EXTS:
                    fails.append(f"成果目录出现非 Office 文件：{os.path.relpath(os.path.join(root, fn), args.project_dir)}")

    report = open(args.report, encoding="utf-8").read()
    if args.gap_mode and "缺口版" not in report:
        fails.append("缺口版成稿未在正文标注「缺口版」")
    if not args.gap_mode and "缺口版" in report:
        warns.append("完整版正文出现「缺口版」字样，请确认是否有意")

    out = {"schema": "cardo/boundary/v1",
           "ok": not fails,
           "gapMode": bool(args.gap_mode),
           "fails": fails,
           "warns": warns,
           "counts": {"originals": len(fp.get("files") or {}),
                      "deliverables": len(args.deliverables)}}
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as w:
        json.dump(out, w, ensure_ascii=False, indent=1)
    print(f"[09] 边界检查：原件 {len(fp.get('files') or {})} 份只读，"
          f"交付物 {len(args.deliverables)} 项；{'PASS' if not fails else 'FAIL'}")
    for f in fails:
        print(f"[09][FAIL] {f}")
    for w_ in warns:
        print(f"[09][WARN] {w_}")
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
