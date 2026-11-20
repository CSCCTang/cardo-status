#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第 01 步：分类确认合并 + 项目索引生成。

分类清单由 classify_materials.py 产出。口径（2026-11-14 起）：
  * **低置信度不阻断**：判不准的文件已按最接近类别入档，本脚本把它们汇总成
    `低置信度归类.md` / `数据缺口登记.json`，流水线继续往下走；
  * **只拦真缺口**：仅当文件完全读不出来（空文件 / 损坏 / 加密）时才提示补齐，
    但仍不阻断——缺口登记后照常出稿（缺口另出清单）；
  * `分类确认.yaml` 仍是可选的人工校正入口：写了 decisions 就按人判覆盖自动结果。

另生成 过程文件/00-总控/项目索引.md（项目、投递区、模板、源报告、原件指纹）。
"""
import argparse
import hashlib
import json
import os
import sys

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_categories(rules_path):
    rules = yaml.safe_load(open(rules_path, encoding="utf-8")) or {}
    return {c["id"]: c for c in rules.get("categories", [])}, rules.get("dirs", {})


def write_md(manifest, path):
    s = manifest.get("summary", {})
    lines = [f"# 分类清单（{manifest.get('project','')}）", "",
             f"- 投递区：{manifest.get('delivery','')}",
             f"- 文件总数：{s.get('total', 0)}（自动 {s.get('auto', 0)} / 待确认 {s.get('pending', 0)}）",
             f"- 生成时间：{manifest.get('generated_at','')}", "",
             "| # | 文件 | 大小 | 类别 | 置信度 | 状态 | 去向 |",
             "| --- | --- | --- | --- | --- | --- | --- |"]
    for i, f in enumerate(manifest.get("files", []), 1):
        lines.append(f"| {i} | {f.get('rel','')} | {f.get('size_h','')} | "
                     f"{f.get('category_label','')} | {f.get('confidence',0):.2f} | "
                     f"{'待确认' if f.get('pending') else '已确认'} | {f.get('dest_rel','')} |")
    with open(path, "w", encoding="utf-8") as w:
        w.write("\n".join(lines) + "\n")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--project-dir", required=True)
    ap.add_argument("--manifest", required=True)
    ap.add_argument("--confirm", required=True, help="分类确认.yaml")
    ap.add_argument("--rules", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--source", default="",
                    help="成篇材料路径；散材料模式可不传，由 03 自动汇编")
    ap.add_argument("--project", required=True)
    ap.add_argument("--direction", required=True)
    ap.add_argument("--delivery", default="输入文件")
    ap.add_argument("--index-out", required=True)
    ap.add_argument("--fingerprint", default="", help="原件指纹输出 JSON（边界检查用）")
    args = ap.parse_args(argv)

    manifest = json.load(open(args.manifest, encoding="utf-8"))
    cats, dirs = load_categories(args.rules)
    decisions = {}
    if os.path.exists(args.confirm):
        data = yaml.safe_load(open(args.confirm, encoding="utf-8")) or {}
        for d in data.get("decisions", []) or []:
            if d.get("file"):
                decisions[d["file"]] = d

    for f in manifest.get("files", []):
        if not f.get("pending"):
            continue
        d = decisions.get(f.get("rel")) or decisions.get(os.path.basename(f.get("rel", "")))
        if not d:
            continue
        cid = d.get("category_id")
        if cid not in cats:
            print(f"[01][FAIL] 确认文件里 {f.get('rel')} 的 category_id 非法：{cid}")
            return 1
        cat = cats[cid]
        f["category_id"] = cid
        f["category_label"] = cat.get("label")
        f["confidence"] = 1.0
        f["gap"] = 1.0
        f["pending"] = False
        f["pending_reason"] = ""
        f["confirmedByUser"] = True
        f["dest_rel"] = os.path.join(cat.get("dir", ""), os.path.basename(f["rel"])).replace("\\", "/")
        manifest.setdefault("decisions_applied", []).append(
            {"file": f["rel"], "category_id": cid, "note": d.get("note", "")})

    auto = sum(1 for f in manifest.get("files", []) if not f.get("pending"))
    pending = [f["rel"] for f in manifest.get("files", []) if f.get("pending")]
    low = [f for f in manifest.get("files", []) if f.get("low_confidence")]
    unreadable = [f for f in manifest.get("files", []) if f.get("unreadable")]
    manifest.setdefault("summary", {})
    manifest["summary"]["auto"] = auto
    manifest["summary"]["pending"] = len(pending)
    manifest["summary"]["low_confidence"] = len(low)
    manifest["summary"]["unreadable"] = len(unreadable)
    manifest["summary"]["confirmed"] = len(manifest.get("decisions_applied", []))
    with open(args.manifest, "w", encoding="utf-8") as w:
        json.dump(manifest, w, ensure_ascii=False, indent=1)
    write_md(manifest, os.path.splitext(args.manifest)[0] + ".md")

    # 低置信度与真缺口登记：写给人看的一份，也供 02/06 缺口清单引用。
    reg_dir = os.path.dirname(os.path.abspath(args.manifest))
    reg = {"schema": "cardo/data-gap/v1",
           "project": args.project,
           "lowConfidence": [
               {"file": f.get("rel"), "category": f.get("category_label"),
                "confidence": f.get("confidence"), "reason": f.get("low_confidence_reason"),
                "dest": f.get("dest_rel")}
               for f in low],
           "unreadable": [
               {"file": f.get("rel"), "reason": f.get("unreadable_reason"),
                "dest": f.get("dest_rel")}
               for f in unreadable]}
    with open(os.path.join(reg_dir, "数据缺口登记.json"), "w", encoding="utf-8") as w:
        json.dump(reg, w, ensure_ascii=False, indent=1)
    lines = ["# 低置信度归类与数据缺口（01 建档与材料归集）", "",
             f"- 项目：{args.project}",
             f"- 文件总数：{len(manifest.get('files', []))}；"
             f"低置信度入档：{len(low)}；文件不可读：{len(unreadable)}", ""]
    if low:
        lines += ["## 低置信度入档（不阻断流水线）", "",
                  "| 文件 | 入档类别 | 置信度 | 说明 |", "|---|---|---|---|"]
        for f in low:
            lines.append(f"| {f.get('rel','')} | {f.get('category_label','')} | "
                         f"{float(f.get('confidence') or 0):.2f} | "
                         f"{str(f.get('low_confidence_reason') or '').replace('|', '/')} |")
        lines.append("")
    if unreadable:
        lines += ["## 数据缺口（文件不可读）", "",
                  "| 文件 | 原因 | 处置 |", "|---|---|---|"]
        for f in unreadable:
            lines.append(f"| {f.get('rel','')} | {f.get('unreadable_reason','')} | "
                         "已按兜底类别入档，请补可读版本 |")
        lines.append("")
    if not low and not unreadable:
        lines.append("无低置信度条目，也无不可读文件。")
    with open(os.path.join(reg_dir, "低置信度归类.md"), "w", encoding="utf-8") as w:
        w.write("\n".join(lines) + "\n")

    # 项目索引（即使有待确认也要建，后续齐套检查靠它识别投递区）
    files = manifest.get("files", [])
    fingerprints = {}
    lines = [f"# 项目索引（{args.project}）", "",
             f"- 方向：{args.direction}",
             f"- 投递区：{args.delivery}",
             f"- 投递区实际路径：`{args.delivery}`",
             f"- 模板：{os.path.basename(args.template)}（sha256 {sha256(args.template)[:16]}…）",
             ("- 源报告：{}（sha256 {}…）".format(os.path.basename(args.source), sha256(args.source)[:16])
              if args.source and os.path.isfile(args.source)
              else "- 源报告：无（按散材料汇编，由 03 步骤生成主线）"),
             f"- 资料文件：{len(files)} 个（入档 {auto}，其中低置信度 {len(low)}；"
             f"文件不可读 {len(unreadable)}；人工确认 {len(manifest.get('decisions_applied', []))}）", "",
             "## 原件指纹（只读校验基线）", "",
             "| 文件 | sha256 | 大小 |", "| --- | --- | --- |"]
    for f in files:
        abs_path = os.path.join(args.project_dir, args.delivery, f.get("rel", ""))
        if os.path.exists(abs_path):
            digest = sha256(abs_path)
            fingerprints[f"输入文件/{f['rel']}".replace("/", os.sep)] = {
                "sha256": digest, "size": f.get("size", 0)}
            lines.append(f"| {f['rel']} | {digest} | {f.get('size_h','')} |")
    index_text = "\n".join(lines) + "\n"
    with open(os.path.join(args.project_dir, "00_项目索引.md"), "w", encoding="utf-8") as w:
        w.write(index_text)
    with open(args.index_out, "w", encoding="utf-8") as w:
        w.write(index_text)
    if args.fingerprint:
        with open(args.fingerprint, "w", encoding="utf-8") as w:
            json.dump({"schema": "cardo/fingerprint/v1", "files": fingerprints},
                      w, ensure_ascii=False, indent=1)

    if low:
        print(f"[01] 低置信度入档 {len(low)} 个（不阻断流水线，已登记缺口）："
              f"{os.path.join(reg_dir, '低置信度归类.md')}")
    if unreadable:
        print(f"[01] 文件不可读 {len(unreadable)} 个（真缺口，已登记；建议补可读版本后重跑）：")
        for f in unreadable:
            print(f"   不可读：{f.get('rel')}（{f.get('unreadable_reason')}）")
    if pending:
        # 兼容旧清单：pending 已不再作为闸门，照常往下走。
        print(f"[01] 另有 {len(pending)} 个旧口径待确认项，已按现状入档，不阻断流水线")

    print(f"[01] 分类确认完成：入档 {auto} / 低置信度 {len(low)} / 不可读 {len(unreadable)}"
          f" / 人工校正 {len(manifest.get('decisions_applied', []))}；索引 → {args.index_out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
