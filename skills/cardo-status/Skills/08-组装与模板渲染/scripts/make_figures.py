#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_figures.py -- 由聚合数据直接出两张图（PIL 手绘，无 matplotlib 依赖）。

只画能诚实的图：
  图 1  实测样区用地构成占比 —— 四类表面覆盖面积来自 OSGB 掩膜/纹理聚合值；
  图 2  实测样区建筑高度分位 + 片区台账层数 —— 高度取 P50/P90 聚合值，层数取台账逐小区登记值。

不画密度分级图、不画高度直方图：样区只有十几栋建筑，直方图画出来的形状是假的。
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhl.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
]

INK = (43, 43, 43)
GREY = (120, 120, 120)
LIGHT = (232, 232, 228)
COLORS = [(107, 122, 143), (127, 169, 139), (201, 183, 156), (217, 214, 207)]
BAR = (150, 160, 172)


def font(size: int):
    for p in FONT_CANDIDATES:
        if os.path.isfile(p):
            try:
                return ImageFont.truetype(p, size)
            except OSError:
                continue
    return ImageFont.load_default()


def text(d: ImageDraw.ImageDraw, xy, s: str, size=30, fill=INK, anchor="la"):
    d.text(xy, s, font=font(size), fill=fill, anchor=anchor)


def fmt_area(v: float) -> str:
    return f"{v:,.2f}"


def fig_land_use(data: dict, out: str) -> None:
    W, H = 1600, 760
    im = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(im)
    s = data.get("sample") or {}
    items = data.get("landUse") or []
    total = sum(float(i.get("area") or 0) for i in items) or 1.0

    text(d, (80, 46), "实测样区表面覆盖构成", 46)
    text(d, (80, 112), f"实测样区 {float(s.get('ha') or 0):.4f} 公顷 / {fmt_area(float(s.get('area') or 0))} 平方米；"
                       f"四类表面覆盖之和为样区面积（样区口径，不代表片区整体）", 26, GREY)

    x0, x1, y0, y1 = 80, W - 80, 232, 356
    bar_w = x1 - x0
    x = x0
    for i, it in enumerate(items):
        w = bar_w * float(it["area"]) / total
        xa, xb = x, x + w
        d.rectangle([xa, y0, xb, y1], fill=COLORS[i % len(COLORS)])
        ratio = f"{float(it['ratio']):.2f}%"
        # 窄段用小字号写在色块内，避免 7.72% 孤零零飘在色块外
        if w >= 150:
            text(d, ((xa + xb) / 2, (y0 + y1) / 2 - 4), ratio, 34, "white", "mm")
        elif w >= 100:
            text(d, ((xa + xb) / 2, (y0 + y1) / 2 - 4), ratio, 24, "white", "mm")
        else:
            text(d, ((xa + xb) / 2, y0 - 44), ratio, 28, INK, "mm")
        x = xb

    # 图例：一行两项，色块 + 名称 + 面积 + 占比（色块与文字留足间距，不压字）
    lx, ly = x0, 430
    for i, it in enumerate(items):
        col = lx + (i % 2) * 760
        row = ly + (i // 2) * 78
        d.rectangle([col, row + 8, col + 30, row + 42], fill=COLORS[i % len(COLORS)])
        text(d, (col + 56, row + 4),
             f"{it['name']}　{fmt_area(float(it['area']))} 平方米　{float(it['ratio']):.2f}%", 30)

    d.line([x0, 620, x1, 620], fill=LIGHT, width=2)
    text(d, (x0, 640), "数据来源：OSGB 现状指标.json（实测样区口径，分母为样区面积，不代表片区整体）", 24, GREY)
    im.save(out, dpi=(200, 200))


def fig_height_floors(data: dict, out: str) -> None:
    W, H = 1600, 1080
    im = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(im)
    h = data.get("height") or {}
    p50 = float(h.get("p50") or 0)
    p90 = float(h.get("p90") or 0)
    vmax = 25.0

    text(d, (80, 44), "实测样区建筑高度分位与片区台账层数", 46)
    text(d, (80, 106), "上：实测样区建筑高度（OSGB 高程聚合值）；下：片区住宅小区最高层数（台账登记值）", 26, GREY)

    # ── 上：高度分位标尺 ─────────────────────────────
    x0, x1, y = 120, W - 120, 330
    scale = (x1 - x0) / vmax
    max_h = float(h.get("maxHeight") or 0)
    d.rectangle([x0, y - 84, x0 + p90 * scale, y - 44], fill=(214, 219, 208))
    d.rectangle([x0, y - 84, x0 + p50 * scale, y - 44], fill=(127, 169, 139))
    d.line([x0, y, x1, y], fill=(200, 200, 196), width=3)
    for t in range(0, 26, 5):
        tx = x0 + t * scale
        d.line([tx, y, tx, y + 12], fill=(200, 200, 196), width=3)
        text(d, (tx, y + 22), f"{t} 米", 24, GREY, "ma")

    # 两段色带的图例：不给图例，读者分不清深浅两段各是哪一档
    lx = x0
    for label, color in [("样区 0–P50（半数建筑低于中位）", (127, 169, 139)),
                         ("样区 P50–P90", (214, 219, 208))]:
        d.rectangle([lx, y - 158, lx + 28, y - 130], fill=color)
        text(d, (lx + 40, y - 164), label, 24, GREY)
        lx += 40 + d.textlength(label, font=font(24)) + 46

    for v, label, color in [(p50, f"中位 P50　{p50:.2f} 米", (56, 96, 74)),
                            (p90, f"P90　{p90:.2f} 米", (96, 88, 60))]:
        tx = x0 + v * scale
        d.line([tx, y - 104, tx, y + 12], fill=color, width=4)
        text(d, (tx, y + 60), label, 30, color, "ma")

    # 台账最高住宅高度的参考线：这个值只在文字里出现，轴上没有锚点
    if max_h:
        mx = x0 + min(max_h, vmax) * scale
        for yy in range(y - 110, y + 8, 18):
            d.line([mx, yy, mx, yy + 9], fill=(150, 90, 90), width=3)
        text(d, (mx - 10, y - 128), f"台账最高 {max_h:.2f} 米", 24, (150, 90, 90), "rm")

    text(d, (x0, y + 110),
         f"P50–P90 区间 {float(h.get('diff') or 0):.2f} 米；样区按基底折算平均 {float(h.get('floors') or 0):.2f} 层；"
         f"台账住宅最高 {float(h.get('maxHeight') or 0):.2f} 米", 28)

    # ── 下：台账层数柱状 ─────────────────────────────
    rows = data.get("floors") or []
    top, bottom = 616, 904
    text(d, (80, top - 50), "片区台账住宅小区最高层数（单位：层）", 32)
    bx0, bx1 = 120, W - 120
    ymax = max([float(r.get("floors") or 0) for r in rows] + [1.0]) * 1.18
    n = max(1, len(rows))
    slot = (bx1 - bx0) / n
    bw = slot * 0.5
    for i, r in enumerate(rows):
        f = float(r.get("floors") or 0)
        cx = bx0 + slot * (i + 0.5)
        yb = bottom - (f / ymax) * (bottom - top)
        d.rectangle([cx - bw / 2, yb, cx + bw / 2, bottom], fill=BAR)
        text(d, (cx, yb - 14), f"{f:.0f} 层", 28, INK, "md")
        name = str(r.get("name") or "").replace("（电信科学技术第五研究所有限公司宿舍）", "")
        if len(name) > 8:
            name = name[:8]
        text(d, (cx, bottom + 14), name, 23, INK, "ma")
    d.line([bx0, bottom, bx1, bottom], fill=(180, 180, 176), width=2)

    d.line([80, 972, W - 80, 972], fill=LIGHT, width=2)
    text(d, (80, 990), "数据来源：OSGB 现状指标.json；社区改造小区基本信息及改造内容统计表", 24, GREY)
    text(d, (80, 1028), "口径：高度为实测样区高程聚合值，层数为台账逐小区登记值；不画高度直方图——样区建筑样本过少，直方图会失真", 22, GREY)
    im.save(out, dpi=(200, 200))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True, help="图表数据.json")
    ap.add_argument("--out-dir", required=True)
    args = ap.parse_args()
    with open(args.data, encoding="utf-8") as f:
        data = json.load(f)
    os.makedirs(args.out_dir, exist_ok=True)
    fig_land_use(data, os.path.join(args.out_dir, "fig_land_use.png"))
    fig_height_floors(data, os.path.join(args.out_dir, "fig_height_floors.png"))
    print(f"[图表] 已生成 2 张：{args.out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
