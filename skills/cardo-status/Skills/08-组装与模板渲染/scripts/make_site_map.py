# -*- coding: utf-8 -*-
"""用 OSGB 成果生成报告“片区基础概况”现状图。

输入（全部来自上游 OSGB 成果文件，不改原文件）：
  raster/ortho_northup.png      正射影像（north-up）
  raster/ortho_georef.json      影像地理参考（extentLocalM / resM / crs）
  raster/class_rgba.png         建筑—道路—绿化地类掩膜
  raster/class_georef.json      掩膜地理参考（须与正射同网格）
  vector/buildings-outline.geojson  建筑轮廓（局地米制）

输出：一张含项目范围、正射底图、地类、建筑轮廓、图例、指北针、比例尺、
坐标系与数据来源说明的 PNG。底图质量不足时在图内明确标注限制，不另造示意图。
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

W, H = 1800, 1370
MAP_BOX = (90, 190, 1310, 1075)  # left, top, right, bottom
FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
]


def font(size: int):
    for p in FONT_CANDIDATES:
        if Path(p).exists():
            try:
                return ImageFont.truetype(p, size, index=0)
            except Exception:
                try:
                    return ImageFont.truetype(p, size)
                except Exception:
                    continue
    return ImageFont.load_default()


def find_first(root: Path, name: str) -> Path | None:
    hits = sorted(root.rglob(name))
    return hits[0] if hits else None


def dashed_line(draw, p0, p1, color, width=3, dash=14, gap=9):
    x0, y0 = p0
    x1, y1 = p1
    dx, dy = x1 - x0, y1 - y0
    dist = math.hypot(dx, dy)
    if dist <= 0:
        return
    ux, uy = dx / dist, dy / dist
    pos = 0.0
    while pos < dist:
        end = min(pos + dash, dist)
        draw.line([(x0 + ux * pos, y0 + uy * pos), (x0 + ux * end, y0 + uy * end)],
                  fill=color, width=width)
        pos = end + gap


def dashed_rect(draw, box, color, width=3):
    l, t, r, b = box
    dashed_line(draw, (l, t), (r, t), color, width)
    dashed_line(draw, (r, t), (r, b), color, width)
    dashed_line(draw, (r, b), (l, b), color, width)
    dashed_line(draw, (l, b), (l, t), color, width)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--workspace", required=True)
    ap.add_argument("--baseline", required=False)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    ws = Path(args.workspace)
    inputs = ws / "输入文件"
    ortho = find_first(inputs, "ortho_northup.png")
    class_png = find_first(inputs, "class_rgba.png")
    georef = find_first(inputs, "ortho_georef.json")
    class_ref = find_first(inputs, "class_georef.json")
    bld = find_first(inputs, "buildings-outline.geojson")
    missing = [n for n, p in [("ortho_northup.png", ortho), ("class_rgba.png", class_png),
                              ("ortho_georef.json", georef), ("buildings-outline.geojson", bld)]
               if not p]
    if missing:
        raise SystemExit("现状图生成失败：OSGB 成果缺少 " + "、".join(missing) +
                         "；按“不静默降级”口径不生成示意图。")

    ref = json.loads(Path(georef).read_text(encoding="utf-8"))
    minx, miny, maxx, maxy = [float(x) for x in ref["extentLocalM"]]
    res = float(ref.get("resM") or 0.25)
    crs = (ref.get("crs") or {})
    crs_name = crs.get("name") or "CGCS2000 / 3-degree Gauss-Kruger CM 105E"
    epsg = crs.get("epsg") or 4544
    nx = int((ref.get("grid") or {}).get("nx") or round((maxx - minx) / res))
    ny = int((ref.get("grid") or {}).get("ny") or round((maxy - miny) / res))

    ortho_im = Image.open(ortho).convert("RGB")
    class_im = Image.open(class_png).convert("RGBA")
    if ortho_im.size != (nx, ny):
        ortho_im = ortho_im.resize((nx, ny), Image.LANCZOS)
    if class_im.size != (nx, ny):
        class_im = class_im.resize((nx, ny), Image.LANCZOS)

    # 地图区按等比例放入画布
    l, t, r, b = MAP_BOX
    box_w, box_h = r - l, b - t
    mpp = max((maxx - minx) / box_w, (maxy - miny) / box_h)
    draw_w = (maxx - minx) / mpp
    draw_h = (maxy - miny) / mpp
    ox = l + (box_w - draw_w) / 2
    oy = t + (box_h - draw_h) / 2

    map_w, map_h = int(round(draw_w)), int(round(draw_h))
    mask = class_im.split()[-1].point(lambda a: 255 if a > 8 else 0)
    base = Image.new("RGB", (map_w, map_h), (255, 255, 255))
    base.paste(ortho_im.resize((map_w, map_h), Image.LANCZOS), (0, 0), mask.resize((map_w, map_h), Image.LANCZOS))
    # 正射影像是模型纹理，压低饱和度后再叠地类，避免纹理拉伸干扰判读
    from PIL import ImageEnhance
    base = ImageEnhance.Color(base).enhance(0.55)
    base = ImageEnhance.Brightness(base).enhance(1.18)
    land = class_im.resize((map_w, map_h), Image.LANCZOS)
    land.putalpha(land.split()[-1].point(lambda a: int(a * 0.55)))
    base = Image.alpha_composite(base.convert("RGBA"), land).convert("RGB")

    canvas = Image.new("RGB", (W, H), (255, 255, 255))
    canvas.paste(base, (int(ox), int(oy)))
    d = ImageDraw.Draw(canvas)

    def px(x, y):
        return (ox + (x - minx) / mpp, oy + (maxy - y) / mpp)

    # 建筑轮廓
    try:
        fc = json.loads(Path(bld).read_text(encoding="utf-8"))
        for f in fc.get("features") or []:
            geom = f.get("geometry") or {}
            if geom.get("type") != "Polygon":
                continue
            rings = geom.get("coordinates") or []
            if not rings:
                continue
            pts = [px(float(x), float(y)) for x, y in rings[0]]
            if len(pts) >= 3:
                d.line(pts + [pts[0]], fill=(32, 48, 96), width=3)
    except Exception:
        pass

    # 项目范围（模型覆盖范围）虚线框
    dashed_rect(d, (ox, oy, ox + draw_w, oy + draw_h), (176, 32, 48), 4)

    # 内外框
    d.rectangle(MAP_BOX, outline=(90, 90, 90), width=2)

    # 标题
    d.text((90, 54), "图 2-1  项目测绘样区现状图", font=font(42), fill=(20, 20, 20))
    d.text((92, 112), f"项目范围：模型覆盖范围 {maxx - minx:.2f} 米 × {maxy - miny:.2f} 米"
                     f"（{((maxx - minx) * (maxy - miny)):.0f} 平方米）",
           font=font(24), fill=(70, 70, 70))

    # 指北针
    nx0, ny0 = 1240, 70
    d.polygon([(nx0, ny0 - 34), (nx0 - 18, ny0 + 18), (nx0, ny0 + 6), (nx0 + 18, ny0 + 18)],
              fill=(30, 30, 30))
    d.text((nx0 - 8, ny0 + 24), "N", font=font(26), fill=(30, 30, 30))

    # 比例尺 0—30 米
    sx, sy = ox + 24, oy + draw_h - 54
    seg = 10 / mpp
    for i in range(3):
        d.rectangle([sx + i * seg, sy, sx + (i + 1) * seg, sy + 14],
                    fill=(30, 30, 30) if i % 2 == 0 else (255, 255, 255),
                    outline=(30, 30, 30))
        d.text((sx + i * seg - 8, sy + 20), str(i * 10), font=font(22), fill=(30, 30, 30))
    d.text((sx + 3 * seg - 14, sy + 20), "30", font=font(22), fill=(30, 30, 30))
    d.text((sx + 3 * seg + 18, sy + 20), "m", font=font(22), fill=(30, 30, 30))

    # 图例
    lx, ly = 1350, 250
    d.text((lx, ly), "图  例", font=font(30), fill=(20, 20, 20))
    items = [
        ((230, 126, 34), "建筑屋面（地类识别）"),
        ((158, 158, 158), "道路与硬化表面"),
        ((76, 175, 80), "绿化与植被"),
        ((255, 255, 255), "模型未覆盖 / 无数据"),
    ]
    for i, (color, label) in enumerate(items):
        yy = ly + 58 + i * 42
        d.rectangle([lx, yy, lx + 30, yy + 24], fill=color, outline=(120, 120, 120))
        d.text((lx + 42, yy - 2), label, font=font(23), fill=(40, 40, 40))
    yy = ly + 58 + len(items) * 42 + 12
    d.line([(lx, yy + 10), (lx + 30, yy + 10)], fill=(176, 32, 48), width=4)
    d.text((lx + 42, yy - 2), "项目范围（模型覆盖，非红线）", font=font(23), fill=(40, 40, 40))
    yy += 42
    d.line([(lx, yy + 10), (lx + 30, yy + 10)], fill=(32, 48, 96), width=3)
    d.text((lx + 42, yy - 2), "建筑轮廓", font=font(23), fill=(40, 40, 40))

    # 说明：放在图下两栏，避免右侧竖排被画布裁切
    notes_left = [
        "坐标系：CGCS2000 / 3 度带高斯—克吕格，中央经线 105°E",
        f"（EPSG:{epsg}）；高程基准 EPSG:{crs.get('vertical', '5737')}。",
        "数据来源：OSGB 低空三维解析成果——正射影像、",
        "地类掩膜、建筑轮廓（同包《现状指标》依据数据）。",
    ]
    notes_right = [
        f"影像地面分辨率 {res:.2f} 米；底图为模型顶视正射，",
        "纹理存在拉伸、遮挡与色差，仅用于判读空间关系，",
        "不展示建筑立面与风貌细节。",
        "项目范围为模型覆盖范围，非宗地红线；四至以现状道路中心线为准。",
    ]
    ty = 1105
    for i, s in enumerate(notes_left):
        d.text((90, ty + i * 34), s, font=font(21), fill=(60, 60, 60))
    for i, s in enumerate(notes_right):
        d.text((760, ty + i * 34), s, font=font(21), fill=(60, 60, 60))

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out, "PNG", optimize=True)
    print(str(out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
