# -*- coding: utf-8 -*-
"""把 designkit 下载的 Logo 产物整理成规范命名，并生成多尺寸 .ico。

用法（需要 Pillow）：
    python assets/icons/make_ico.py

输入（同目录下 designkit download 的产物，按文件名前缀定位）：
    01..04  logo_candidate_1..4.jpg   首轮候选（图形+文字）
    05      icon_white_bg.jpg         白底圆角图标（主图标来源）
    06      icon_dark_bg.jpg          深底圆角图标
    07      icon_transparent.png      纯图形 透明底（备用图标来源，与 09 内容相同）
    08      icon_symbol_only.jpg      纯图形 白底
    09      icon_symbol_transparent.png  纯图形 透明底

主图标逻辑：
    icon-white-bg.jpg 是「白色 squircle 底板 + 图形」铺满画布的 JPG（无 alpha），
    圆角外被压成黑色。利用这一点反解 alpha：非黑区为底板，边界过渡带用亮度做软 alpha，
    白区提亮到纯白，得到带真透明的圆角图标；再缩放输出多尺寸 .ico（PNG 压缩帧）。
    若白色版缺失，则退回「纯图形(带 alpha) + 自绘 squircle 底板」合成。

输出：
    icon-*.png / logo-candidate-*.jpg  规范命名
    ExcelSplitter.ico                  多尺寸图标（16~256，含 PNG 压缩帧，圆角外透明）
"""

import os
import sys

from PIL import Image, ImageChops, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))

# (下载产物文件名, 规范输出文件名)
RENAME_MAP = [
    ("01_536cb97ddbb640578aac3be7b5a89aa0.jpg", "logo-candidate-1.jpg"),
    ("02_ae4fd39e0b1f495fb3ad7611d65eb34d.jpg", "logo-candidate-2.jpg"),
    ("03_5ad943a5bbc54998bf6e16fe040e033a.jpg", "logo-candidate-3.jpg"),
    ("04_6eb2359e9ad046538e8cf742540b644f.jpg", "logo-candidate-4.jpg"),
    ("05_224329c73ff94a3d8bd95b9a64baa116.jpg", "icon-white-bg.jpg"),
    ("06_f6c705653de64eaeadf67aab71064aa6.jpg", "icon-dark-bg.jpg"),
    ("07_2bffa21a409c45308709fdc4ad2ca608.png", "icon-rounded-transparent.png"),
    ("08_0b990c72cd084e7f8e22673c87098ec0.jpg", "icon-symbol-white.jpg"),
    ("09_ac7cc125fa5d427bb96219eedf93ecf9.png", "icon-symbol-transparent.png"),
]

ICO_OUT = "ExcelSplitter.ico"
ICO_SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

# 主图标来源（按优先级）
SRC_WHITE = "icon-white-bg.jpg"            # 白色 squircle，反解 alpha
SRC_GLYPH = "icon-symbol-transparent.png"  # 纯图形透明底（回退方案用）

# 回退方案的自绘底板参数
TILE_CANVAS = 1024          # 底板画布边长
TILE_RADIUS_RATIO = 0.225   # 圆角半径比例（squircle 观感）
GLYPH_RATIO = 0.60          # 图形占底板边长比例
TILE_FILL_TOP = (255, 255, 255, 255)
TILE_FILL_BOTTOM = (240, 245, 251, 255)   # 极浅蓝渐变，避免纯白死板
TILE_BORDER = (222, 230, 240, 255)        # 发丝描边，白底页面上也能看出圆角轮廓


def icon_from_white_bg() -> Image.Image:
    """白色 squircle JPG（圆角外黑）-> 带 alpha 的 RGBA 图标。"""
    im = Image.open(os.path.join(HERE, SRC_WHITE)).convert("RGB")
    lum = im.convert("L")

    # 硬边界：非黑即底板
    hard = lum.point(lambda v: 255 if v > 30 else 0)
    # 边界过渡带（膨胀-腐蚀差集），在带内用亮度做软 alpha，圆角边缘才不会锯齿/发黑
    dil = hard.filter(ImageFilter.MaxFilter(7))
    ero = hard.filter(ImageFilter.MinFilter(7))
    rim = ImageChops.difference(dil, ero)
    alpha = Image.composite(lum, hard, rim)

    # 白区提亮到纯白（消除 JPG 的 254 灰）
    white_mask = lum.point(lambda v: 255 if v > 240 else 0)
    rgb = Image.composite(Image.new("RGB", im.size, (255, 255, 255)), im, white_mask)
    return Image.merge("RGBA", (*rgb.split(), alpha))


def icon_from_glyph_tile() -> Image.Image:
    """回退：纯图形(带 alpha) + 自绘 squircle 底板 -> RGBA 图标。"""
    glyph = Image.open(os.path.join(HERE, SRC_GLYPH)).convert("RGBA")
    bbox = glyph.getchannel("A").point(lambda a: 255 if a > 8 else 0).getbbox()
    if bbox:
        glyph = glyph.crop(bbox)

    size = TILE_CANVAS
    tile = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    radius = round(size * TILE_RADIUS_RATIO)

    # 垂直渐变底板：先画渐变图，再用圆角矩形做蒙版
    grad = Image.new("RGBA", (size, size))
    top, bottom = TILE_FILL_TOP, TILE_FILL_BOTTOM
    grad.putdata(
        [
            (
                round(top[0] + (bottom[0] - top[0]) * y / (size - 1)),
                round(top[1] + (bottom[1] - top[1]) * y / (size - 1)),
                round(top[2] + (bottom[2] - top[2]) * y / (size - 1)),
                255,
            )
            for y in range(size)
        ]
    )
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, size - 1, size - 1], radius=radius, fill=255
    )
    tile.paste(grad, (0, 0), mask)
    # 发丝描边（只描圆角矩形边界）
    border = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(border).rounded_rectangle(
        [1, 1, size - 2, size - 2], radius=radius - 1, outline=TILE_BORDER, width=6
    )
    tile.alpha_composite(border)

    # 居中贴图形
    target = round(size * GLYPH_RATIO)
    scale = target / max(glyph.size)
    glyph = glyph.resize((round(glyph.width * scale), round(glyph.height * scale)), Image.LANCZOS)
    tile.alpha_composite(glyph, ((size - glyph.width) // 2, (size - glyph.height) // 2))
    return tile


def build_icon() -> Image.Image:
    if os.path.exists(os.path.join(HERE, SRC_WHITE)):
        img = icon_from_white_bg()
        print(f"[info] ico 源: {SRC_WHITE}（反解 alpha） {img.size} {img.mode}")
    elif os.path.exists(os.path.join(HERE, SRC_GLYPH)):
        img = icon_from_glyph_tile()
        print(f"[info] ico 源: {SRC_GLYPH} + 自绘 squircle 底板 {img.size} {img.mode}")
    else:
        raise SystemExit("[error] 找不到任何可用于生成 ico 的源图")
    if img.mode != "RGBA":
        img = img.convert("RGBA")
    return img


def main() -> int:
    renamed = []
    for src, dst in RENAME_MAP:
        src_path = os.path.join(HERE, src)
        dst_path = os.path.join(HERE, dst)
        if not os.path.exists(src_path):
            print(f"[skip] 缺少 {src}")
            continue
        if os.path.abspath(src_path) != os.path.abspath(dst_path):
            os.replace(src_path, dst_path)
        renamed.append(dst)
        print(f"[ok] {src} -> {dst}")

    img = build_icon()
    ico_path = os.path.join(HERE, ICO_OUT)
    img.save(
        ico_path,
        format="ICO",
        sizes=ICO_SIZES,
        bitmap_format="png",  # 全部用 PNG 压缩帧，避免 BMP 帧过大
    )
    size_kb = os.path.getsize(ico_path) / 1024
    print(f"[ok] {ICO_OUT}  sizes={[s[0] for s in ICO_SIZES]}  {size_kb:.0f} KB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
