# -*- coding: utf-8 -*-
"""
生成 DocFormatter 应用图标 (ICO)。
设计：几个图案组合的「简洁扁平」图标 ——
  - 圆角方块背景（主题蓝，扁平单色）
  - 文档页（白色圆角矩形 + 右上折角，代表「文档」）
  - 三条文字行（浅蓝圆角条，代表「排版/内容」）
  - 右下角对勾徽标（白底蓝勾，代表「一键规范完成」）
整体干净不花哨，比旧版去掉了闪光星与对齐网格，更克制。
"""
import os
from PIL import Image, ImageDraw

BRAND = (45, 127, 249)        # 主题蓝
BRAND_DARK = (28, 99, 209)
WHITE = (255, 255, 255)
LINE = (120, 168, 247)        # 文字行（浅蓝）

SS = 4  # 超采样倍数，保证缩放后边缘平滑


def _draw_at(s: int) -> Image.Image:
    """在 s×s（已超采样）画布上绘制图标，返回 RGBA 图像。"""
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # 背景圆角方块
    pad = int(s * 0.09)
    d.rounded_rectangle((pad, pad, s - pad, s - pad), radius=int(s * 0.22), fill=BRAND)

    # 文档页（白色，略小，居中偏上；右上折角）
    px = int(s * 0.27)
    pw = int(s * 0.46)
    py = int(s * 0.19)
    ph = int(s * 0.60)
    fold = int(s * 0.11)
    # 页身（多边形：左上 -> 右上折角起点 -> 折角尖 -> 右下 -> 左下）
    page = [
        (px, py),
        (px + pw - fold, py),
        (px + pw, py + fold),
        (px + pw, py + ph),
        (px, py + ph),
    ]
    d.polygon(page, fill=WHITE)
    # 折角缺口（用背景色三角覆盖，形成翻页效果）
    d.polygon([(px + pw - fold, py), (px + pw, py + fold), (px + pw - fold, py + fold)], fill=BRAND)

    # 三条文字行（末行短一点）
    lx1 = px + int(s * 0.08)
    lx2 = px + pw - int(s * 0.08)
    ly = py + int(s * 0.16)
    gap = int(s * 0.135)
    bar_h = max(2, int(s * 0.034))
    for k in range(3):
        lw = (lx2 - lx1) if k < 2 else int((lx2 - lx1) * 0.60)
        d.rounded_rectangle((lx1, ly + k * gap, lx1 + lw, ly + k * gap + bar_h),
                            radius=max(2, int(s * 0.014)), fill=LINE)

    # 右下角对勾徽标（白底圆 + 蓝勾），放在蓝底上、页面右下角外侧一点
    cx = px + pw - int(s * 0.01)
    cy = py + ph - int(s * 0.01)
    r = int(s * 0.13)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=WHITE)
    # 对勾：两段线
    t = int(s * 0.045)
    p1 = (cx - int(s * 0.06), cy + int(s * 0.00))
    p2 = (cx - int(s * 0.01), cy + int(s * 0.06))
    p3 = (cx + int(s * 0.07), cy - int(s * 0.05))
    d.line((p1, p2), fill=BRAND, width=t, joint="curve")
    d.line((p2, p3), fill=BRAND, width=t, joint="curve")

    return img


def generate_icon(out_path: str, sizes=(16, 24, 32, 48, 64, 128, 256)) -> str:
    """生成多尺寸 ICO，并附一张预览 PNG（同目录 icon_preview.png）。

    说明：本机 Pillow 版本的 ICO 保存，只有「单张高分辨率基准图 + sizes 参数」才会
    真正内嵌多分辨率；append_images 多张图的方式只会写出第一张(16x16)，造成桌面图标缺失。
    因此这里用超采样画一张最大尺寸基准图，再让 Pillow 自动缩放出全套尺寸。
    """
    big = _draw_at(max(sizes) * SS)
    base = big.resize((max(sizes), max(sizes)), Image.LANCZOS).convert("RGBA")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    base.save(out_path, format="ICO", sizes=[(sz, sz) for sz in sizes])
    # 预览
    preview = os.path.join(os.path.dirname(out_path), "icon_preview.png")
    _draw_at(256 * SS).resize((256, 256), Image.LANCZOS).save(preview)
    return out_path


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "app_icon_simple.ico")
    p = generate_icon(out)
    print("icon ->", p)
    print("preview ->", os.path.join(here, "icon_preview.png"))
