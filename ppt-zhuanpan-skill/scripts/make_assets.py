# -*- coding: utf-8 -*-
"""生成暗灰机械风 PPT 素材：透明底旋转齿轮 GIF ×2 + 机械纹理背景 PNG。"""
import math
import os
from PIL import Image, ImageDraw, ImageFilter, ImageChops

BASE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.normpath(os.path.join(BASE, '..', 'assets'))
os.makedirs(ASSETS, exist_ok=True)


def gear_mask(size, teeth=12, r_tip_frac=0.48, r_root_ratio=0.80,
              hole_ratio=0.20, spoke_holes=6, supersample=4):
    """齿轮轮廓 L 掩膜：梯形齿 + 中心孔 + 轮辐减重孔。supersample 抗锯齿。"""
    S = size * supersample
    m = Image.new('L', (S, S), 0)
    d = ImageDraw.Draw(m)
    cx = cy = S / 2
    r_tip = S * r_tip_frac
    r_root = r_tip * r_root_ratio
    step = 2 * math.pi / teeth
    pts = []
    for k in range(teeth):
        a = k * step
        for rr, aa in [(r_root, a), (r_tip, a + step * 0.16),
                       (r_tip, a + step * 0.40), (r_root, a + step * 0.56)]:
            pts.append((cx + rr * math.cos(aa), cy + rr * math.sin(aa)))
    d.polygon(pts, fill=255)
    # 中心孔
    rh = r_tip * hole_ratio
    d.ellipse([cx - rh, cy - rh, cx + rh, cy + rh], fill=0)
    # 轮辐减重孔
    if spoke_holes:
        rr = r_root * 0.58
        hr = r_root * 0.15
        for k in range(spoke_holes):
            a = k * 2 * math.pi / spoke_holes + step * 0.28
            x, y = cx + rr * math.cos(a), cy + rr * math.sin(a)
            d.ellipse([x - hr, y - hr, x + hr, y + hr], fill=0)
    return m.resize((size, size), Image.LANCZOS)


def paint_gear(size, teeth=12, reverse=False):
    """金属灰齿轮 RGBA 图：对角渐变 + 深色描边。"""
    mask = gear_mask(size, teeth=teeth)
    g = Image.linear_gradient('L').resize((size, size)).rotate(-45)
    grad = Image.merge('RGB', tuple(
        g.point(lambda v, a=a, b=b: int(a + (b - a) * v / 255))
        for a, b in zip((178, 184, 190), (96, 102, 109))))
    out = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    out.paste(grad, (0, 0), mask)
    # 深色描边：mask 膨胀 - 原 mask
    edge = ImageChops.subtract(mask.filter(ImageFilter.MaxFilter(5)), mask)
    dark = Image.new('RGBA', (size, size), (40, 43, 47, 255))
    out.paste(dark, (0, 0), edge.point(lambda v: v // 2))
    return out


def make_gif(path, size, teeth, n_frames=24, duration=60, reverse=False):
    gear = paint_gear(size * 2, teeth=teeth).resize((size, size), Image.LANCZOS)
    frames = []
    for i in range(n_frames):
        ang = (360.0 / n_frames) * i * (-1 if reverse else 1)
        frames.append(gear.rotate(ang, resample=Image.BICUBIC))

    # 统一调色板 + 显式透明索引 255（Pillow 自动转换会丢透明）
    pal = frames[0].convert('RGB').convert('P', palette=Image.ADAPTIVE, colors=255)
    pframes = []
    for f in frames:
        alpha = f.getchannel('A')
        p = f.convert('RGB').quantize(palette=pal)
        p.paste(255, alpha.point(lambda a: 255 if a < 128 else 0))
        pframes.append(p)
    pframes[0].save(path, save_all=True, append_images=pframes[1:],
                    duration=duration, loop=0, disposal=2, transparency=255,
                    optimize=False)
    print('gif:', path, os.path.getsize(path), 'bytes')


def make_bg(path, w=1920, h=1080):
    img = Image.new('RGB', (w, h))
    d = ImageDraw.Draw(img)
    for y in range(h):
        t = y / (h - 1)
        c = tuple(int(a + (b - a) * t) for a, b in zip((19, 20, 23), (30, 32, 36)))
        d.line([(0, y), (w, y)], fill=c)
    # 细网格
    for x in range(0, w, 96):
        d.line([(x, 0), (x, h)], fill=(35, 38, 43))
    for y in range(0, h, 96):
        d.line([(0, y), (w, y)], fill=(35, 38, 43))
    # 右下大齿轮暗纹
    gm = gear_mask(1100, teeth=16)
    layer = Image.new('RGB', (1100, 1100), (54, 58, 64))
    img.paste(layer, (int(w * 0.62), int(h * 0.40)), gm.point(lambda v: v * 0.55))
    # 左上小齿轮暗纹
    gm2 = gear_mask(560, teeth=12)
    layer2 = Image.new('RGB', (560, 560), (46, 49, 55))
    img.paste(layer2, (-160, -140), gm2.point(lambda v: v * 0.5))
    # 顶部刻度尺
    d.line([(0, 64), (w, 64)], fill=(62, 66, 72), width=2)
    for x in range(0, w, 24):
        tall = (x % 120 == 0)
        d.line([(x, 64), (x, 64 - (16 if tall else 8))], fill=(70, 74, 80), width=2 if tall else 1)
    # 底部信息条细线
    d.line([(0, h - 72), (w, h - 72)], fill=(52, 56, 61), width=1)
    # 四角螺丝
    for sx, sy in [(36, 108), (w - 36, 108), (36, h - 108), (w - 36, h - 108)]:
        d.ellipse([sx - 9, sy - 9, sx + 9, sy + 9], fill=(58, 62, 67), outline=(86, 91, 97))
        d.line([(sx - 5, sy), (sx + 5, sy)], fill=(24, 26, 29), width=2)
        d.line([(sx, sy - 5), (sx, sy + 5)], fill=(24, 26, 29), width=2)
    img.save(path, 'PNG')
    print('bg:', path, os.path.getsize(path), 'bytes')


if __name__ == '__main__':
    make_bg(os.path.join(ASSETS, 'bg_mech.png'))
    make_gif(os.path.join(ASSETS, 'gear_big.gif'), 480, teeth=12, duration=60)
    make_gif(os.path.join(ASSETS, 'gear_small.gif'), 300, teeth=9, duration=50, reverse=True)
