# -*- coding: utf-8 -*-
"""生成 examples 用的抽象演示图（塔影/雾/灯）。正式使用请换成自己的图或 AI 生图。

用法：
  python make_demo_assets.py
输出：examples/photos/cover.jpg、examples/photos/tower.jpg
"""
import math
import os
import random

from PIL import Image, ImageDraw, ImageFilter

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, '..', 'examples', 'photos')


def vgrad(w, h, top, bottom):
    im = Image.new('RGB', (w, h))
    d = ImageDraw.Draw(im)
    for y in range(h):
        t = y / float(h - 1)
        c = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
        d.line([(0, y), (w, y)], fill=c)
    return im


def glow(im, cx, cy, r, color, alpha):
    layer = Image.new('RGBA', im.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=color + (alpha,))
    d.ellipse([cx - r * 0.45, cy - r * 0.45, cx + r * 0.45, cy + r * 0.45],
              fill=color + (min(255, int(alpha * 1.6)),))
    layer = layer.filter(ImageFilter.GaussianBlur(r * 0.30))
    im.alpha_composite(layer)


def window_lights(im, cx, top_y, tier_h, tiers, w):
    layer = Image.new('RGBA', im.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    for i in range(tiers):
        y = top_y + (i + 0.55) * tier_h
        for dx in (-w * 0.015, w * 0.015):
            d.rectangle([cx + dx - w * 0.009, y - w * 0.014,
                         cx + dx + w * 0.009, y + w * 0.014],
                        fill=(255, 238, 190, 255))
    layer = layer.filter(ImageFilter.GaussianBlur(max(2, w * 0.010)))
    im.alpha_composite(layer)


def fog(im, blobs, color=(232, 210, 207), alpha=14):
    layer = Image.new('RGBA', im.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    for cx, cy, rx, ry in blobs:
        d.ellipse([cx - rx, cy - ry, cx + rx, cy + ry], fill=color + (alpha,))
    layer = layer.filter(ImageFilter.GaussianBlur(60))
    im.alpha_composite(layer)


def pagoda(d, cx, base_y, w, h, color):
    """塔剪影：多层飞檐 + 塔身 + 塔刹。"""
    tiers = 6
    tier_h = h / (tiers + 1.6)
    for i in range(tiers):
        t = i / float(tiers - 1)
        tw = w * (1.0 - 0.55 * t)
        roof_y = base_y - (i + 1) * tier_h
        body_w = tw * 0.58
        body_top = roof_y + tier_h * 0.42
        d.rectangle([cx - body_w / 2, body_top, cx + body_w / 2,
                     roof_y + tier_h * 1.02], fill=color)
        eave = tw * 0.10
        d.polygon([
            (cx - tw / 2 - eave, roof_y + tier_h * 0.10),
            (cx - tw / 2, roof_y - tier_h * 0.16),
            (cx + tw / 2, roof_y - tier_h * 0.16),
            (cx + tw / 2 + eave, roof_y + tier_h * 0.10),
            (cx + tw / 2 * 0.92, roof_y + tier_h * 0.34),
            (cx - tw / 2 * 0.92, roof_y + tier_h * 0.34),
        ], fill=color)
    top_y = base_y - (tiers + 0.55) * tier_h
    d.rectangle([cx - w * 0.02, top_y, cx + w * 0.02, top_y + tier_h * 0.55],
                fill=color)
    d.ellipse([cx - w * 0.035, top_y - w * 0.035, cx + w * 0.035,
               top_y + w * 0.035], fill=color)


def grain(im, amp=7, seed=7):
    rnd = random.Random(seed)
    px = im.load()
    w, h = im.size
    for y in range(h):
        for x in range(w):
            n = rnd.randint(-amp, amp)
            r, g, b = px[x, y][:3]
            px[x, y] = (max(0, min(255, r + n)), max(0, min(255, g + n)),
                        max(0, min(255, b + n)))
    return im


def vignette(im, strength=0.55):
    w, h = im.size
    mask = Image.new('L', (w, h), 0)
    d = ImageDraw.Draw(mask)
    d.ellipse([-w * 0.18, -h * 0.18, w * 1.18, h * 1.18], fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(min(w, h) * 0.18))
    dark = Image.new('RGB', (w, h), (2, 0, 0))
    return Image.composite(im, Image.blend(im, dark, strength), mask)


def make(size, tower_cx, tower_base, tower_w, tower_h, seed):
    w, h = size
    im = vgrad(w, h, (46, 0, 3), (3, 0, 0)).convert('RGBA')
    glow(im, w * tower_cx, h * (tower_base - tower_h * 0.55),
         min(w, h) * 0.11, (255, 242, 204), 95)
    glow(im, w * 0.80, h * 0.18, min(w, h) * 0.08, (232, 210, 207), 30)
    fog(im, [(w * 0.2, h * 0.72, w * 0.35, h * 0.10),
             (w * 0.75, h * 0.60, w * 0.30, h * 0.08),
             (w * 0.5, h * 0.86, w * 0.42, h * 0.09)])
    d = ImageDraw.Draw(im)
    pagoda(d, w * tower_cx, h * tower_base, w * tower_w, h * tower_h, (6, 0, 1))
    tiers = 6
    tier_h = (h * tower_h) / (tiers + 1.6)
    window_lights(im, w * tower_cx, h * tower_base - (tiers + 1) * tier_h,
                  tier_h, tiers, w * tower_w)
    base = ImageDraw.Draw(im)
    base.rectangle([0, h * 0.94, w, h], fill=(4, 0, 1))
    im = vignette(im.convert('RGB'), 0.5)
    im = grain(im, amp=6, seed=seed)
    return im


def main():
    os.makedirs(OUT, exist_ok=True)
    cover = make((1792, 1008), tower_cx=0.72, tower_base=0.86, tower_w=0.20,
                 tower_h=0.52, seed=11)
    cover.save(os.path.join(OUT, 'cover.jpg'), quality=90)
    tower = make((1024, 1280), tower_cx=0.50, tower_base=0.90, tower_w=0.34,
                 tower_h=0.62, seed=23)
    tower.save(os.path.join(OUT, 'tower.jpg'), quality=90)
    print('saved:', os.path.join(OUT, 'cover.jpg'))
    print('saved:', os.path.join(OUT, 'tower.jpg'))


if __name__ == '__main__':
    main()