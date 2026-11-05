# -*- coding: utf-8 -*-
"""生成技能自带的抽象山景演示底图；非照片、非外部版权素材。"""
import math
import os
import random

from PIL import Image, ImageDraw, ImageFilter

W, H = 1920, 1080
OUT = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', 'assets', 'demo-bg.png'))


def main():
    im = Image.new('RGB', (W, H))
    px = im.load()
    for y in range(H):
        t = y / float(H - 1)
        for x in range(W):
            glow = max(0.0, 1.0 - math.hypot((x - W * .73) / W, (y - H * .32) / H) * 2.3)
            px[x, y] = (int(16 + 38 * t + 80 * glow),
                        int(31 + 24 * t + 40 * glow),
                        int(48 + 25 * t + 18 * glow))
    sun = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    sd = ImageDraw.Draw(sun)
    sd.ellipse((1370, 195, 1560, 385), fill=(245, 184, 118, 225))
    sun = sun.filter(ImageFilter.GaussianBlur(3))
    im = Image.alpha_composite(im.convert('RGBA'), sun)
    random.seed(7)
    layers = [
        (570, (43, 65, 76, 255), 155, 7),
        (710, (27, 48, 57, 255), 210, 8),
        (850, (14, 30, 35, 255), 180, 9),
    ]
    for base_y, color, amplitude, peaks in layers:
        points = [(-80, H)]
        phase = random.random() * math.pi
        for x in range(-80, W + 160, 80):
            y = base_y - amplitude * (0.45 + 0.55 * math.sin(x / W * math.pi * peaks + phase))
            y += random.uniform(-45, 45)
            points.append((x, int(y)))
        points.extend([(W + 80, H), (-80, H)])
        layer = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        ImageDraw.Draw(layer).polygon(points, fill=color)
        im = Image.alpha_composite(im, layer)
    haze = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    hd = ImageDraw.Draw(haze)
    for i in range(12):
        y = 490 + i * 20
        hd.rectangle((0, y, W, y + 16), fill=(165, 185, 183, max(0, 18 - i)))
    im = Image.alpha_composite(im, haze)
    im.convert('RGB').save(OUT, optimize=True)
    print('OK: 抽象山景演示底图已生成：%s' % OUT)


if __name__ == '__main__':
    main()
