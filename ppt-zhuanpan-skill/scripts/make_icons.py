# -*- coding: utf-8 -*-
"""生成 4 个钢色线性小图标（透明底 PNG）：笔、扳手、钳子、图纸。
用于大齿轮轮盘上当前章标签左侧的小元素，配合 Morph 做缩放+径向移动。
"""
import os
from PIL import Image, ImageDraw

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.normpath(os.path.join(BASE, '..', 'assets'))
STEEL = (201, 206, 212, 255)
S = 2  # 超采样倍数
W = 480


def canvas():
    img = Image.new('RGBA', (W * S, W * S), (0, 0, 0, 0))
    return img, ImageDraw.Draw(img)


def save(img, name):
    img = img.resize((W, W), Image.LANCZOS)
    img.save(os.path.join(OUT, name))
    print('saved', name)


def L(d, pts, w):
    d.line([(int(x * S), int(y * S)) for x, y in pts], fill=STEEL, width=int(w * S))


def pen():
    img, d = canvas()
    L(d, [(140, 340), (288, 192)], 44)                     # 笔杆
    d.polygon([(272 * S, 208 * S), (304 * S, 176 * S), (352 * S, 128 * S)], fill=STEEL)  # 笔尖
    L(d, [(128, 352), (148, 332)], 30)                     # 尾帽
    save(img, 'icon_pen.png')


def wrench():
    img, d = canvas()
    L(d, [(150, 330), (285, 195)], 34)                     # 手柄
    d.arc([253 * S, 103 * S, 377 * S, 227 * S], start=100, end=350, fill=STEEL, width=int(30 * S))  # 开口钳口
    d.ellipse([126 * S, 306 * S, 174 * S, 354 * S], outline=STEEL, width=int(20 * S))  # 尾部圆环
    save(img, 'icon_wrench.png')


def pliers():
    img, d = canvas()
    L(d, [(140, 340), (262, 218)], 30)                     # 左手柄
    L(d, [(340, 340), (218, 218)], 30)                     # 右手柄
    L(d, [(240, 210), (200, 110)], 30)                     # 左钳口
    L(d, [(240, 210), (280, 110)], 30)                     # 右钳口
    d.ellipse([224 * S, 194 * S, 256 * S, 226 * S], fill=STEEL)  # 轴铆钉
    save(img, 'icon_pliers.png')


def paper():
    img, d = canvas()
    L(d, [(150, 110), (280, 110), (330, 160), (330, 370), (150, 370), (150, 110)], 24)  # 外框带折角
    L(d, [(280, 110), (280, 160), (330, 160)], 20)         # 折角线
    L(d, [(180, 215), (300, 215)], 18)                     # 图线 1
    L(d, [(180, 265), (300, 265)], 18)                     # 图线 2
    L(d, [(180, 315), (265, 315)], 18)                     # 图线 3
    save(img, 'icon_paper.png')


if __name__ == '__main__':
    pen()
    wrench()
    pliers()
    paper()
