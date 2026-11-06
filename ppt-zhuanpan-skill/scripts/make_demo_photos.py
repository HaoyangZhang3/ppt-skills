# -*- coding: utf-8 -*-
"""生成 examples/double4.json 用的演示照片（抽象渐变 + 几何剪影，非真实照片）。

技能本身不附带任何图片素材；这个脚本只是为了让 examples 里的双盘配置开箱能跑。
正式使用请把自己/客户的照片路径写进 pages[].photo。

用法：
  python scripts/make_demo_photos.py                 # 输出到 examples/photos/
  python scripts/make_demo_photos.py -n 5 -o photos  # 5 张，输出到指定目录
"""
import argparse
import os

from PIL import Image, ImageDraw

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCENES = [
    ((214, 138, 66), (46, 32, 22)),
    ((120, 150, 190), (30, 40, 62)),
    ((96, 150, 108), (24, 44, 30)),
    ((188, 120, 140), (52, 26, 40)),
    ((110, 170, 170), (26, 48, 56)),
    ((150, 130, 190), (40, 32, 62)),
]


def make(path, top, bottom, tag):
    im = Image.new('RGB', (1600, 900), top)
    d = ImageDraw.Draw(im)
    for y in range(900):                              # 竖向渐变
        t = y / 900.0
        d.line([(0, y), (1600, y)],
               fill=tuple(int(top[c] * (1 - t) + bottom[c] * t) for c in range(3)))
    for k in range(6):                                # 方向感明显的山形剪影
        x0 = k * 300 - 100
        d.polygon([(x0, 900), (x0 + 260, 380 + k * 40), (x0 + 520, 900)],
                  fill=(bottom[0] // 2 + 10, bottom[1] // 2 + 10, bottom[2] // 2 + 10))
    d.rectangle([60, 60, 700, 150], fill=(255, 255, 255))   # 左上白块：一转就看得出转向
    d.text((80, 84), tag, fill=(20, 20, 20))
    im.save(path)
    return path


def main():
    ap = argparse.ArgumentParser(description='生成双盘示例用的抽象演示照片')
    ap.add_argument('-n', '--count', type=int, default=4, help='张数（1~6）')
    ap.add_argument('-o', '--out', default=os.path.join('examples', 'photos'))
    args = ap.parse_args()
    n = max(1, min(6, args.count))
    out = args.out if os.path.isabs(args.out) else os.path.join(BASE, args.out)
    os.makedirs(out, exist_ok=True)
    made = []
    for i in range(n):
        top, bottom = SCENES[i % len(SCENES)]
        made.append(make(os.path.join(out, 'demo-%d.png' % (i + 1)), top, bottom,
                         'DEMO PHOTO %d' % (i + 1)))
    print('demo photos: %d -> %s' % (n, out))
    for p in made:
        print('  ', p)


if __name__ == '__main__':
    main()
