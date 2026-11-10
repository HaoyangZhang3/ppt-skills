# -*- coding: utf-8 -*-
"""静态核验镂空字 PPTX：页数、透明字孔、对象身份、位移和 Morph。"""
import argparse
import hashlib
import io
import json
import re
import sys

from lxml import etree
from PIL import Image
from pptx import Presentation

MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
P159 = '{http://schemas.microsoft.com/office/powerpoint/2015/09/main}morph'
NAMES = ('!!loukong-bg', '!!loukong-mask-top', '!!loukong-mask-bottom')


def fail(condition, message):
    if not condition:
        raise ValueError(message)


def main(argv=None):
    parser = argparse.ArgumentParser(description='核验 ppt-loukong-skill 生成的 PPTX。')
    parser.add_argument('pptx')
    parser.add_argument('--mode', choices=['opening', 'closing'], required=True)
    parser.add_argument('--morph-ms', type=int, default=2000)
    args = parser.parse_args(argv)
    try:
        prs = Presentation(args.pptx)
        expected = 3 if args.mode == 'opening' else 2
        fail(len(prs.slides) == expected, '页数应为 %d，实际 %d' % (expected, len(prs.slides)))
        hashes = {name: set() for name in NAMES}
        positions = []
        durations = []
        for index, slide in enumerate(prs.slides):
            shapes = {shape.name: shape for shape in slide.shapes}
            for name in NAMES:
                fail(name in shapes, '第 %d 页缺少 %s' % (index + 1, name))
                hashes[name].add(hashlib.sha256(shapes[name].image.blob).hexdigest())
            positions.append({name: [shapes[name].left, shapes[name].top,
                                     shapes[name].width, shapes[name].height]
                              for name in NAMES})
            for name in NAMES[1:]:
                with Image.open(io.BytesIO(shapes[name].image.blob)) as im:
                    fail(im.mode == 'RGBA', '遮罩必须为 RGBA')
                    lo, hi = im.getchannel('A').getextrema()
                    fail(lo == 0 and hi > 0, '遮罩必须同时含透明字孔与不透明区域')
            morph = slide._element.find(MC)
            if index == 0:
                fail(morph is None, '第一页不应有进入本页的 Morph')
            else:
                fail(morph is not None and morph.find('.//' + P159) is not None,
                     '第 %d 页缺少 byObject Morph' % (index + 1))
                xml = etree.tostring(morph, encoding='unicode')
                match = re.search(r'p14:dur="(\d+)"', xml)
                fail(match and int(match.group(1)) == args.morph_ms,
                     '第 %d 页 Morph 时长不符' % (index + 1))
                durations.append(int(match.group(1)))
        for name, values in hashes.items():
            fail(len(values) == 1, '%s 跨页图像内容不一致' % name)
        sw, sh = prs.slide_width, prs.slide_height
        center = 0 if args.mode == 'closing' else 1

        def hole_span(blob):
            """透明字孔（alpha<128）在遮罩图里的水平范围与图宽。"""
            with Image.open(io.BytesIO(blob)) as im:
                w_px, _ = im.size
                holes = im.getchannel('A').point(lambda v: 255 if v < 128 else 0)
                bb = holes.getbbox()
            fail(bb is not None, '遮罩缺少透明字孔')
            return bb, w_px

        spans = {name: hole_span(
            {s.name: s for s in prs.slides[0].shapes}[name].image.blob)
            for name in NAMES[1:]}

        def hole_page_x(slide_index, name, side):
            left, _, width, _ = positions[slide_index][name]
            bb, w_px = spans[name]
            k = width / float(w_px)
            return left + (bb[0] if side == 'left' else bb[2]) * k, k

        for name in NAMES[1:]:
            widths = [p[name][2] for p in positions]
            fail(len(set(widths)) == 1, '%s 跨页显示宽度不一致' % name)
            fail(widths[0] > sw, '%s 应为宽幅遮罩（比页面宽）' % name)
        overhang = (positions[center]['!!loukong-mask-top'][2] - sw) // 2
        fail(positions[center]['!!loukong-mask-top'][0] == -overhang and
             positions[center]['!!loukong-mask-top'][1] == 0, '居中页上遮罩位置错误')
        fail(positions[center]['!!loukong-mask-bottom'][0] == -overhang and
             positions[center]['!!loukong-mask-bottom'][1] == sh // 2,
             '居中页下遮罩位置错误')
        for name in NAMES[1:]:
            hx0, k = hole_page_x(center, name, 'left')
            hx1, _ = hole_page_x(center, name, 'right')
            fail(hx0 >= -k and hx1 <= sw + k, '居中页 %s 字孔未落在页面内' % name)
        if args.mode == 'opening':
            for name in NAMES[1:]:
                left = positions[0][name][0]
                fail(left <= 0 and left + positions[0][name][2] >= sw,
                     '首页 %s 未盖满页面' % name)
            hx0, k = hole_page_x(0, '!!loukong-mask-top', 'left')
            fail(hx0 >= sw - k, '首页上条字孔未整体停在页面右侧之外')
            hx1, _ = hole_page_x(0, '!!loukong-mask-bottom', 'right')
            fail(hx1 <= k, '首页下条字孔未整体停在页面左侧之外')
        last = positions[-1]
        fail(last['!!loukong-mask-top'][1] + last['!!loukong-mask-top'][3] <= 0,
             '上遮罩未完全退出画面')
        fail(last['!!loukong-mask-bottom'][1] >= sh, '下遮罩未完全退出画面')
        fail(last['!!loukong-bg'][2] > positions[0]['!!loukong-bg'][2] and
             last['!!loukong-bg'][3] > positions[0]['!!loukong-bg'][3], '底图未放大')
        if args.mode == 'closing':
            last_names = {shape.name for shape in prs.slides[-1].shapes}
            fail('loukong-end-message' in last_names and 'loukong-end-subtitle' in last_names,
                 '谢幕页缺少结束语')
        print(json.dumps({'status': 'PASS', 'mode': args.mode, 'slides': len(prs.slides),
                          'morph_ms': durations, 'transparent_masks': True,
                          'stable_image_content': True, 'mask_positions': positions},
                         ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print('FAIL: %s' % exc, file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
