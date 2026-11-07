# -*- coding: utf-8 -*-
"""静态核验喜鹊谋杀案风 PPTX：图层、灯光参数、柔边、动画、Morph、跨页同名对象。

用法：
  python verify_deck.py deck.pptx
  python verify_deck.py deck.pptx --wash-alpha 47 --cone-alpha 45 --morph-ms 1400
退出码 0 = PASS。
"""
import argparse
import json
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
P159 = '{http://schemas.microsoft.com/office/powerpoint/2015/09/main}morph'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
REQUIRED = ('xq-bg', 'xq-wash', 'xq-cone', 'xq-lamp-a', 'xq-lamp-b', 'xq-title')


def fail(cond, msg):
    if not cond:
        raise ValueError(msg)


def shape_alpha(shape):
    """取填充纯色的 a:alpha val（千分比），没有则 None。"""
    spPr = shape._element.spPr
    fill = spPr.find(qn('a:solidFill'))
    if fill is None:
        return None
    clr = fill.find(qn('a:srgbClr'))
    if clr is None:
        return None
    alpha = clr.find(qn('a:alpha'))
    return int(alpha.get('val')) if alpha is not None else None


def has_soft_edge(shape):
    spPr = shape._element.spPr
    return spPr.find(qn('a:effectLst')) is not None and \
        spPr.find(qn('a:effectLst')).find(qn('a:softEdge')) is not None


def main(argv=None):
    ap = argparse.ArgumentParser(description='核验 ppt-xqmsaf-skill 生成的 PPTX。')
    ap.add_argument('pptx')
    ap.add_argument('--wash-alpha', type=int, default=47)
    ap.add_argument('--cone-alpha', type=int, default=45)
    ap.add_argument('--morph-ms', type=int, default=1400)
    ap.add_argument('--wipe-ms', type=int, default=500)
    args = ap.parse_args(argv)

    try:
        prs = Presentation(args.pptx)
        fail(len(prs.slides) >= 2, '至少 2 页')
        report = []
        prev_names = None
        for index, slide in enumerate(prs.slides):
            n = index + 1
            shapes = {s.name: s for s in slide.shapes if s.name}
            for req in REQUIRED:
                fail(req in shapes, '第 %d 页缺少 %s' % (n, req))

            wash = shape_alpha(shapes['xq-wash'])
            fail(wash == (100 - args.wash_alpha) * 1000,
                 '第 %d 页红蒙版 alpha 应为 %d‰（透明度 %d%%），实际 %r'
                 % (n, (100 - args.wash_alpha) * 1000, args.wash_alpha, wash))
            cone = shape_alpha(shapes['xq-cone'])
            fail(cone == (100 - args.cone_alpha) * 1000,
                 '第 %d 页光锥 alpha 应为 %d‰，实际 %r'
                 % (n, (100 - args.cone_alpha) * 1000, cone))
            fail(has_soft_edge(shapes['xq-cone']), '第 %d 页光锥缺少柔化边缘' % n)
            fail(has_soft_edge(shapes['xq-glow']), '第 %d 页光晕缺少柔化边缘' % n)

            xml = etree.tostring(slide._element, encoding='unicode')
            morph = slide._element.find(MC)
            if index > 0:
                fail(morph is not None and morph.find('.//' + P159) is not None,
                     '第 %d 页缺少 byObject Morph' % n)
                m = re.search(r'p14:dur="(\d+)"', etree.tostring(morph, encoding='unicode'))
                fail(m and int(m.group(1)) == args.morph_ms,
                     '第 %d 页 Morph 时长不符' % n)
            timing = slide._element.find(qn('p:timing'))
            wipes = re.findall(r'filter="wipe\(down\)"', xml)
            timing_info = None
            if timing is not None:
                fail(len(wipes) >= 1, '第 %d 页有 timing 却没有擦除(自顶部)动画' % n)
                timing_info = '%d wipe' % len(wipes)
                if 'randombar(' in xml:
                    timing_info += ' + bars'
            lit = 'lit' if 'xq-lit' in shapes else ''
            report.append({'slide': n, 'layers': sorted(shapes.keys()),
                           'cone_alpha': cone, 'timing': timing_info, 'lit': lit,
                           'morph': index > 0})
        print(json.dumps({'status': 'PASS', 'slides': len(prs.slides),
                          'wash_alpha': args.wash_alpha,
                          'cone_alpha': args.cone_alpha,
                          'morph_ms': args.morph_ms,
                          'report': report}, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print('FAIL: %s' % exc, file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())