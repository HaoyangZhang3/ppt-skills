# -*- coding: utf-8 -*-
"""ppt-zhuanpan-skill — 机械风大齿轮转盘 PPT 生成器（自包含）。

效果：左侧大齿轮只露右半边（约 1/4 版面），每翻一页逆时针转 90°，
当前章「part + 数字」横排标签永远落在右缘 3 点钟方向，正对右侧内容区；
标签旁有工具小图标，翻页时旧图标缩小收回轮心、新图标自轮心放大伸出；
右上角常驻转动齿轮 GIF；暗灰机械风；全程 Morph 平滑切换。

用法：
  python zhuanpan_deck.py -o out.pptx                 # 内置演示（4 章 + 封面）
  python zhuanpan_deck.py -o out.pptx --config deck.json

依赖：python-pptx >= 0.6.19（素材已随技能附带，无需 Pillow，除非要重新生成素材）。
"""
import argparse
import json
import math
import os

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn, nsdecls

BASE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.normpath(os.path.join(BASE, '..', 'assets'))
PT = 12700
SW, SH = 12192000, 6858000  # 16:9

STEEL = RGBColor(0xC9, 0xCE, 0xD4)
TXT_MAIN = RGBColor(0xEC, 0xEE, 0xF0)
TXT_SUB = RGBColor(0xB6, 0xBD, 0xC4)
TXT_DIM = RGBColor(0x8A, 0x90, 0x97)
ACCENT = RGBColor(0x9F, 0xB4, 0xC6)

MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
P14_NS = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
P159_NS = 'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"'


# ---------------------------------------------------------------- XML  helpers

def set_slide_bg_picture(slide, img_path):
    image_part, rId = slide.part.get_or_add_image_part(img_path)
    cSld = slide._element.find(qn('p:cSld'))
    old = cSld.find(qn('p:bg'))
    if old is not None:
        cSld.remove(old)
    xml = ('<p:bg %s><p:bgPr><a:blipFill dpi="0" rotWithShape="1">'
           '<a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch>'
           '</a:blipFill><a:effectLst/></p:bgPr></p:bg>') % (nsdecls('p', 'a', 'r'), rId)
    cSld.insert(0, parse_xml(xml))


def add_outer_shadow(shape, blur_pt=8.0, dist_pt=14.0, dir_deg=30.0, alpha_pct=55):
    spPr = shape._element.spPr
    old = spPr.find(qn('a:effectLst'))
    if old is not None:
        spPr.remove(old)
    xml = ('<a:effectLst %s><a:outerShdw blurRad="%d" dist="%d" dir="%d" rotWithShape="0">'
           '<a:srgbClr val="000000"><a:alpha val="%d"/></a:srgbClr>'
           '</a:outerShdw></a:effectLst>') % (nsdecls('a'), int(blur_pt * PT),
                                              int(dist_pt * PT), int(dir_deg * 60000) % 21600000,
                                              int(alpha_pct * 1000))
    spPr.append(parse_xml(xml))


def set_gradient_fill(shape, color='767D84', ang_deg=45, alpha_start=100, alpha_end=100):
    spPr = shape._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'):
        el = spPr.find(qn(tag))
        if el is not None:
            spPr.remove(el)
    xml = ('<a:gradFill %s rotWithShape="1"><a:gsLst>'
           '<a:gs pos="0"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:gs>'
           '<a:gs pos="100000"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:gs>'
           '</a:gsLst><a:lin ang="%d" scaled="1"/></a:gradFill>') % (
        nsdecls('a'), color, int(alpha_start * 1000), color,
        int(alpha_end * 1000), int(ang_deg * 60000) % 21600000)
    spPr.find(qn('a:prstGeom')).addnext(parse_xml(xml))


def set_morph_transition(slide, dur_ms=1600, auto_advance_ms=None):
    sld = slide._element
    MC_TAG = '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
    for child in list(sld):
        if child.tag == MC_TAG or child.tag == qn('p:transition'):
            sld.remove(child)
    adv = ' advTm="%d"' % auto_advance_ms if auto_advance_ms else ''
    xml = ('<mc:AlternateContent %s %s %s><mc:Choice %s Requires="p159">'
           '<p:transition spd="slow" p14:dur="%d"%s><p159:morph option="byObject"/>'
           '</p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"%s/>'
           '</mc:Fallback></mc:AlternateContent>') % (MC_NS, nsdecls('p'), P14_NS,
                                                      P159_NS, dur_ms, adv, adv)
    sld.append(parse_xml(xml))


# ---------------------------------------------------------------- shape helpers

def add_text(shapes, x, y, w, h, text, size, color=TXT_MAIN, bold=False,
             align=PP_ALIGN.LEFT, name=None, spacing=None):
    box = shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = '微软雅黑'
    if spacing is not None:
        run.font._rPr.set('spc', str(int(spacing * 100)))
    return box


def add_bar(shapes, x, y, w, h, rgb, name=None):
    bar = shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x)), Emu(int(y)),
                           Emu(int(w)), Emu(int(h)))
    if name:
        bar.name = name
    bar.fill.solid()
    bar.fill.fore_color.rgb = rgb
    bar.line.fill.background()
    bar.shadow.inherit = False
    return bar


def add_gif_pair(shapes, big_cx, big_cy, big_w, scale=0.62):
    big = shapes.add_picture(os.path.join(ASSETS, 'gear_big.gif'),
                             Emu(int(big_cx - big_w / 2)), Emu(int(big_cy - big_w / 2)),
                             Emu(int(big_w)), Emu(int(big_w)))
    big.name = 'gear-anim-big'
    small_w = big_w * scale
    small = shapes.add_picture(os.path.join(ASSETS, 'gear_small.gif'),
                               Emu(int(big_cx + big_w * 0.68 - small_w / 2)),
                               Emu(int(big_cy + big_w * 0.66 - small_w / 2)),
                               Emu(int(small_w)), Emu(int(small_w)))
    small.name = 'gear-anim-small'


def add_footer(shapes, idx, total, footer_left, footer_name, left_x=None):
    y = SH * 0.945
    if left_x is None:
        left_x = SW * 0.035
    add_text(shapes, left_x, y, SW * 0.60, SH * 0.04, footer_left, 9, color=TXT_DIM)
    add_text(shapes, SW * 0.62, y, SW * 0.345, SH * 0.04,
             '%s · %02d / %02d' % (footer_name, idx, total), 9, color=TXT_DIM,
             align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------- wheel

def add_gear_roulette(shapes, cx, cy, r, offset, labels, current, icons):
    """大齿轮转盘。offset 为整盘旋转角（度，顺时针为正；每页 -90 即逆时针）。
    所有形状跨页同名，保证 Morph 按对象匹配。"""
    gear = shapes.add_shape(MSO_SHAPE.GEAR_9, Emu(int(cx - r)), Emu(int(cy - r)),
                            Emu(int(2 * r)), Emu(int(2 * r)))
    gear.name = 'roulette-gear'
    gear.rotation = offset % 360
    gear.line.fill.background()
    gear.shadow.inherit = False
    set_gradient_fill(gear, color='767D84', ang_deg=45, alpha_start=100, alpha_end=100)
    grad = gear._element.spPr.find(qn('a:gradFill'))
    grad.find(qn('a:gsLst'))[1].find(qn('a:srgbClr')).set('val', '454A51')
    add_outer_shadow(gear, blur_pt=8, dist_pt=14, dir_deg=30, alpha_pct=55)

    # 放射线：过圆心细矩形 + rotation，与标签错开 45°
    w = int(1.4 * PT)
    length = int(2 * r * 0.94)
    for k in range(4):
        ang = offset + (k + 0.5) * 90.0
        spoke = shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(cx - length // 2)),
                                 Emu(int(cy - w // 2)), Emu(length), Emu(w))
        spoke.name = 'roulette-spoke-%d' % k
        spoke.rotation = ang % 360
        spoke.fill.solid()
        spoke.fill.fore_color.rgb = STEEL
        spoke.line.fill.background()
        spoke.shadow.inherit = False

    # 工具图标接力：当前章放大在标签左侧，其余缩小藏于轮心（被轴孔遮住）
    if icons:
        big_size = int(0.85 * 914400)
        small_size = int(0.10 * 914400)
        big_cx = cx + int(r * 0.40)
        for k in range(4):
            if k == current:
                icx, icy, isize = big_cx, cy, big_size
            else:
                icx, icy, isize = cx, cy, small_size
            pic = shapes.add_picture(os.path.join(ASSETS, 'icon_%s.png' % icons[k]),
                                     Emu(int(icx - isize / 2)), Emu(int(icy - isize / 2)),
                                     Emu(isize), Emu(isize))
            pic.name = 'chapter-icon-%d' % k

    # 中心轴孔（盖住缩小的非当前图标）
    hr = int(r * 0.20)
    hub = shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx - hr)), Emu(int(cy - hr)),
                           Emu(2 * hr), Emu(2 * hr))
    hub.name = 'roulette-hub'
    hub.fill.solid()
    hub.fill.fore_color.rgb = RGBColor(0x1B, 0x1D, 0x20)
    hub.line.color.rgb = STEEL
    hub.line.width = Pt(1.0)
    hub.shadow.inherit = False

    # 横排两段式标签：小写 part 40pt + 数字 20pt，全部水平不旋转。
    # 基准角 k*90°，每页 offset -90° → 当前章永远落在右缘 3 点钟方向。
    label_w, label_h = int(1.5 * 914400), int(1.1 * 914400)
    rr = int(r * 0.70)
    for k, text in enumerate(labels):
        num = text.split()[-1]
        ang = offset + k * 90.0
        rad = math.radians(ang)
        px, py = cx + rr * math.cos(rad), cy + rr * math.sin(rad)
        box = shapes.add_textbox(Emu(int(px - label_w / 2)), Emu(int(py - label_h / 2)),
                                 Emu(label_w), Emu(label_h))
        box.name = 'roulette-label-%d' % k
        box.rotation = 0
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        run1 = p1.add_run()
        run1.text = 'part'
        run1.font.size = Pt(40)
        run1.font.color.rgb = TXT_MAIN
        run1.font.name = '微软雅黑'
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(6)
        run2 = p2.add_run()
        run2.text = num
        run2.font.size = Pt(20)
        run2.font.color.rgb = TXT_SUB
        run2.font.name = '微软雅黑'


def add_content(shapes, part, section, title, points):
    left = SW * 0.535
    width = SW * 0.425
    add_text(shapes, left, SH * 0.215, width, SH * 0.05,
             '%s · %s' % (part, section), 14, color=ACCENT, bold=True, spacing=2)
    add_text(shapes, left, SH * 0.285, width, SH * 0.13, title, 28, bold=True)
    add_bar(shapes, left, SH * 0.435, SW * 0.055, Pt(2.2), STEEL)
    y = SH * 0.49
    for pt_text in points:
        add_bar(shapes, left, y + Pt(6), Pt(5), Pt(5), ACCENT)
        add_text(shapes, left + Pt(14), y - Pt(2), width - Pt(14), SH * 0.10,
                 pt_text, 14, color=TXT_SUB)
        y += SH * 0.105


# ---------------------------------------------------------------- build

DEMO = {
    'footer_left': '演示汇报 · 齿轮转盘',
    'footer_name': '汇报人',
    'morph_ms': 1600,
    'direction': 'ccw',
    'cover': {
        'eyebrow': '演示汇报 · 2026',
        'title': '演示标题',
        'claim': '一句话核心主张，写在封面上。',
        'target': '主题副标题：写在这里',
        'name_line': '汇报人 · 单位 · 日期',
        'teacher': '',
    },
    'pages': [
        {'part': 'PART 01', 'section': '第一章', 'title': '第一章核心观点',
         'points': ['要点一：一句话', '要点二：一句话', '要点三：一句话'], 'icon': 'pen'},
        {'part': 'PART 02', 'section': '第二章', 'title': '第二章核心观点',
         'points': ['要点一：一句话', '要点二：一句话', '要点三：一句话'], 'icon': 'wrench'},
        {'part': 'PART 03', 'section': '第三章', 'title': '第三章核心观点',
         'points': ['要点一：一句话', '要点二：一句话', '要点三：一句话'], 'icon': 'pliers'},
        {'part': 'PART 04', 'section': '第四章', 'title': '第四章核心观点',
         'points': ['要点一：一句话', '要点二：一句话', '要点三：一句话'], 'icon': 'paper'},
    ],
}


def build(out_path, cfg):
    pages = cfg['pages']
    assert len(pages) == 4, '本技能固定 4 章（90° 步长对应 4 扇区）'
    labels = [pg.get('part', 'PART %02d' % (k + 1)) for k, pg in enumerate(pages)]
    icons = [pg.get('icon', '') for pg in pages]
    if not all(icons):
        icons = None
    footer_left = cfg.get('footer_left', '')
    footer_name = cfg.get('footer_name', '')
    morph_ms = cfg.get('morph_ms', 1600)
    step = -90.0 if cfg.get('direction', 'ccw') == 'ccw' else 90.0
    bg = os.path.join(ASSETS, 'bg_mech.png')
    total = len(pages) + (1 if cfg.get('cover') else 0)

    prs = Presentation()
    prs.slide_width = Emu(SW)
    prs.slide_height = Emu(SH)
    blank = prs.slide_layouts[6]

    cover = cfg.get('cover')
    if cover:
        s = prs.slides.add_slide(blank)
        set_slide_bg_picture(s, bg)
        sh = s.shapes
        add_text(sh, SW * 0.06, SH * 0.20, SW * 0.60, SH * 0.05,
                 cover.get('eyebrow', ''), 13, color=ACCENT, spacing=2)
        add_text(sh, SW * 0.06, SH * 0.29, SW * 0.60, SH * 0.15,
                 cover.get('title', ''), 52, bold=True)
        add_bar(sh, SW * 0.062, SH * 0.47, SW * 0.07, Pt(2.5), STEEL)
        add_text(sh, SW * 0.06, SH * 0.51, SW * 0.58, SH * 0.08,
                 cover.get('claim', ''), 17, color=TXT_SUB)
        add_text(sh, SW * 0.06, SH * 0.63, SW * 0.58, SH * 0.05,
                 cover.get('target', ''), 15, color=TXT_MAIN, bold=True)
        add_text(sh, SW * 0.06, SH * 0.71, SW * 0.62, SH * 0.05,
                 cover.get('name_line', ''), 12, color=TXT_DIM)
        if cover.get('teacher'):
            add_text(sh, SW * 0.06, SH * 0.77, SW * 0.62, SH * 0.05,
                     cover['teacher'], 12, color=TXT_DIM)
        add_gif_pair(sh, SW * 0.80, SH * 0.40, SW * 0.20)
        add_footer(sh, 1, total, footer_left, footer_name)
        set_morph_transition(s, dur_ms=morph_ms)

    cx, cy = int(SW * 0.06), int(SH * 0.50)
    r = int(SH * 0.56)
    for i, pg in enumerate(pages):
        s = prs.slides.add_slide(blank)
        set_slide_bg_picture(s, bg)
        sh = s.shapes
        add_gear_roulette(sh, cx, cy, r, i * step, labels, i, icons)
        add_content(sh, pg.get('part', labels[i]), pg.get('section', ''),
                    pg.get('title', ''), pg.get('points', []))
        add_gif_pair(sh, SW * 0.885, SH * 0.15, SW * 0.07)
        add_footer(sh, i + total - 3, total, footer_left, footer_name, left_x=SW * 0.30)
        set_morph_transition(s, dur_ms=morph_ms)

    prs.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser(description='机械风大齿轮转盘 PPT 生成器')
    ap.add_argument('-o', '--out', default='zhuanpan.pptx')
    ap.add_argument('--config', default=None, help='JSON 配置（cover/pages/footer/morph_ms/direction）')
    args = ap.parse_args()
    cfg = DEMO
    if args.config:
        with open(args.config, encoding='utf-8') as f:
            cfg = json.load(f)
    out = build(args.out, cfg)
    print('saved:', os.path.abspath(out), os.path.getsize(out), 'bytes | pages:',
          len(cfg['pages']) + (1 if cfg.get('cover') else 0),
          '| direction:', cfg.get('direction', 'ccw'))


if __name__ == '__main__':
    main()
