# -*- coding: utf-8 -*-
"""ppt-xqmsaf-skill — 喜鹊谋杀案风（暗场·酒红·暖光）PPT 生成器。

复刻自抖音教程视频《喜鹊谋杀案风 PPT》的三大手法：

1. 灯光
   - 灯源：一大一小两个淡黄圆叠出灯泡
   - 光锥：PIE 扇形（"被咬了一口的饼形"），同色、透明度 45%、柔化边缘
   - 红蒙版：满页红色矩形 + 47% 透明度，压出 #450001 深酒红底色
   - 光照图片：图片与光锥"合并形状→相交"的等价实现 ——
     用 Pillow 按光锥几何把图片裁成带羽化 alpha 的 PNG，垫在暗图上层
   - 开灯动画：光锥 + 压扁光晕 + 光照图，擦除(自顶部)、0.5s、同时播放
   - 聚光灯转移：跨页同名对象 + 文字透明度 74% → 0%，灯动文字亮
2. 图片处理
   - 直接接收 AI 生成/处理好的同风格图片（视频里用"百度搭子"）
   - 暗底色 + 满页红蒙版统一色调
3. 页面衔接
   - 全部内容页挂 Morph（平滑切换），跨页同名对象自动补间
   - "拉大图片 + 移出画面外"配方：上一页把大图放到画外，下一页摆正，
     播放时大图从画外丝滑滑入

用法：
  python build_xqmsaf.py -o out.pptx                       # 内置演示（无图片）
  python build_xqmsaf.py -o out.pptx --config demo.json    # 相对路径以配置文件为基准

依赖：python-pptx >= 0.6.19、Pillow。
"""
import argparse
import json
import math
import os
import sys

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn, nsdecls
from PIL import Image, ImageDraw, ImageFilter

SW, SH = 12192000, 6858000        # 16:9 页面
PT = 12700                        # 1pt = 12700 EMU
EMU_PX = 9525                     # 96dpi 下 1px = 9525 EMU

MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
P14_NS = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
P159_NS = 'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"'

# 视频实测配色：黑底 + 红 47% 蒙版 → #450001 深酒红；暖白文字；淡黄灯光
PALETTE = {
    'bg': '000000',        # 底色（黑）
    'wash': 'A50004',      # 满页红蒙版
    'lamp': 'FFF2CC',      # 灯的淡黄（近 PPT 标准淡黄）
    'cream': 'E8D2CF',     # 暖白正文/标题
    'dim': '8E6A62',       # 弱化文字（暖灰红）
    'accent': 'B2554A',    # 小标签/短线的砖红
    'red': '9E0007',       # 强调深红
}

DEFAULTS = {
    'font_title': 'STZhongsong',   # 华文中宋（没有就自行换成 宋体 等衬线）
    'font_body': '宋体',
    'font_en': 'Georgia',
    'morph_ms': 1400,              # 平滑切换时长
    'wipe_ms': 500,                # 开灯擦除动画时长
    'bars_ms': 800,                # 随机线条动画时长
    'wash_alpha': 47,              # 红蒙版透明度（%）
    'cone_alpha': 45,              # 光锥透明度（%）
    'feather': 0.035,              # 光照图羽化（占页高比例）
    'palette': {},
    'pages': [],
}

# 各页型的灯光/光锥默认几何（lamp = [x, y, r]，均为占页宽/页高的比例；r 相对页高）
TYPE_DEFAULTS = {
    'cover': {
        'lamp': (0.705, 0.255, 0.078),
        'cone': {'dir': 118, 'spread': 66, 'length': 1.55},
        'glow': True, 'light_on': True,
    },
    'section': {
        'lamp': (0.80, 0.285, 0.062),
        'cone': {'dir': 148, 'spread': 60, 'length': 1.5},
        'glow': True, 'light_on': True,
    },
    'content': {
        'lamp': (0.285, 0.265, 0.055),
        'cone': {'dir': 15, 'spread': 50, 'length': 1.0},
        'glow': True, 'light_on': False,
    },
    'toc': {
        'lamp': (0.285, 0.265, 0.055),
        'cone': {'dir': 15, 'spread': 50, 'length': 0.02},   # 灯先藏进灯源
        'glow': True, 'light_on': False,
    },
    'end': {
        'lamp': (0.50, 0.40, 0.048),
        'cone': {'dir': 90, 'spread': 52, 'length': 0.02},
        'glow': False, 'light_on': False,
    },
}
TYPES = tuple(TYPE_DEFAULTS)

# 各页型的图片默认摆放（未写 image_box 时生效）：toc 默认把大图放在画外，
# 作为"下一内容页大图滑入"的衔接起点。
TYPE_IMAGE_BOX = {'toc': [0.90, 0.085, 0.42, 0.83]}

ALLOWED_TOP = set(DEFAULTS)
ALLOWED_PAGE = {'type', 'kicker', 'title', 'en', 'sub', 'body', 'num', 'image',
                'image_box', 'lamp', 'cone', 'glow', 'light_on', 'title_reveal',
                'text_transparency', 'title_size', 'body_size', 'foot', 'lit'}
ALLOWED_CONE = {'dir', 'spread', 'length'}
ALLOWED_PALETTE = set(PALETTE)

_FONT_TITLE = DEFAULTS['font_title']
_FONT_BODY = DEFAULTS['font_body']
_FONT_EN = DEFAULTS['font_en']
_PAL = dict(PALETTE)


# ---------------------------------------------------------------- xml helpers

def _set_solid(shape, hexcolor, opacity=None):
    """纯色填充；opacity 为不透明度 %（None = 100%）。"""
    spPr = shape._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
                'a:pattFill', 'a:grpFill'):
        el = spPr.find(qn(tag))
        if el is not None:
            spPr.remove(el)
    if opacity is None or opacity >= 100:
        xml = ('<a:solidFill %s><a:srgbClr val="%s"/></a:solidFill>'
               % (nsdecls('a'), hexcolor))
    else:
        xml = ('<a:solidFill %s><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr>'
               '</a:solidFill>' % (nsdecls('a'), hexcolor, int(opacity * 1000)))
    spPr.find(qn('a:prstGeom')).addnext(parse_xml(xml))


def _no_line(shape):
    shape.line.fill.background()


def _soft_edge(shape, rad_pt):
    """柔化边缘（视频里的"把它的角给磨平"）。"""
    spPr = shape._element.spPr
    old = spPr.find(qn('a:effectLst'))
    if old is not None:
        spPr.remove(old)
    xml = ('<a:effectLst %s><a:softEdge rad="%d"/></a:effectLst>'
           % (nsdecls('a'), max(0, int(rad_pt * PT))))
    spPr.append(parse_xml(xml))


def _run_style(run, size, color_hex, font, bold=False, italic=False,
               transparency=None, spacing_pt=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    run.font.color.rgb = rgb(color_hex)
    if spacing_pt is not None:
        run.font._rPr.set('spc', str(int(spacing_pt * 100)))
    if transparency:
        solid = run.font._rPr.find(qn('a:solidFill'))
        if solid is not None:
            clr = solid.find(qn('a:srgbClr'))
            if clr is not None:
                alpha = clr.makeelement(
                    qn('a:alpha'), {'val': str(int((100 - transparency) * 1000))})
                clr.append(alpha)


def set_morph_transition(slide, dur_ms=1400):
    """Morph（平滑）切换：Choice 走 p159:morph byObject，Fallback 普通切换。"""
    sld = slide._element
    MC_TAG = '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
    for child in list(sld):
        if child.tag in (MC_TAG, qn('p:transition')):
            sld.remove(child)
    xml = ('<mc:AlternateContent %s %s %s><mc:Choice %s Requires="p159">'
           '<p:transition spd="slow" p14:dur="%d"><p159:morph option="byObject"/>'
           '</p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"/>'
           '</mc:Fallback></mc:AlternateContent>') % (MC_NS, nsdecls('p'), P14_NS,
                                                      P159_NS, dur_ms)
    sld.append(parse_xml(xml))


def set_timing(slide, groups):
    """写入 <p:timing>。

    groups: [{'delay': ms, 'effects': [{'spid': int, 'kind': 'wipe'|'bars',
                                        'dur': ms, 'delay': ms}]}]
    group 全部自动播放（onBegin of mainSeq），组内第一个 afterEffect、其余 withEffect。
    wipe 用 擦除·自顶部（wipe(down)/subtype 4）；bars 用 随机线条 水平。
    """
    if not groups:
        return
    nid = [2]                     # 1=tmRoot、2=mainSeq，其余从 3 起，绝不能重复

    def next_id():
        nid[0] += 1
        return nid[0]

    KINDS = {'wipe': (22, 4, 'wipe(down)'), 'bars': (14, 10, 'randombar(horizontal)')}
    body = []
    for gi, group in enumerate(groups):
        inner = []
        for ei, eff in enumerate(group['effects']):
            pid, sub, filt = KINDS[eff['kind']]
            node = 'afterEffect' if ei == 0 else 'withEffect'
            if eff['kind'] == 'bars':      # 文本按段落目标（对齐 PowerPoint 原生写法）
                tgt = ('<p:spTgt spid="%d"><p:txEl><p:pRg st="0" end="0"/></p:txEl>'
                       '</p:spTgt>' % eff['spid'])
            else:                          # 形状目标
                tgt = ('<p:spTgt spid="%d"><p:bg/></p:spTgt>' % eff['spid'])
            inner.append(
                '<p:par><p:cTn id="%d" presetID="%d" presetClass="entr" '
                'presetSubtype="%d" fill="hold" grpId="0" nodeType="%s">'
                '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
                '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
                '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                '<p:tgtEl>%s</p:tgtEl>'
                '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                '</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                '<p:animEffect transition="in" filter="%s"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d"/><p:tgtEl>%s</p:tgtEl>'
                '</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>'
                % (next_id(), pid, sub, node, int(eff.get('delay', 0)), next_id(),
                   tgt, filt, next_id(), int(eff['dur']), tgt))
        body.append(
            '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst>'
            '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="%d">'
            '<p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="0"/>'
            '</p:stCondLst><p:childTnLst>%s</p:childTnLst></p:cTn></p:par>'
            '</p:childTnLst></p:cTn></p:par>'
            % (next_id(), int(group.get('delay', 0)), next_id(), ''.join(inner)))

    xml = ('<p:timing %s><p:tnLst><p:par><p:cTn id="1" dur="indefinite" '
           'restart="never" nodeType="tmRoot"><p:childTnLst>'
           '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" '
           'nodeType="mainSeq"><p:childTnLst>%s</p:childTnLst></p:cTn>'
           '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/>'
           '</p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" '
           'delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
           '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>') % (
        nsdecls('p'), ''.join(body))
    sld = slide._element
    old = sld.find(qn('p:timing'))
    if old is not None:
        sld.remove(old)
    sld.append(parse_xml(xml))


# ---------------------------------------------------------------- utils

def rgb(hexstr):
    if hexstr is None:
        return None
    h = str(hexstr).strip().lstrip('#')
    if not h:
        return None
    if len(h) == 3:
        h = ''.join(c * 2 for c in h)
    if len(h) != 6 or any(c not in '0123456789abcdefABCDEF' for c in h):
        raise ValueError('颜色要写成 6 位十六进制（如 E8D2CF），收到 %r' % hexstr)
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def resolve_path(path, base_dir=None):
    if not path:
        return None
    p = str(path)
    if base_dir and not os.path.isabs(p):
        p = os.path.join(base_dir, p)
    return os.path.normpath(p)


def _cover_crop_fractions(img_path, box_w, box_h):
    """按"cover"方式填满框所需裁剪比例（返回 l/t/r/b，0~1）。"""
    with Image.open(img_path) as im:
        iw, ih = im.size
    src = iw / float(ih)
    dst = box_w / float(box_h)
    if src > dst:                      # 图更宽 → 左右裁
        f = (1.0 - dst / src) / 2.0
        return f, 0.0, f, 0.0
    if src < dst:                      # 图更高 → 上下裁
        f = (1.0 - src / dst) / 2.0
        return 0.0, f, 0.0, f
    return 0.0, 0.0, 0.0, 0.0


def make_lit_picture(img_path, box, lamp, cone_dir, cone_spread, cone_length,
                     feather, out_dir):
    """光照图：把图片裁成"光锥范围内可见"的 RGBA PNG（≈ 合并形状→相交 + 柔化）。

    box             (x, y, w, h)，EMU，图片在页面上的摆放框
    lamp            (x, y)，EMU，灯心（可位于框外）
    cone_dir        度，0=右，顺时针
    cone_spread     度，扇形张角
    cone_length     EMU，光锥长度
    feather         羽化半径占页高比例
    """
    bx, by, bw, bh = box
    W = max(320, min(2600, int(round(bw / float(EMU_PX) * 2))))   # 2x 超采样
    H = max(200, min(2600, int(round(bh / float(EMU_PX) * 2))))
    with Image.open(img_path) as im:
        im = im.convert('RGB')
        iw, ih = im.size
        scale = max(W / float(iw), H / float(ih))
        nw, nh = int(math.ceil(iw * scale)), int(math.ceil(ih * scale))
        im = im.resize((nw, nh), Image.LANCZOS)
        left = (nw - W) // 2
        top = (nh - H) // 2
        im = im.crop((left, top, left + W, top + H))

    px_per_emu = W / float(bw)
    lx = (lamp[0] - bx) * px_per_emu
    ly = (lamp[1] - by) * px_per_emu
    R = cone_length * px_per_emu

    mask = Image.new('L', (W, H), 0)
    d = ImageDraw.Draw(mask)
    pts = [(lx, ly)]
    a0 = cone_dir - cone_spread / 2.0
    a1 = cone_dir + cone_spread / 2.0
    steps = 48
    for k in range(steps + 1):
        a = math.radians(a0 + (a1 - a0) * k / steps)
        pts.append((lx + R * math.cos(a), ly + R * math.sin(a)))
    d.polygon(pts, fill=255)
    feather_px = max(2.0, feather * SH / float(bh) * H)
    mask = mask.filter(ImageFilter.GaussianBlur(feather_px))
    out = Image.new('RGBA', (W, H))
    out.paste(im, (0, 0))
    out.putalpha(mask)
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, 'lit_%dx%d_%d.png' % (W, H, abs(hash(
        (img_path, box, lamp, cone_dir, cone_spread, cone_length)))))
    out.save(path)
    return path


# ---------------------------------------------------------------- shape helpers

def add_text(shapes, x, y, w, h, text, size, color_hex, font=None, bold=False,
             italic=False, align=PP_ALIGN.LEFT, transparency=None,
             spacing_pt=None, line_spacing=None, name=None,
             anchor=MSO_ANCHOR.TOP):
    box = shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = str(text).split('\n') if text is not None else ['']
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing is not None:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        _run_style(run, size, color_hex, font or _FONT_BODY, bold=bold,
                   italic=italic, transparency=transparency,
                   spacing_pt=spacing_pt)
    return box


def add_rect(shapes, x, y, w, h, color_hex, opacity=None, name=None):
    sp = shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x)), Emu(int(y)),
                          Emu(int(w)), Emu(int(h)))
    if name:
        sp.name = name
    _set_solid(sp, color_hex, opacity)
    _no_line(sp)
    sp.shadow.inherit = False
    return sp


def add_lamp(shapes, cx, cy, r):
    """一大一小两个淡黄圆叠成灯源（视频：插入圆形→淡黄→无线条→复制→调整大小）。"""
    big = shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx - r)), Emu(int(cy - r)),
                           Emu(int(2 * r)), Emu(int(2 * r)))
    big.name = 'xq-lamp-a'
    _set_solid(big, _PAL['lamp'], opacity=62)
    _no_line(big)
    _soft_edge(big, max(1.0, r / float(PT) * 0.30))
    big.shadow.inherit = False

    rs = r * 0.55
    off = r * 0.16
    small = shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx - off - rs)),
                             Emu(int(cy - off - rs)), Emu(int(2 * rs)), Emu(int(2 * rs)))
    small.name = 'xq-lamp-b'
    _set_solid(small, _PAL['lamp'], opacity=96)
    _no_line(small)
    _soft_edge(small, max(0.8, rs / float(PT) * 0.25))
    small.shadow.inherit = False
    return big, small


def add_cone(shapes, cx, cy, radius, direction, spread):
    """光锥：PIE 扇形，指向 direction，张角 spread（扇形默认朝右，旋转对齐）。"""
    pie = shapes.add_shape(MSO_SHAPE.PIE, Emu(int(cx - radius)),
                           Emu(int(cy - radius)), Emu(int(2 * radius)),
                           Emu(int(2 * radius)))
    pie.name = 'xq-cone'
    pie.adjustments[0] = -spread / 2.0
    pie.adjustments[1] = spread / 2.0
    pie.rotation = direction % 360
    _set_solid(pie, _PAL['lamp'], opacity=100 - 45)   # 透明度由调用方覆盖
    _no_line(pie)
    pie.shadow.inherit = False
    return pie


def add_glow(shapes, cx, cy, r, direction):
    """压扁拉伸的光晕（视频里"给它做下按摩"的扁圆），沿光锥方向铺开。"""
    gw = r * 4.6
    gh = r * 1.15
    sp = shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx - gw / 2)),
                          Emu(int(cy - gh / 2)), Emu(int(gw)), Emu(int(gh)))
    sp.name = 'xq-glow'
    sp.rotation = direction % 360
    _set_solid(sp, _PAL['lamp'], opacity=26)
    _no_line(sp)
    _soft_edge(sp, max(1.0, gh / float(PT) * 0.55))
    sp.shadow.inherit = False
    return sp


# ---------------------------------------------------------------- page builders

def _draw_base(slide, page, cfg):
    """底图 → 暗底 → 红蒙版 → 光照图 → 光锥 → 光晕 → 灯源。返回绘制结果。"""
    sh = slide.shapes
    add_rect(sh, 0, 0, SW, SH, _PAL['bg'], name='xq-bg')

    hero = None
    box = None
    img_path = page.get('_image_abs')
    if img_path:
        if not os.path.isfile(img_path):
            raise ValueError('图片不存在：%s' % img_path)
        box = _page_box(page, cfg)
        hero = sh.add_picture(img_path, Emu(int(box[0])), Emu(int(box[1])),
                              Emu(int(box[2])), Emu(int(box[3])))
        hero.name = 'xq-hero'
        cl, ct, cr, cb = _cover_crop_fractions(img_path, box[2], box[3])
        hero.crop_left, hero.crop_top = cl, ct
        hero.crop_right, hero.crop_bottom = cr, cb
        hero.shadow.inherit = False

    add_rect(sh, 0, 0, SW, SH, _PAL['wash'],
             opacity=100 - cfg['wash_alpha'], name='xq-wash')

    lit = None
    lamp = _lamp_emu(page)
    cone = dict(TYPE_DEFAULTS[page['type']]['cone'])
    cone.update(page.get('cone') or {})
    if page.get('lit') and img_path and box:
        lit_path = _lit_of(page, cfg)
        lit = sh.add_picture(lit_path, Emu(int(box[0])), Emu(int(box[1])),
                             Emu(int(box[2])), Emu(int(box[3])))
        lit.name = 'xq-lit'

    length = float(cone['length']) * SW
    radius = max(int(SW * 0.01), int(length))
    pie = add_cone(sh, lamp[0], lamp[1], radius,
                   float(cone['dir']), float(cone['spread']))
    _set_solid(pie, _PAL['lamp'], opacity=100 - cfg['cone_alpha'])
    _soft_edge(pie, min(120.0, max(2.0, radius / float(PT) * 0.06)))

    glow = None
    if page.get('glow', TYPE_DEFAULTS[page['type']]['glow']):
        glow = add_glow(sh, lamp[0], lamp[1], lamp[2], float(cone['dir']))
    else:
        glow = add_glow(sh, lamp[0], lamp[1], lamp[2] * 0.16, float(cone['dir']))
        glow.name = 'xq-glow'      # 仍保留同名小光晕（Morph 跨页补间用）

    add_lamp(sh, lamp[0], lamp[1], lamp[2])
    return {'hero': hero, 'lit': lit, 'cone': pie, 'glow': glow,
            'lamp': lamp, 'box': box, 'img': img_path}


def _lamp_emu(page):
    lx, ly, lr = page['lamp']
    return int(SW * float(lx)), int(SH * float(ly)), int(SH * float(lr))


def _page_box(page, cfg):
    b = (page.get('image_box') or TYPE_IMAGE_BOX.get(page['type'])
         or [0.52, 0.085, 0.42, 0.83])
    if isinstance(b, dict):
        b = [b.get('x', 0.52), b.get('y', 0.085), b.get('w', 0.42), b.get('h', 0.83)]
    return (int(SW * float(b[0])), int(SH * float(b[1])),
            int(SW * float(b[2])), int(SH * float(b[3])))


def _lit_of(page, cfg):
    box = _page_box(page, cfg)
    lamp = _lamp_emu(page)
    cone = dict(TYPE_DEFAULTS[page['type']]['cone'])
    cone.update(page.get('cone') or {})
    return make_lit_picture(page['_image_abs'], box, (lamp[0], lamp[1]),
                            float(cone['dir']), float(cone['spread']),
                            float(cone['length']) * SW, float(cfg['feather']),
                            page['_tmp_dir'])


def _texts_cover(sh, page):
    add_text(sh, SW * 0.070, SH * 0.150, SW * 0.45, SH * 0.045,
             page.get('kicker') or 'PRESENTATION', 12, _PAL['accent'],
             font=_FONT_EN, bold=True, spacing_pt=5, name='xq-kicker')
    title = page.get('title') or '标题'
    size = page.get('title_size') or (52 if max(len(x) for x in title.split('\n')) <= 9 else 44)
    add_text(sh, SW * 0.067, SH * 0.205, SW * 0.62, SH * 0.26, title, size,
             _PAL['cream'], font=_FONT_TITLE, bold=True, line_spacing=1.05,
             name='xq-title')
    add_rect(sh, SW * 0.072, SH * 0.505, SW * 0.072, int(2.4 * PT),
             _PAL['red'], name='xq-rule')
    add_text(sh, SW * 0.070, SH * 0.538, SW * 0.58, SH * 0.05,
             page.get('en') or '', 15, _PAL['cream'], font=_FONT_EN, italic=True,
             transparency=30, name='xq-en')
    add_text(sh, SW * 0.070, SH * 0.628, SW * 0.55, SH * 0.08,
             page.get('sub') or '', 13, _PAL['dim'], name='xq-sub')


def _texts_section(sh, page):
    num = page.get('num') or 'Ⅰ'
    add_text(sh, SW * 0.068, SH * 0.135, SW * 0.42, SH * 0.30, num,
             page.get('title_size') or 108, _PAL['red'],
             font=_FONT_TITLE, bold=True, transparency=45, name='xq-num')
    add_rect(sh, SW * 0.073, SH * 0.492, SW * 0.062, int(2.2 * PT),
             _PAL['red'], name='xq-rule')
    add_text(sh, SW * 0.070, SH * 0.512, SW * 0.5, SH * 0.04,
             page.get('kicker') or '', 12, _PAL['accent'], font=_FONT_EN,
             bold=True, spacing_pt=5, name='xq-kicker')
    add_text(sh, SW * 0.067, SH * 0.560, SW * 0.60, SH * 0.12,
             page.get('title') or '', page.get('title_size') or 40,
             _PAL['cream'], font=_FONT_TITLE, bold=True, name='xq-title')
    add_text(sh, SW * 0.070, SH * 0.690, SW * 0.55, SH * 0.05,
             page.get('en') or '', 14, _PAL['cream'], font=_FONT_EN,
             italic=True, transparency=40, name='xq-en')


def _texts_content(sh, page, transparency):
    add_text(sh, SW * 0.065, SH * 0.115, SW * 0.38, SH * 0.04,
             page.get('kicker') or '', 11, _PAL['accent'], font=_FONT_EN,
             bold=True, spacing_pt=4, name='xq-kicker')
    title = page.get('title') or ''
    size = page.get('title_size') or (34 if max(len(x) for x in title.split('\n')) <= 8 else 28)
    add_text(sh, SW * 0.062, SH * 0.160, SW * 0.41, SH * 0.20, title, size,
             _PAL['cream'], font=_FONT_TITLE, bold=True, line_spacing=1.12,
             transparency=transparency, name='xq-title')
    add_rect(sh, SW * 0.067, SH * 0.395, SW * 0.052, int(2.0 * PT),
             _PAL['red'], name='xq-rule')
    add_text(sh, SW * 0.065, SH * 0.425, SW * 0.38, SH * 0.045,
             page.get('en') or '', 13, _PAL['cream'], font=_FONT_EN, italic=True,
             transparency=35 + (transparency or 0) // 2, name='xq-en')
    body = page.get('body') or []
    body_text = '\n'.join(body) if isinstance(body, list) else str(body)
    add_text(sh, SW * 0.065, SH * 0.500, SW * 0.385, SH * 0.33, body_text,
             page.get('body_size') or (15 if page['type'] == 'toc' else 13),
             _PAL['cream'], line_spacing=1.55,
             transparency=transparency, name='xq-body')


def _texts_end(sh, page):
    add_text(sh, SW * 0.18, SH * 0.330, SW * 0.64, SH * 0.13,
             page.get('title') or 'THE END', page.get('title_size') or 44,
             _PAL['cream'], font=_FONT_TITLE, bold=True, align=PP_ALIGN.CENTER,
             name='xq-title')
    add_rect(sh, SW * 0.468, SH * 0.492, SW * 0.064, int(2.0 * PT),
             _PAL['red'], name='xq-rule')
    add_text(sh, SW * 0.20, SH * 0.530, SW * 0.60, SH * 0.06,
             page.get('sub') or '', 15, _PAL['cream'], align=PP_ALIGN.CENTER,
             transparency=45, name='xq-sub')
    add_text(sh, SW * 0.20, SH * 0.605, SW * 0.60, SH * 0.05,
             page.get('en') or '', 12, _PAL['cream'], font=_FONT_EN, italic=True,
             align=PP_ALIGN.CENTER, transparency=60, name='xq-en')


def _texts(sh, page, transparency=None):
    t = page['type']
    if t == 'cover':
        _texts_cover(sh, page)
    elif t == 'section':
        _texts_section(sh, page)
    elif t in ('content', 'toc'):
        _texts_content(sh, page, transparency if transparency is not None
                       else page.get('text_transparency') or 0)
    elif t == 'end':
        _texts_end(sh, page)


def _add_footer(sh, idx, total):
    add_text(sh, SW * 0.62, SH * 0.938, SW * 0.345, SH * 0.04,
             '%02d / %02d' % (idx, total), 9, _PAL['dim'],
             font=_FONT_EN, align=PP_ALIGN.RIGHT, name='xq-foot')


def render_page(slide, idx, total, page, cfg):
    """单页：底图/蒙版/灯光/文字 + Morph + 动画。"""
    t = page['type']
    _draw_base(slide, page, cfg)

    trans = page.get('text_transparency')
    if t in ('content', 'toc') and trans is None:
        trans = 0
    _texts(slide.shapes, page, trans)
    _add_footer(slide.shapes, idx, total)

    if idx > 1:
        set_morph_transition(slide, dur_ms=cfg['morph_ms'])

    groups = []
    if page.get('light_on', TYPE_DEFAULTS[t]['light_on']):
        effects = []
        shapes = {sp.name: sp for sp in slide.shapes}
        for nm in ('xq-cone', 'xq-glow', 'xq-lit'):
            if nm in shapes:
                effects.append({'spid': shapes[nm].shape_id, 'kind': 'wipe',
                                'dur': cfg['wipe_ms']})
        if effects:
            groups.append({'delay': 0, 'effects': effects})
    if page.get('title_reveal') == 'bars':
        shapes = {sp.name: sp for sp in slide.shapes}
        if 'xq-title' in shapes:
            groups.append({'delay': cfg['wipe_ms'],
                           'effects': [{'spid': shapes['xq-title'].shape_id,
                                        'kind': 'bars', 'dur': cfg['bars_ms']}]})
    if groups:
        set_timing(slide, groups)


# ---------------------------------------------------------------- config

def resolve_config(cfg):
    if not isinstance(cfg, dict):
        raise ValueError('配置必须是 JSON 对象')
    unknown = sorted(set(cfg) - ALLOWED_TOP)
    if unknown:
        raise ValueError('未知配置项：%s' % ', '.join(unknown))
    out = dict(DEFAULTS)
    out.update(cfg)
    pal = dict(PALETTE)
    overrides = out.get('palette') or {}
    if not isinstance(overrides, dict):
        raise ValueError('palette 必须是对象')
    bad = sorted(set(overrides) - ALLOWED_PALETTE)
    if bad:
        raise ValueError('未知配色项：%s（可用：%s）'
                         % (', '.join(bad), ', '.join(sorted(ALLOWED_PALETTE))))
    pal.update(overrides)
    for value in pal.values():
        rgb(value)
    out['palette'] = pal
    for key in ('morph_ms', 'wipe_ms', 'bars_ms', 'wash_alpha', 'cone_alpha'):
        try:
            out[key] = int(out[key])
        except (TypeError, ValueError):
            raise ValueError('%s 必须是整数' % key)
    if not 0 <= out['wash_alpha'] <= 95:
        raise ValueError('wash_alpha 需在 0~95 之间（47 = 视频同款深酒红）')
    if not 0 <= out['cone_alpha'] <= 95:
        raise ValueError('cone_alpha 需在 0~95 之间（45 = 视频同款柔光）')
    try:
        out['feather'] = float(out['feather'])
    except (TypeError, ValueError):
        raise ValueError('feather 必须是数字（占页高比例，如 0.035）')
    pages = out.get('pages') or []
    if not isinstance(pages, list) or not pages:
        raise ValueError('pages 必须是非空数组')
    if len(pages) > 12:
        raise ValueError('页数最多 12 页，当前 %d' % len(pages))
    merged = []
    for i, pg in enumerate(pages):
        if not isinstance(pg, dict):
            raise ValueError('pages[%d] 必须是对象' % i)
        bad = sorted(set(pg) - ALLOWED_PAGE)
        if bad:
            raise ValueError('pages[%d] 未知字段：%s' % (i, ', '.join(bad)))
        if pg.get('type') not in TYPES:
            raise ValueError('pages[%d].type 只能是 %s，收到 %r'
                             % (i, ' / '.join(TYPES), pg.get('type')))
        base = {'type': pg['type'],
                'lamp': list(TYPE_DEFAULTS[pg['type']]['lamp']),
                'cone': dict(TYPE_DEFAULTS[pg['type']]['cone']),
                'glow': TYPE_DEFAULTS[pg['type']]['glow'],
                'light_on': TYPE_DEFAULTS[pg['type']]['light_on']}
        base.update(pg)
        if isinstance(pg.get('lamp'), (list, tuple)) and len(pg['lamp']) == 3:
            base['lamp'] = [float(v) for v in pg['lamp']]
        cone = dict(TYPE_DEFAULTS[pg['type']]['cone'])
        if pg.get('cone'):
            if not isinstance(pg['cone'], dict):
                raise ValueError('pages[%d].cone 必须是对象' % i)
            bad = sorted(set(pg['cone']) - ALLOWED_CONE)
            if bad:
                raise ValueError('pages[%d].cone 未知字段：%s' % (i, ', '.join(bad)))
            cone.update(pg['cone'])
        base['cone'] = {k: float(v) for k, v in cone.items()}
        merged.append(base)
    out['pages'] = merged
    return out


DEMO = {
    'pages': [
        {'type': 'cover',
         'kicker': 'PRESENTATION · XQMSAF',
         'title': '雾锁重塔',
         'en': 'PAGODAS IN THE FOG',
         'sub': '把悬疑写进每一页：暗场、灯火与平滑的转场。'},
        {'type': 'section', 'num': 'Ⅰ', 'kicker': 'PART ONE',
         'title': '灯与影', 'en': 'THE LAMP & THE SHADOW'},
        {'type': 'toc',
         'kicker': 'Ⅱ · 目录', 'title': '三座塔',
         'en': 'Three Pagodas',
         'body': ['Ⅰ　燃灯塔 · Randeng Pagoda',
                  'Ⅱ　铁塔 · Iron Pagoda',
                  'Ⅲ　瑞光塔 · Ruiguang Pagoda']},
        {'type': 'content',
         'kicker': 'Ⅲ · 其一', 'title': '燃灯塔',
         'en': 'Randeng Buddha Relic Pagoda',
         'body': ['木架撑起层层飞檐，', '像一只收拢翅膀的鸟。',
                  '灰瓦覆着倾斜的屋顶，', '灯火在雾里亮了一千年。'],
         'text_transparency': 74},
        {'type': 'content',
         'kicker': 'Ⅲ · 其二', 'title': '铁塔',
         'en': 'Iron Pagoda · Youguo Temple',
         'body': ['十三层铁色琉璃，', '风吹过时像翻书的声音。',
                  '灯光移过来，', '故事就翻到了下一页。']},
        {'type': 'end',
         'title': '第一幕 · 终',
         'sub': '灯光藏进灯源，故事留给下一页。',
         'en': 'END OF ACT ONE'},
    ],
}


def build(out_path, cfg, base_dir=None):
    cfg = resolve_config(cfg)
    global _FONT_TITLE, _FONT_BODY, _FONT_EN, _PAL
    _FONT_TITLE, _FONT_BODY, _FONT_EN = (cfg['font_title'], cfg['font_body'],
                                         cfg['font_en'])
    _PAL = cfg['palette']

    tmp_dir = os.path.join(os.path.dirname(os.path.abspath(out_path)) or '.',
                           '.xqmsaf_tmp')
    prs = Presentation()
    prs.slide_width = Emu(SW)
    prs.slide_height = Emu(SH)
    blank = prs.slide_layouts[6]
    total = len(cfg['pages'])
    for i, page in enumerate(cfg['pages']):
        if page.get('image'):
            page['_image_abs'] = resolve_path(page['image'], base_dir)
            page['_tmp_dir'] = tmp_dir
        s = prs.slides.add_slide(blank)
        render_page(s, i + 1, total, page, cfg)
    prs.save(out_path)

    # 清理光照图临时文件
    try:
        if os.path.isdir(tmp_dir):
            for f in os.listdir(tmp_dir):
                os.remove(os.path.join(tmp_dir, f))
            os.rmdir(tmp_dir)
    except OSError:
        pass
    return out_path


def main():
    ap = argparse.ArgumentParser(description='喜鹊谋杀案风 PPT 生成器')
    ap.add_argument('-o', '--out', default='xqmsaf.pptx')
    ap.add_argument('--config', default=None,
                    help='JSON 配置；相对路径以配置文件所在目录为基准')
    args = ap.parse_args()

    cfg = dict(DEMO)
    base_dir = None
    if args.config:
        with open(args.config, encoding='utf-8') as f:
            cfg = json.load(f)
        base_dir = os.path.dirname(os.path.abspath(args.config))

    try:
        out = build(args.out, cfg, base_dir=base_dir)
    except Exception as exc:                            # noqa: BLE001
        print('ERROR: %s' % exc)
        return 1

    types = [p.get('type') for p in cfg.get('pages') or []]
    print('saved: %s | %d bytes | pages: %d | types: %s'
          % (os.path.abspath(out), os.path.getsize(out), len(types),
             ', '.join(types)))
    return 0


if __name__ == '__main__':
    sys.exit(main())