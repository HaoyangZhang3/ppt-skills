# -*- coding: utf-8 -*-
"""ppt-zhuanpan-skill — 翻页旋转转盘 PPT 生成器（自包含，风格无关）。

这个技能只做一件事：页面左侧放一个大转盘，只露右半边，每翻一页转盘旋转
360/N 度（N = 章数），当前章的标签永远落在右缘 3 点钟方向、正对右侧内容区；
标签旁的图标在翻页时缩回轮心 / 放大伸出接力。全程 Morph（平滑）切换按对象名
补间，所以转盘是真的在转，不是换图。

三种造型：
  disc    中性圆盘（纯色/渐变 + 分格线 + 轴孔）
  gear9   PowerPoint 自带 9 齿齿轮 + 轴孔
  double  内外双盘：外环 + 内盘各一个圆，同一张照片同时填进两个盘，
          翻页时外环逆时针、内盘顺时针反向转；标签沿圆周切向排布（随盘转，
          转到下方会倒过来），分格线只出现在环带上。

造型、章数、配色、照片、背景、图标全部由配置决定，默认中性无风格。
技能不附带任何图片素材（double 的照片必须由配置给路径）。

用法：
  python zhuanpan_deck.py -o out.pptx                    # 内置演示（4 章 + 封面）
  python zhuanpan_deck.py -o out.pptx --config deck.json

依赖：python-pptx >= 0.6.19
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

BASE = os.path.dirname(os.path.abspath(__file__))
PT = 12700
EMU_IN = 914400
SW, SH = 12192000, 6858000  # 16:9

MIN_CHAPTERS, MAX_CHAPTERS = 2, 12
WHEELS = ('disc', 'gear9', 'double')
DIRECTIONS = ('ccw', 'cw')
LABEL_MODES = ('flat', 'tangential')

MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
P14_NS = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
P159_NS = 'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"'

# 默认中性配色（白底可读）。config 的 colors 可逐项覆盖；写成 '' 表示该项不用
# （例如 wheel_line: '' → 转盘不描边，spoke: '' → 不画分格线）。
PALETTE = {
    'wheel_fill': 'E6E9ED',   # 转盘填充
    'wheel_fill2': '',        # 非空则与 wheel_fill 组成 45° 渐变
    'wheel_line': 'C0C7CE',   # 转盘描边
    'spoke': 'AEB6BE',        # 分格放射线
    'hub_fill': 'FFFFFF',     # 中心轴孔（必须不透明，要盖住缩小的图标）
    'hub_line': 'AEB6BE',
    'ring_line': '',          # double：外环描边
    'inner_line': '',         # double：内盘描边
    'label': '1F2429',        # 标签前缀（part）
    'label_num': '6A727A',    # 标签编号
    'title': '1F2429',        # 内容区大标题
    'text': '4E565E',         # 内容区要点
    'accent': '5A7186',       # 小标签、短横线、要点方块
    'dim': '8C939A',          # 页脚
}
# gear9 造型的专属默认（金属感渐变），仍可被 colors 覆盖
GEAR9_PALETTE = {
    'wheel_fill': 'D8DCE1',
    'wheel_fill2': '9BA3AB',
    'wheel_line': '',
    'hub_fill': 'F2F4F6',
}
# double 造型的专属默认：照片上做白色字与白线最稳，仍可被 colors 覆盖
DOUBLE_PALETTE = {
    'label': 'FFFFFF',
    'label_num': 'FFFFFF',
    'spoke': 'FFFFFF',
    'ring_line': 'FFFFFF',
    'inner_line': 'FFFFFF',
    'hub_fill': '',
    'hub_line': '',
}

DEFAULTS = {
    'font': '微软雅黑',
    'morph_ms': 1600,
    'auto_advance_ms': None,
    'direction': 'ccw',
    'wheel': 'disc',
    'wheel_size': 1.12,       # 转盘直径 = wheel_size × 页高
    'wheel_cx': 0.06,         # 圆心横坐标占页宽比例（0.06 → 只露右半边）
    'wheel_cy': 0.50,         # 圆心纵坐标占页高比例
    'wheel_shadow': None,     # None → disc 不加、gear9/double 加
    'wheel_line_pt': 1.25,
    'spoke_pt': 1.4,          # 分格线粗细（磅）
    'spoke_len': 0.94,        # 分格线外端占半径比例
    'hub_ratio': 0.20,        # 轴孔半径占转盘半径比例；0 或 '' → 不画轴孔
    'label_prefix': 'part',   # 标签前缀文字；'' → 只显示编号
    'label_radius': 0.70,     # 标签中心到轮心距离占半径比例
    'label_size': 40,         # 前缀字号
    'label_num_size': 20,     # 编号字号
    'label_w_in': 1.5,        # 标签文本框宽（英寸，横排模式）
    'label_h_in': 1.1,        # 标签文本框高（英寸，横排模式）
    'label_mode': None,       # None → 自动：double 用 tangential，其余 flat
    'label_tan_w_in': 3.2,    # 切向标签文本框长边（英寸）
    'label_tan_h_in': 0.7,    # 切向标签文本框短边（英寸）
    'inner_ratio': 0.44,      # double：内盘半径占外盘半径比例
    'inner_edge': None,       # double 且未写 wheel_cx 时：内盘左缘落在页宽的这一比例处（负 = 越出左边界）
    'inner_spin': 1.0,        # double：内盘转角 = 外环转角 × inner_spin（反向）
    'ring_alpha': 55,         # double：外环照片不透明度（%）
    'inner_alpha': 100,       # double：内盘照片不透明度（%）
    'photo': None,            # 照片路径（章节没写 photo 时的统一默认）
    'photo_match': True,      # double：内盘按中央裁切，与外环照片等比衔接
    'ring_line_pt': 1.0,      # double：外环描边粗细（磅）
    'inner_line_pt': 1.0,     # double：内盘描边粗细（磅）
    'inner_shadow': None,     # None → double 给内盘加柔和投影
    'icon_size_in': 0.85,     # 当前章图标边长（英寸）
    'icon_hidden_in': 0.10,   # 非当前章图标缩到多小（藏在轴孔下）
    'icon_offset': 0.40,      # 当前章图标在轮心右侧多远（占半径比例）
    'content_left': 0.535,    # 右侧内容区起点（占页宽比例）
    'content_width': 0.425,
    'content_title_size': 28,
    'content_point_size': 14,
    'background': None,       # 背景图片路径；None = 不设背景，用 PowerPoint 默认白底
    'footer_left': '',
    'footer_name': '',
    'footer_left_x': None,    # None → 自动避开转盘
    'colors': {},
    'cover': None,
    'pages': [],
}

# double 造型的几何默认（照参考视频量出来的比例），配置里显式写了就以配置为准
DOUBLE_DEFAULTS = {
    'wheel_size': 1.17,
    'wheel_cx': None,         # 留空 → 由 inner_edge 反推轮心，改照片尺寸/内盘比例都自动成立
    'inner_edge': -0.012,     # 内盘左缘越出页面左边界一点点（刚好被裁掉一线），与参考图一致
    'content_left': 0.49,
    'content_width': 0.46,
    'hub_ratio': 0.0,
    'label_prefix': 'PART',
    'label_size': 28,
    'label_mode': 'tangential',
    'morph_ms': 1200,
    'spoke_pt': 1.2,
    'wheel_line_pt': 1.0,
    'icon_offset': 0.62,      # 双盘的图标要落在环带上，别被内盘盖住
}

ALLOWED_TOP = set(DEFAULTS)
ALLOWED_PAGE = {'part', 'num', 'section', 'title', 'points', 'icon', 'photo'}
ALLOWED_COVER = {'eyebrow', 'title', 'claim', 'target', 'name_line', 'teacher'}

_FONT = DEFAULTS['font']       # build() 按配置覆盖，供 add_text 取用


# ---------------------------------------------------------------- XML  helpers
# 以下函数是本技能对外暴露的底层配方，签名保持稳定（外部脚本会直接 import）。

def set_slide_bg_picture(slide, img_path):
    """把图片设为整页背景（写在 p:cSld 的第一个子元素 p:bg 里）。"""
    image_part, rId = slide.part.get_or_add_image_part(img_path)
    cSld = slide._element.find(qn('p:cSld'))
    old = cSld.find(qn('p:bg'))
    if old is not None:
        cSld.remove(old)
    xml = ('<p:bg %s><p:bgPr><a:blipFill dpi="0" rotWithShape="1">'
           '<a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch>'
           '</a:blipFill><a:effectLst/></p:bgPr></p:bg>') % (nsdecls('p', 'a', 'r'), rId)
    cSld.insert(0, parse_xml(xml))


def set_picture_fill(shape, img_path, alpha=None, crop_frac=None):
    """把照片填进形状（spPr 里的 a:blipFill）。

    rotWithShape="1" 是关键：形状被 Morph 旋转时，照片跟着一起转，
    转盘才真的像「照片在盘里转」而不是「盘在照片上转」。
    alpha      1~100 的不透明度（写 None 表示完全不透明），外环用它压淡。
    crop_frac  0~1，只取照片中央这一比例再铺满形状。内盘传 inner_ratio，
               照片尺度就和外环完全一致，两个盘看起来是同一张图的延续。
    """
    if not img_path or not os.path.isfile(img_path):
        raise ValueError('照片文件不存在：%r' % img_path)
    image_part, rId = shape.part.get_or_add_image_part(img_path)
    spPr = shape._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
                'a:pattFill', 'a:grpFill'):
        el = spPr.find(qn(tag))
        if el is not None:
            spPr.remove(el)
    alpha_xml = ('<a:alphaModFix amt="%d"/>' % int(alpha * 1000)
                 if alpha is not None and alpha < 100 else '')
    src_xml = ''
    if crop_frac is not None and 0 < float(crop_frac) < 1:
        inset = int(round((1.0 - float(crop_frac)) / 2.0 * 100000))
        src_xml = '<a:srcRect l="%d" t="%d" r="%d" b="%d"/>' % (
            inset, inset, inset, inset)
    xml = ('<a:blipFill %s rotWithShape="1"><a:blip r:embed="%s">%s</a:blip>'
           '%s<a:stretch><a:fillRect/></a:stretch></a:blipFill>') % (
        nsdecls('a', 'r'), rId, alpha_xml, src_xml)
    geom = spPr.find(qn('a:prstGeom'))
    if geom is not None:
        geom.addnext(parse_xml(xml))
    else:
        spPr.append(parse_xml(xml))
    return rId


def add_outer_shadow(shape, blur_pt=8.0, dist_pt=14.0, dir_deg=30.0, alpha_pct=55):
    """给形状加外阴影。"""
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


def set_gradient_fill(shape, color='767D84', ang_deg=45, alpha_start=100, alpha_end=100,
                      color2=None):
    """线性渐变填充。color2 为空时两端同色（等价于带 alpha 的纯色）。"""
    spPr = shape._element.spPr
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'):
        el = spPr.find(qn(tag))
        if el is not None:
            spPr.remove(el)
    xml = ('<a:gradFill %s rotWithShape="1"><a:gsLst>'
           '<a:gs pos="0"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:gs>'
           '<a:gs pos="100000"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:gs>'
           '</a:gsLst><a:lin ang="%d" scaled="1"/></a:gradFill>') % (
        nsdecls('a'), color, int(alpha_start * 1000), color2 or color,
        int(alpha_end * 1000), int(ang_deg * 60000) % 21600000)
    spPr.find(qn('a:prstGeom')).addnext(parse_xml(xml))


def set_morph_transition(slide, dur_ms=1600, auto_advance_ms=None):
    """Morph（平滑）切换：mc:AlternateContent，Choice 走 p159:morph byObject，
    Fallback 退化成普通切换（旧版 PowerPoint / WPS 不会报错，只是不平滑）。"""
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


# ---------------------------------------------------------------- small utils

def rgb(hexstr):
    """'1F2429' / '#1F2429' / 'abc' → RGBColor；空值返回 None。"""
    if hexstr is None:
        return None
    h = str(hexstr).strip().lstrip('#')
    if not h:
        return None
    if len(h) == 3:
        h = ''.join(c * 2 for c in h)
    if len(h) != 6 or any(c not in '0123456789abcdefABCDEF' for c in h):
        raise ValueError('颜色要写成 6 位十六进制（如 1F2429），收到 %r' % hexstr)
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def resolve_path(path, base_dir=None):
    """配置里的相对路径以配置文件所在目录为基准。"""
    if not path:
        return None
    p = str(path)
    if base_dir and not os.path.isabs(p):
        p = os.path.join(base_dir, p)
    return os.path.normpath(p)


def label_number(page, k):
    """标签第二行的编号：优先 page.num，其次从 part 末尾取，最后退回序号。"""
    if page.get('num'):
        return str(page['num'])
    text = str(page.get('part') or '').strip()
    if not text:
        return '%02d' % (k + 1)
    toks = text.split()
    return toks[-1] if len(toks) > 1 else text


# ---------------------------------------------------------------- shape helpers

def add_text(shapes, x, y, w, h, text, size, color=None, bold=False,
             align=PP_ALIGN.LEFT, name=None, spacing=None, font=None):
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
    run.font.color.rgb = color if color is not None else rgb(PALETTE['title'])
    run.font.name = font or _FONT
    if spacing is not None:
        run.font._rPr.set('spc', str(int(spacing * 100)))
    return box


def add_bar(shapes, x, y, w, h, color, name=None):
    bar = shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x)), Emu(int(y)),
                           Emu(int(w)), Emu(int(h)))
    if name:
        bar.name = name
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()
    bar.shadow.inherit = False
    return bar


def add_footer(shapes, idx, total, footer_left, footer_name, left_x=None, pal=None):
    pal = pal or PALETTE
    y = SH * 0.945
    if left_x is None:
        left_x = int(SW * 0.035)
    dim = rgb(pal['dim'])
    if footer_left:
        add_text(shapes, left_x, y, SW * 0.60, SH * 0.04, footer_left, 9, color=dim)
    right = '%02d / %02d' % (idx, total)
    if footer_name:
        right = '%s · %s' % (footer_name, right)
    add_text(shapes, SW * 0.62, y, SW * 0.345, SH * 0.04, right, 9, color=dim,
             align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------- wheel

def label_mode(cfg):
    """标签排法：flat 永远横排；tangential 沿圆周切向、跟着盘转。"""
    mode = cfg.get('label_mode')
    if mode:
        return mode
    return 'tangential' if cfg['wheel'] == 'double' else 'flat'


def spoke_band(shapes, cx, cy, r0, r1, ang, w, color, name):
    """一段分格线：从半径 r0 延伸到 r1 的细矩形。

    必须用「矩形 + rotation」而不是画直线，Morph 才能补间出旋转。
    r0=0 就是一端在轮心的传统放射线。
    """
    length = r1 - r0
    if length <= 0:
        return None
    rad = math.radians(ang)
    mid = (r0 + r1) / 2.0
    sx, sy = cx + mid * math.cos(rad), cy + mid * math.sin(rad)
    sp = shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(sx - length / 2)),
                          Emu(int(sy - w / 2)), Emu(int(length)), Emu(int(w)))
    sp.name = name
    sp.rotation = ang % 360
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def add_labels(shapes, cx, cy, r, offset, step, nums, cfg, pal):
    """章节标签。第 k 个的极角 = offset - k*step，所以第 k 页 k 个恒在 3 点钟。

    flat        文本框不旋转（rotation=0），前缀一行 + 编号一行。
    tangential  文本框 rotation = 极角 + 90，一行「前缀 编号」，字贴着圆周走，
                转到右缘竖着、转到底下会倒过来（参考视频里的排法）。
    """
    mode = label_mode(cfg)
    rr = int(r * cfg['label_radius'])
    prefix = str(cfg['label_prefix'] or '')
    for k, num in enumerate(nums):
        ang = offset - k * step
        rad = math.radians(ang)
        px, py = cx + rr * math.cos(rad), cy + rr * math.sin(rad)
        if mode == 'tangential':
            w = int(cfg['label_tan_w_in'] * EMU_IN)
            h = int(cfg['label_tan_h_in'] * EMU_IN)
        else:
            w = int(cfg['label_w_in'] * EMU_IN)
            h = int(cfg['label_h_in'] * EMU_IN)
        box = shapes.add_textbox(Emu(int(px - w / 2)), Emu(int(py - h / 2)),
                                 Emu(w), Emu(h))
        box.name = 'roulette-label-%d' % k
        box.rotation = (ang + 90) % 360 if mode == 'tangential' else 0
        box.shadow.inherit = False
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        def style(run, size, color, bold=False, spacing=None):
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = color
            run.font.name = _FONT
            if spacing is not None:
                run.font._rPr.set('spc', str(int(spacing * 100)))

        if mode == 'tangential':
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = ('%s %s' % (prefix, num)).strip()
            style(run, cfg['label_size'], rgb(pal['label']), bold=True, spacing=2)
            continue
        if prefix:
            p1 = tf.paragraphs[0]
            p1.alignment = PP_ALIGN.CENTER
            run1 = p1.add_run()
            run1.text = prefix
            style(run1, cfg['label_size'], rgb(pal['label']))
            p2 = tf.add_paragraph()
            p2.space_before = Pt(6)
            num_size, num_color = cfg['label_num_size'], rgb(pal['label_num'])
        else:
            p2 = tf.paragraphs[0]
            num_size, num_color = cfg['label_size'], rgb(pal['label'])
        p2.alignment = PP_ALIGN.CENTER
        run2 = p2.add_run()
        run2.text = num
        style(run2, num_size, num_color)


def add_double_wheel(shapes, cx, cy, r, offset, step, nums, current, cfg, pal,
                     photos, icons=None):
    """内外双盘反转（参考视频实测比例）：

    外环  roulette-gear   照片铺满 + ring_alpha 压淡，rotation = offset（逆时针）
    内盘  roulette-inner  同一张照片按中央裁切（尺度与外环一致），
                          rotation = -offset × inner_spin（顺时针，反向）
    分格线只在环带上（内盘外缘 → 外盘内缘），标签沿圆周切向，画在最上层。
    内盘不透明，所以缩回轮心的图标直接被它盖住，双盘模式不需要轴孔。
    """
    inner_r = int(r * max(0.05, min(0.95, float(cfg['inner_ratio']))))

    ring = shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx - r)), Emu(int(cy - r)),
                            Emu(int(2 * r)), Emu(int(2 * r)))
    ring.name = 'roulette-gear'          # 名字沿用，Morph 与外部脚本都按此匹配
    ring.rotation = offset % 360
    ring.shadow.inherit = False
    set_picture_fill(ring, photos[current], alpha=cfg['ring_alpha'])
    ring_line = rgb(pal.get('ring_line'))
    if ring_line is not None:
        ring.line.color.rgb = ring_line
        ring.line.width = Pt(cfg['ring_line_pt'])
    else:
        ring.line.fill.background()
    if cfg['wheel_shadow'] if cfg['wheel_shadow'] is not None else True:
        add_outer_shadow(ring, blur_pt=10, dist_pt=16, dir_deg=90, alpha_pct=30)

    spoke_color = rgb(pal.get('spoke'))
    if spoke_color is not None:
        w = int(cfg['spoke_pt'] * PT)
        for k in range(len(nums)):
            spoke_band(shapes, cx, cy, int(inner_r * 1.03),
                       int(r * cfg['spoke_len']), offset - (k + 0.5) * step, w,
                       spoke_color, 'roulette-spoke-%d' % k)

    if icons:
        big_size = int(cfg['icon_size_in'] * EMU_IN)
        small_size = int(cfg['icon_hidden_in'] * EMU_IN)
        big_cx = cx + int(r * cfg['icon_offset'])
        for k in range(len(nums)):
            if not os.path.isfile(icons[k]):
                raise ValueError('图标文件不存在：%s' % icons[k])
            icx, icy, isize = ((big_cx, cy, big_size) if k == current
                               else (cx, cy, small_size))
            pic = shapes.add_picture(icons[k], Emu(int(icx - isize / 2)),
                                     Emu(int(icy - isize / 2)), Emu(isize), Emu(isize))
            pic.name = 'chapter-icon-%d' % k

    inner = shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx - inner_r)),
                             Emu(int(cy - inner_r)), Emu(int(2 * inner_r)),
                             Emu(int(2 * inner_r)))
    inner.name = 'roulette-inner'
    inner.rotation = (-offset * float(cfg['inner_spin'])) % 360
    inner.shadow.inherit = False
    crop = (inner_r / float(r)) if cfg['photo_match'] else None
    set_picture_fill(inner, photos[current], alpha=cfg['inner_alpha'],
                     crop_frac=crop)
    inner_line = rgb(pal.get('inner_line'))
    if inner_line is not None:
        inner.line.color.rgb = inner_line
        inner.line.width = Pt(cfg['inner_line_pt'])
    else:
        inner.line.fill.background()
    if cfg['inner_shadow'] if cfg['inner_shadow'] is not None else True:
        add_outer_shadow(inner, blur_pt=12, dist_pt=18, dir_deg=90, alpha_pct=40)

    hub_ratio = float(cfg['hub_ratio'] or 0)
    if hub_ratio > 0:
        hr = int(r * hub_ratio)
        hub = shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx - hr)), Emu(int(cy - hr)),
                               Emu(2 * hr), Emu(2 * hr))
        hub.name = 'roulette-hub'
        hub_line = rgb(pal.get('hub_line'))
        hub_fill = rgb(pal.get('hub_fill'))
        if hub_fill is None:
            hub.fill.solid()
            hub.fill.fore_color.rgb = rgb('FFFFFF')
        else:
            hub.fill.solid()
            hub.fill.fore_color.rgb = hub_fill
        if hub_line is not None:
            hub.line.color.rgb = hub_line
            hub.line.width = Pt(1.0)
        else:
            hub.line.fill.background()
        hub.shadow.inherit = False

    add_labels(shapes, cx, cy, r, offset, step, nums, cfg, pal)


def add_wheel(shapes, cx, cy, r, offset, step, nums, current, icons, cfg, pal,
              photos=None):
    """转盘本体。

    cx/cy/r     轮心与半径（EMU）
    offset      本页外盘整盘旋转角（度，PowerPoint 顺时针为正）
    step        相邻扇区夹角（度，带符号：ccw 为负、cw 为正）
    nums        每章标签编号文本
    current     当前章索引
    icons       每章图标文件路径列表，或 None（关闭图标接力）
    photos      每章照片文件路径列表，或 None（不启用照片填充）

    关键：所有形状跨页同名（roulette-gear / roulette-inner / roulette-spoke-k /
    chapter-icon-k / roulette-label-k / roulette-hub），Morph 才会按对象补间出真旋转。
    标签第 k 个的极角 = offset - k*step，所以第 i 页 k=i 恒为 0°（3 点钟）。
    double 造型另走 add_double_wheel：内外盘同名跨页，内盘反向旋转。
    """
    if cfg['wheel'] == 'double':
        if not photos:
            raise ValueError('wheel=double 需要照片：在每章写 "photo": "路径"，'
                             '或在顶层写一个 photo')
        add_double_wheel(shapes, cx, cy, r, offset, step, nums, current, cfg, pal,
                         photos, icons)
        return

    gear9 = cfg['wheel'] == 'gear9'
    body = shapes.add_shape(MSO_SHAPE.GEAR_9 if gear9 else MSO_SHAPE.OVAL,
                            Emu(int(cx - r)), Emu(int(cy - r)),
                            Emu(int(2 * r)), Emu(int(2 * r)))
    body.name = 'roulette-gear'          # 名字沿用历史值，外部脚本按此名匹配
    body.rotation = offset % 360
    body.shadow.inherit = False
    if photos:
        set_picture_fill(body, photos[current], alpha=cfg['inner_alpha'])
        line = rgb(pal.get('wheel_line'))
        if line is not None:
            body.line.color.rgb = line
            body.line.width = Pt(cfg['wheel_line_pt'])
        else:
            body.line.fill.background()
    elif pal.get('wheel_fill2'):
        body.line.fill.background()
        set_gradient_fill(body, color=pal['wheel_fill'], ang_deg=45,
                          color2=pal['wheel_fill2'])
    else:
        body.fill.solid()
        body.fill.fore_color.rgb = rgb(pal['wheel_fill'])
        line = rgb(pal.get('wheel_line'))
        if line is not None:
            body.line.color.rgb = line
            body.line.width = Pt(cfg['wheel_line_pt'])
        else:
            body.line.fill.background()
    shadow = cfg['wheel_shadow'] if cfg['wheel_shadow'] is not None else gear9
    if shadow:
        add_outer_shadow(body, blur_pt=8, dist_pt=14, dir_deg=30, alpha_pct=35)

    # 分格线：每条一个细矩形，一端在轮心、一端指向扇区分界（与标签错开半步）。
    spoke_color = rgb(pal.get('spoke'))
    if spoke_color is not None:
        w = int(cfg['spoke_pt'] * PT)
        length = int(r * cfg['spoke_len'])
        for k in range(len(nums)):
            spoke_band(shapes, cx, cy, 0, length, offset - (k + 0.5) * step, w,
                       spoke_color, 'roulette-spoke-%d' % k)

    # 图标接力：当前章放大在轮心右侧（标签左边），其余缩到 0.1in 藏在轮心，
    # 由后画的轴孔盖住。图标必须加在轴孔之前。
    if icons:
        big_size = int(cfg['icon_size_in'] * EMU_IN)
        small_size = int(cfg['icon_hidden_in'] * EMU_IN)
        big_cx = cx + int(r * cfg['icon_offset'])
        for k in range(len(nums)):
            if k == current:
                icx, icy, isize = big_cx, cy, big_size
            else:
                icx, icy, isize = cx, cy, small_size
            if not os.path.isfile(icons[k]):
                raise ValueError('图标文件不存在：%s' % icons[k])
            pic = shapes.add_picture(icons[k], Emu(int(icx - isize / 2)),
                                     Emu(int(icy - isize / 2)), Emu(isize), Emu(isize))
            pic.name = 'chapter-icon-%d' % k

    # 中心轴孔（不透明，负责盖住缩小的非当前图标）
    hub_ratio = float(cfg['hub_ratio'] or 0)
    if hub_ratio > 0:
        hr = int(r * hub_ratio)
        hub = shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(cx - hr)), Emu(int(cy - hr)),
                               Emu(2 * hr), Emu(2 * hr))
        hub.name = 'roulette-hub'
        hub.fill.solid()
        hub.fill.fore_color.rgb = rgb(pal['hub_fill'])
        hub_line = rgb(pal.get('hub_line'))
        if hub_line is not None:
            hub.line.color.rgb = hub_line
            hub.line.width = Pt(1.0)
        else:
            hub.line.fill.background()
        hub.shadow.inherit = False

    add_labels(shapes, cx, cy, r, offset, step, nums, cfg, pal)


def add_gear_roulette(shapes, cx, cy, r, offset, labels, current, icons=None):
    """旧签名的兼容封装（4 章时代的外部脚本按此名调用）：齿轮造型、逆时针、
    扇区数按 labels 长度自动推。icons 现在要传文件路径，不再是素材名。"""
    n = max(1, len(labels))
    cfg = dict(DEFAULTS)
    cfg['wheel'] = 'gear9'
    pal = resolve_palette(cfg)
    step = -360.0 / n
    nums = [label_number({'part': t}, k) for k, t in enumerate(labels)]
    add_wheel(shapes, cx, cy, r, offset, step, nums, current, icons, cfg, pal)


def add_content(shapes, part, section, title, points, cfg, pal):
    left = int(SW * cfg['content_left'])
    width = int(SW * cfg['content_width'])
    accent = rgb(pal['accent'])
    head = ' · '.join(x for x in (part, section) if x)
    add_text(shapes, left, SH * 0.215, width, SH * 0.05, head, 14,
             color=accent, bold=True, spacing=2)
    add_text(shapes, left, SH * 0.285, width, SH * 0.13, title,
             cfg['content_title_size'], color=rgb(pal['title']), bold=True)
    add_bar(shapes, left, SH * 0.435, int(SW * 0.055), Pt(2.2), accent)
    y = SH * 0.49
    for point in points:
        add_bar(shapes, left, y + Pt(6), Pt(5), Pt(5), accent)
        add_text(shapes, left + Pt(14), y - Pt(2), width - Pt(14), SH * 0.10,
                 point, cfg['content_point_size'], color=rgb(pal['text']))
        y += SH * 0.105


def footer_left_x(cfg, cx, cy, r):
    """页脚左边界：自动避开转盘在页脚那一行探出来的右缘。"""
    if cfg.get('footer_left_x') is not None:
        return int(SW * float(cfg['footer_left_x']))
    y = SH * 0.945 + SH * 0.02
    dy = abs(y - cy)
    if dy >= r:
        return int(SW * 0.035)
    edge = cx + math.sqrt(r * r - dy * dy)
    return int(max(SW * 0.035, min(edge + SW * 0.02, SW * 0.45)))


# ---------------------------------------------------------------- config

def resolve_config(cfg):
    if not isinstance(cfg, dict):
        raise ValueError('配置必须是 JSON 对象')
    unknown = sorted(set(cfg) - ALLOWED_TOP)
    if unknown:
        raise ValueError('未知配置项：%s' % ', '.join(unknown))
    out = dict(DEFAULTS)
    if cfg.get('wheel') == 'double':
        for key, value in DOUBLE_DEFAULTS.items():
            if key not in cfg:
                out[key] = value
    out.update(cfg)
    if out['wheel'] not in WHEELS:
        raise ValueError('wheel 只能是 %s，收到 %r' % (' / '.join(WHEELS), out['wheel']))
    if out['direction'] not in DIRECTIONS:
        raise ValueError('direction 只能是 %s，收到 %r'
                         % (' / '.join(DIRECTIONS), out['direction']))
    if out.get('label_mode') not in (None, '') + LABEL_MODES:
        raise ValueError('label_mode 只能是 %s（或留空自动），收到 %r'
                         % (' / '.join(LABEL_MODES), out['label_mode']))
    try:
        out['inner_ratio'] = float(out['inner_ratio'])
        out['inner_spin'] = float(out['inner_spin'])
        out['ring_alpha'] = max(0, min(100, float(out['ring_alpha'])))
        out['inner_alpha'] = max(0, min(100, float(out['inner_alpha'])))
    except (TypeError, ValueError):
        raise ValueError('inner_ratio / inner_spin / ring_alpha / inner_alpha '
                         '必须是数字')
    if not 0.05 <= out['inner_ratio'] <= 0.95:
        raise ValueError('inner_ratio 需在 0.05~0.95 之间，收到 %r' % out['inner_ratio'])
    if out['wheel'] == 'double':
        if out.get('wheel_cx') is None and out.get('inner_edge') is None:
            raise ValueError('double：wheel_cx 与 inner_edge 不能同时留空')
        if out.get('inner_edge') is not None:
            try:
                out['inner_edge'] = float(out['inner_edge'])
            except (TypeError, ValueError):
                raise ValueError('inner_edge 必须是数字（占页宽比例，负值 = 越出页面左边界）')
            if not -0.5 <= out['inner_edge'] <= 0.5:
                raise ValueError('inner_edge 需在 -0.5~0.5 之间，收到 %r' % out['inner_edge'])
    colors = out.get('colors')
    if colors is None:
        colors = {}
    if not isinstance(colors, dict):
        raise ValueError('colors 必须是对象')
    bad = sorted(set(colors) - set(PALETTE))
    if bad:
        raise ValueError('未知配色项：%s（可用：%s）'
                         % (', '.join(bad), ', '.join(sorted(PALETTE))))
    out['colors'] = colors
    pages = out.get('pages') or []
    if not isinstance(pages, list):
        raise ValueError('pages 必须是数组')
    for i, pg in enumerate(pages):
        if not isinstance(pg, dict):
            raise ValueError('pages[%d] 必须是对象' % i)
        bad = sorted(set(pg) - ALLOWED_PAGE)
        if bad:
            raise ValueError('pages[%d] 未知字段：%s' % (i, ', '.join(bad)))
        if not isinstance(pg.get('points') or [], list):
            raise ValueError('pages[%d].points 必须是数组' % i)
    cover = out.get('cover')
    if cover is not None:
        if not isinstance(cover, dict):
            raise ValueError('cover 必须是对象或 null')
        bad = sorted(set(cover) - ALLOWED_COVER)
        if bad:
            raise ValueError('cover 未知字段：%s' % ', '.join(bad))
    return out


def resolve_palette(cfg):
    pal = dict(PALETTE)
    if cfg.get('wheel') == 'gear9':
        pal.update(GEAR9_PALETTE)
    if cfg.get('wheel') == 'double':
        pal.update(DOUBLE_PALETTE)
    pal.update(cfg.get('colors') or {})
    for value in pal.values():
        rgb(value)                      # 提前校验，别让错色值跑到一半才炸
    return pal


# ---------------------------------------------------------------- build

DEMO = {
    'footer_left': '演示汇报 · 转盘',
    'footer_name': '汇报人',
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
         'points': ['要点一：一句话', '要点二：一句话', '要点三：一句话']},
        {'part': 'PART 02', 'section': '第二章', 'title': '第二章核心观点',
         'points': ['要点一：一句话', '要点二：一句话', '要点三：一句话']},
        {'part': 'PART 03', 'section': '第三章', 'title': '第三章核心观点',
         'points': ['要点一：一句话', '要点二：一句话', '要点三：一句话']},
        {'part': 'PART 04', 'section': '第四章', 'title': '第四章核心观点',
         'points': ['要点一：一句话', '要点二：一句话', '要点三：一句话']},
    ],
}


def build(out_path, cfg, base_dir=None):
    cfg = resolve_config(cfg)
    pages = cfg['pages']
    n = len(pages)
    if not MIN_CHAPTERS <= n <= MAX_CHAPTERS:
        raise ValueError('章数需在 %d~%d 之间，当前 %d' % (MIN_CHAPTERS, MAX_CHAPTERS, n))

    global _FONT
    _FONT = cfg['font']

    pal = resolve_palette(cfg)
    step_abs = 360.0 / n
    step = -step_abs if cfg['direction'] == 'ccw' else step_abs

    icons = [pg.get('icon') for pg in pages]
    icons = [resolve_path(p, base_dir) for p in icons] if all(icons) else None
    photo_src = [pg.get('photo') or cfg.get('photo') for pg in pages]
    photos = None
    if any(photo_src):
        photos = [resolve_path(p, base_dir) for p in photo_src]
        bad = [i for i, p in enumerate(photos) if not p or not os.path.isfile(p)]
        if bad:
            raise ValueError('photo 未提供或文件不存在：pages[%s]'
                             % '], pages['.join(str(i) for i in bad))
    bg = resolve_path(cfg['background'], base_dir)
    if bg and not os.path.isfile(bg):
        raise ValueError('背景图不存在：%s' % bg)
    nums = [label_number(pg, k) for k, pg in enumerate(pages)]

    total = n + (1 if cfg.get('cover') else 0)
    prs = Presentation()
    prs.slide_width = Emu(SW)
    prs.slide_height = Emu(SH)
    blank = prs.slide_layouts[6]

    cover = cfg.get('cover')
    if cover:
        s = prs.slides.add_slide(blank)
        if bg:
            set_slide_bg_picture(s, bg)
        sh = s.shapes
        add_text(sh, SW * 0.06, SH * 0.20, SW * 0.60, SH * 0.05,
                 cover.get('eyebrow', ''), 13, color=rgb(pal['accent']), spacing=2)
        add_text(sh, SW * 0.06, SH * 0.29, SW * 0.60, SH * 0.15,
                 cover.get('title', ''), 52, color=rgb(pal['title']), bold=True)
        add_bar(sh, SW * 0.062, SH * 0.47, int(SW * 0.07), Pt(2.5), rgb(pal['accent']))
        add_text(sh, SW * 0.06, SH * 0.51, SW * 0.58, SH * 0.08,
                 cover.get('claim', ''), 17, color=rgb(pal['text']))
        add_text(sh, SW * 0.06, SH * 0.63, SW * 0.58, SH * 0.05,
                 cover.get('target', ''), 15, color=rgb(pal['title']), bold=True)
        add_text(sh, SW * 0.06, SH * 0.71, SW * 0.62, SH * 0.05,
                 cover.get('name_line', ''), 12, color=rgb(pal['dim']))
        if cover.get('teacher'):
            add_text(sh, SW * 0.06, SH * 0.77, SW * 0.62, SH * 0.05,
                     cover['teacher'], 12, color=rgb(pal['dim']))
        add_footer(sh, 1, total, cfg['footer_left'], cfg['footer_name'], pal=pal)
        set_morph_transition(s, dur_ms=cfg['morph_ms'],
                             auto_advance_ms=cfg['auto_advance_ms'])

    r = int(SH * cfg['wheel_size'] / 2.0)
    cy = int(SH * cfg['wheel_cy'])
    if cfg['wheel'] == 'double' and cfg['wheel_cx'] is None:
        # 轮心由「内盘左缘落在哪」反推：inner_edge 为负 → 内盘刚好被页面左边界裁掉一线
        cx = int(r * max(0.05, min(0.95, float(cfg['inner_ratio'])))
                 + SW * float(cfg['inner_edge']))
    else:
        cx = int(SW * float(cfg['wheel_cx']))
    foot_x = footer_left_x(cfg, cx, cy, r)
    for i, pg in enumerate(pages):
        s = prs.slides.add_slide(blank)
        if bg:
            set_slide_bg_picture(s, bg)
        sh = s.shapes
        add_wheel(sh, cx, cy, r, i * step, step, nums, i, icons, cfg, pal, photos)
        add_content(sh, pg.get('part') or 'PART %02d' % (i + 1), pg.get('section', ''),
                    pg.get('title', ''), pg.get('points') or [], cfg, pal)
        add_footer(sh, i + 1 + (1 if cover else 0), total,
                   cfg['footer_left'], cfg['footer_name'], left_x=foot_x, pal=pal)
        set_morph_transition(s, dur_ms=cfg['morph_ms'],
                             auto_advance_ms=cfg['auto_advance_ms'])

    prs.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser(description='翻页旋转转盘 PPT 生成器（风格无关）')
    ap.add_argument('-o', '--out', default='zhuanpan.pptx')
    ap.add_argument('--config', default=None,
                    help='JSON 配置；里面的相对路径以配置文件所在目录为基准')
    ap.add_argument('--wheel', default=None, choices=list(WHEELS),
                    help='覆盖配置里的转盘造型')
    ap.add_argument('--direction', default=None, choices=list(DIRECTIONS),
                    help='覆盖配置里的旋转方向')
    args = ap.parse_args()

    cfg = dict(DEMO)
    base_dir = None
    if args.config:
        with open(args.config, encoding='utf-8') as f:
            cfg = json.load(f)
        base_dir = os.path.dirname(os.path.abspath(args.config))
    if args.wheel:
        cfg['wheel'] = args.wheel
    if args.direction:
        cfg['direction'] = args.direction

    try:
        out = build(args.out, cfg, base_dir=base_dir)
    except Exception as exc:                            # noqa: BLE001 - CLI 统一报错
        print('ERROR: %s' % exc)
        return 1

    n = len(cfg.get('pages') or [])
    wheel = cfg.get('wheel') or DEFAULTS['wheel']
    mode = cfg.get('label_mode') or ('tangential' if wheel == 'double' else 'flat')
    print('saved: %s | %d bytes | pages: %d | chapters: %d (step %.1f) | wheel: %s'
          ' | direction: %s | label: %s%s'
          % (os.path.abspath(out), os.path.getsize(out),
             n + (1 if cfg.get('cover') else 0), n, 360.0 / n, wheel,
             cfg.get('direction') or DEFAULTS['direction'], mode,
             ' | photo fill' if any(p.get('photo') for p in (cfg.get('pages') or []))
             or cfg.get('photo') else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main())
