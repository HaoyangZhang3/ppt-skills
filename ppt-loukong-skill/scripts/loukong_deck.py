# -*- coding: utf-8 -*-
"""ppt-loukong-skill：透明字孔开场 / 谢幕 PPTX 生成器（自包含）。

  python loukong_deck.py -o opening.pptx --mode opening
  python loukong_deck.py -o closing.pptx --config closing.json

依赖：python-pptx、Pillow。字孔以透明 PNG 生成，标题修改后需重新运行。
"""
import argparse
import json
import os
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.normpath(os.path.join(BASE, '..', 'assets'))
SW, SH = 12192000, 6858000  # 16:9，13.333 × 7.5 英寸
PX_W, PX_H = 1920, 1080
MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
P14_NS = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
P159_NS = 'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"'

DEFAULTS = {
    'mode': 'opening',
    'title': None,
    'image': None,
    'font_path': None,
    'font_size_pt': None,
    'text_x': 0.50,
    'text_y': 0.50,
    'mask_color': '#000000',
    'mask_opacity': 100,
    'canvas_color': '#FFFFFF',
    'initial_offset': 0.06,
    'image_scale': 1.10,
    'morph_ms': 2000,
    'closing_message': '愿你我更上一层楼',
    'closing_subtitle': '感谢各位的观看和支持，您的关注是我持续更新的动力',
    'closing_text_font': '微软雅黑',
    'closing_message_size': 34,
    'closing_subtitle_size': 14,
    'scrim_opacity': 28,
}


def rgb(hex_color):
    value = str(hex_color).strip().lstrip('#')
    if len(value) != 6:
        raise ValueError('颜色必须是 #RRGGBB：%s' % hex_color)
    return RGBColor.from_string(value.upper())


def rgba_tuple(hex_color):
    c = rgb(hex_color)
    return c[0], c[1], c[2]


def resolve_path(value, config_dir, default=None):
    path = value or default
    if not path:
        return None
    path = os.path.expanduser(os.path.expandvars(str(path)))
    if not os.path.isabs(path):
        path = os.path.join(config_dir, path)
    return os.path.abspath(path)


def find_font(configured, config_dir):
    candidates = [
        resolve_path(configured, config_dir),
        r'C:\Windows\Fonts\impact.ttf',
        r'C:\Windows\Fonts\msyhbd.ttc',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
        '/System/Library/Fonts/Supplemental/Arial Bold.ttf',
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    raise OSError('找不到粗体字体；请在配置中指定 font_path。')


def set_morph_transition(slide, dur_ms):
    """给“进入本页”的切换写入 Morph；兼容分支让旧版正常打开。"""
    sld = slide._element
    mc_tag = '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
    for child in list(sld):
        if child.tag == mc_tag or child.tag == qn('p:transition'):
            sld.remove(child)
    xml = ('<mc:AlternateContent %s %s %s><mc:Choice %s Requires="p159">'
           '<p:transition spd="slow" p14:dur="%d"><p159:morph option="byObject"/>'
           '</p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"/>'
           '</mc:Fallback></mc:AlternateContent>') % (
               MC_NS, nsdecls('p'), P14_NS, P159_NS, int(dur_ms))
    sld.append(parse_xml(xml))


def set_fill_opacity(shape, opacity_pct):
    solid = shape._element.spPr.find(qn('a:solidFill'))
    color = solid.find(qn('a:srgbClr'))
    old = color.find(qn('a:alpha'))
    if old is not None:
        color.remove(old)
    color.append(parse_xml('<a:alpha %s val="%d"/>' % (
        nsdecls('a'), int(float(opacity_pct) * 1000))))


def add_cover_picture(slide, path, scale):
    """等比居中裁切底图；跨页只改位置和大小，不改图像内容。"""
    with Image.open(path) as im:
        iw, ih = im.size
    image_ratio = iw / float(ih)
    frame_ratio = SW / float(SH)
    w, h = int(SW * scale), int(SH * scale)
    pic = slide.shapes.add_picture(path, Emu(int((SW - w) / 2)),
                                   Emu(int((SH - h) / 2)), Emu(w), Emu(h))
    if image_ratio > frame_ratio:
        crop = (1.0 - frame_ratio / image_ratio) / 2.0
        pic.crop_left = crop
        pic.crop_right = crop
    elif image_ratio < frame_ratio:
        crop = (1.0 - image_ratio / frame_ratio) / 2.0
        pic.crop_top = crop
        pic.crop_bottom = crop
    pic.name = '!!loukong-bg'
    return pic


def make_mask_pair(title, font_path, font_size_pt, text_x, text_y,
                   mask_color, mask_opacity, initial_offset, out_dir):
    """宽幅遮罩：比页面两侧各宽出一段余量，字孔仍按 text_x/text_y 落在页面坐标上。

    余量足够大，首页才能把字孔整孔平移到页面之外（上条在右缘外、下条在左缘外），
    而遮罩主体始终盖满页面；Morph 只补间位置，不换图。
    4 倍像素面积（2×2）超采样，分别裁出上下独立遮罩。
    """
    ss = 2
    emu_per_px = SW / float(PX_W)
    opacity = int(round(float(mask_opacity) * 255 / 100))
    # 先在页面大小的画布上试排：字号仍只受页面约束，与遮罩加宽无关；
    # 重心 text_x 偏离中线时，字宽上限随之收紧，保证标题完整落在页面内
    max_w_px = min(0.90, 2 * min(float(text_x), 1.0 - float(text_x))) * PX_W * ss
    scratch = Image.new('L', (PX_W * ss, PX_H * ss), opacity)
    draw = ImageDraw.Draw(scratch)
    requested = int(round(float(font_size_pt) * 2 * ss))
    spacing_ratio = 0.12
    font = None
    bbox = None
    for font_size in range(requested, 15, -max(2, requested // 80)):
        trial = ImageFont.truetype(font_path, font_size)
        trial_bbox = draw.multiline_textbbox(
            (0, 0), title, font=trial, spacing=int(font_size * spacing_ratio),
            align='center')
        width = trial_bbox[2] - trial_bbox[0]
        height = trial_bbox[3] - trial_bbox[1]
        if width <= max_w_px and height <= PX_H * ss * 0.74:
            font, bbox = trial, trial_bbox
            break
    if font is None:
        raise ValueError('标题过长，无法放入画面：%s' % title)
    hole_half_emu = int(round((bbox[2] - bbox[0]) / 2.0 / ss * emu_per_px))
    # 单侧余量：字孔整孔出页 + 遮罩仍盖满页面，取上/下两条中更苛刻的一方
    overhang_emu = int(round(
        (float(initial_offset) + max(float(text_x), 1.0 - float(text_x))) * SW
        + hole_half_emu + SW * 0.01))
    mask_w_emu = SW + 2 * overhang_emu
    mask_px_w = int(round(mask_w_emu / emu_per_px))
    alpha = Image.new('L', (mask_px_w * ss, PX_H * ss), opacity)
    draw = ImageDraw.Draw(alpha)
    cx = (overhang_emu + float(text_x) * SW) / emu_per_px * ss
    cy = float(text_y) * PX_H * ss
    pos = (cx - (bbox[0] + bbox[2]) / 2.0,
           cy - (bbox[1] + bbox[3]) / 2.0)
    draw.multiline_text(pos, title, font=font, fill=0,
                        spacing=int(font.size * spacing_ratio), align='center')
    half_h = PX_H * ss // 2
    paths = []
    for label, box in [('top', (0, 0, mask_px_w * ss, half_h)),
                       ('bottom', (0, half_h, mask_px_w * ss, PX_H * ss))]:
        part_alpha = alpha.crop(box)
        part = Image.new('RGBA', part_alpha.size, rgba_tuple(mask_color) + (0,))
        part.putalpha(part_alpha)
        part = part.resize((mask_px_w, PX_H // 2), Image.Resampling.LANCZOS)
        path = os.path.join(out_dir, 'mask-%s.png' % label)
        part.save(path, optimize=True)
        paths.append(path)
    geom = {'overhang': overhang_emu, 'width': int(round(mask_w_emu)),
            'hole_center': overhang_emu + int(round(float(text_x) * SW)),
            'hole_half': hole_half_emu}
    return paths, geom


def add_mask(slide, path, name, x, y, width):
    pic = slide.shapes.add_picture(path, Emu(int(x)), Emu(int(y)),
                                   Emu(int(width)), Emu(SH // 2))
    pic.name = name
    return pic


def add_text(slide, text, y, h, size, font_name, name, color='FFFFFF', bold=False):
    box = slide.shapes.add_textbox(Emu(int(SW * 0.07)), Emu(int(y)),
                                   Emu(int(SW * 0.86)), Emu(int(h)))
    box.name = name
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return box


def add_closing_text(slide, cfg):
    opacity = float(cfg['scrim_opacity'])
    if opacity > 0:
        scrim = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Emu(SW), Emu(SH))
        scrim.name = 'loukong-end-scrim'
        scrim.fill.solid()
        scrim.fill.fore_color.rgb = rgb('#000000')
        set_fill_opacity(scrim, opacity)
        scrim.line.fill.background()
    if cfg['closing_message']:
        add_text(slide, cfg['closing_message'], SH * 0.42, SH * 0.11,
                 cfg['closing_message_size'], cfg['closing_text_font'],
                 'loukong-end-message', bold=True)
    if cfg['closing_subtitle']:
        add_text(slide, cfg['closing_subtitle'], SH * 0.565, SH * 0.07,
                 cfg['closing_subtitle_size'], cfg['closing_text_font'],
                 'loukong-end-subtitle', color='F0F0F0')


def build(cfg, output, config_dir):
    mode = str(cfg['mode']).lower()
    if mode not in ('opening', 'closing'):
        raise ValueError('mode 只能是 opening 或 closing。')
    title = str(cfg['title'] or ('BE MYSELF' if mode == 'opening' else 'THANKS')).strip()
    if not title:
        raise ValueError('title 不能为空。')
    image_path = resolve_path(cfg['image'], config_dir,
                              os.path.join(ASSETS, 'demo-bg.png'))
    if not os.path.isfile(image_path):
        raise OSError('底图不存在：%s' % image_path)
    font_path = find_font(cfg['font_path'], config_dir)
    font_size = cfg['font_size_pt'] or (220 if mode == 'opening' else 295)
    morph_ms = int(cfg['morph_ms'])
    if morph_ms < 100:
        raise ValueError('morph_ms 必须不小于 100。')
    for key in ('text_x', 'text_y'):
        if not 0 <= float(cfg[key]) <= 1:
            raise ValueError('%s 必须在 0～1 之间。' % key)
    if not 0 < float(cfg['text_x']) < 1:
        raise ValueError('text_x 必须在 0～1 之间（不能取 0 或 1）。')
    if not 0 <= float(cfg['initial_offset']) <= 1:
        raise ValueError('initial_offset（页外余量）必须在 0～1 之间。')
    if not 0 <= float(cfg['mask_opacity']) <= 100:
        raise ValueError('mask_opacity 必须在 0～100 之间。')
    if not 0 <= float(cfg['scrim_opacity']) <= 100:
        raise ValueError('scrim_opacity 必须在 0～100 之间。')
    if float(cfg['image_scale']) < 1:
        raise ValueError('image_scale 必须不小于 1。')

    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(SW), Emu(SH)
    prs.core_properties.title = title
    prs.core_properties.subject = '镂空字开场' if mode == 'opening' else '镂空字谢幕'
    prs.core_properties.comments = 'Generated by ppt-loukong-skill; transparent PNG masks.'
    blank = prs.slide_layouts[6]
    page_count = 3 if mode == 'opening' else 2
    with tempfile.TemporaryDirectory(prefix='ppt-loukong-') as tmp:
        masks, geom = make_mask_pair(title, font_path, font_size, cfg['text_x'],
                                     cfg['text_y'], cfg['mask_color'],
                                     cfg['mask_opacity'],
                                     float(cfg['initial_offset']), tmp)
        overhang, mask_w = geom['overhang'], geom['width']
        hole_c, hole_u = geom['hole_center'], geom['hole_half']
        io_emu = int(round(SW * float(cfg['initial_offset'])))
        for index in range(page_count):
            slide = prs.slides.add_slide(blank)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = rgb(cfg['canvas_color'])
            is_reveal = index == page_count - 1
            add_cover_picture(slide, image_path,
                              cfg['image_scale'] if is_reveal else 1.0)
            if mode == 'opening' and index == 0:
                # 首页全黑：上条字孔整孔停在右缘之外，下条停在左缘之外
                top_x = SW + io_emu - (hole_c - hole_u)
                bottom_x = -io_emu - (hole_c + hole_u)
                top_y, bottom_y = 0, SH // 2
            else:
                # 居中页与退场页：字孔落在页面坐标上，只动位置不换图
                top_x = bottom_x = -overhang
                if is_reveal:
                    top_y, bottom_y = -SH // 2, SH
                else:
                    top_y, bottom_y = 0, SH // 2
            add_mask(slide, masks[0], '!!loukong-mask-top', top_x, top_y, mask_w)
            add_mask(slide, masks[1], '!!loukong-mask-bottom', bottom_x, bottom_y, mask_w)
            if mode == 'closing' and is_reveal:
                add_closing_text(slide, cfg)
            if index > 0:
                set_morph_transition(slide, morph_ms)
    output = os.path.abspath(output)
    parent = os.path.dirname(output)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    prs.save(output)
    return {'output': output, 'mode': mode, 'slides': page_count,
            'title': title, 'image': image_path, 'font': font_path,
            'morph_ms': morph_ms}


def main(argv=None):
    parser = argparse.ArgumentParser(description='生成镂空字开场或谢幕 PPTX。')
    parser.add_argument('-o', '--output', help='输出 .pptx 路径')
    parser.add_argument('--config', help='JSON 配置路径')
    parser.add_argument('--mode', choices=['opening', 'closing'], help='覆盖配置中的模式')
    args = parser.parse_args(argv)
    try:
        cfg = dict(DEFAULTS)
        config_dir = os.getcwd()
        if args.config:
            config_path = os.path.abspath(args.config)
            config_dir = os.path.dirname(config_path)
            with open(config_path, 'r', encoding='utf-8-sig') as f:
                loaded = json.load(f)
            if not isinstance(loaded, dict):
                raise ValueError('配置根节点必须是 JSON 对象。')
            unknown = sorted(set(loaded) - set(DEFAULTS))
            if unknown:
                raise ValueError('未知配置项：%s' % ', '.join(unknown))
            cfg.update(loaded)
        if args.mode:
            cfg['mode'] = args.mode
            if not args.config:
                cfg['title'] = 'BE MYSELF' if args.mode == 'opening' else 'THANKS'
        output = args.output or ('loukong-%s.pptx' % cfg['mode'])
        result = build(cfg, output, config_dir)
        print('OK: %(slides)d 页 %(mode)s 已生成：%(output)s' % result)
        print('标题：%(title)s；Morph：%(morph_ms)dms' % result)
        print('底图：%(image)s' % result)
        print('字体：%(font)s' % result)
        return 0
    except Exception as exc:
        print('ERROR: %s' % exc, file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
