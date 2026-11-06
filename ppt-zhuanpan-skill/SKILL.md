---
name: ppt-zhuanpan-skill
description: 生成翻页旋转的大转盘 PPT（Morph 真旋转，风格无关）。左侧大转盘只露右半边，每翻一页整盘转 360/N（N=章数，2~12），当前章横排「前缀+编号」标签永远落在右缘 3 点钟方向正对右侧内容区，标签旁图标翻页时缩回轮心/放大伸出接力。默认中性圆盘+白底；造型（圆盘/齿轮）、章数、配色、背景图、图标全部可配，技能不附带任何素材。当用户要"转盘/轮盘/章节转盘 PPT"、点名 ppt-zhuanpan-skill、或要把目录页/章节过渡页做成翻页旋转大转盘时使用。
---

# PPT 翻页旋转转盘（风格无关）

这个技能只负责一件事：把一个"翻页会转的大转盘"放进页面左侧。造型、配色、章数、背景、图标都由配置决定，默认中性无风格。已用 python-pptx 生成并经真实 PowerPoint 逐页渲染核验。

## 机制（改脚本前必读）

- **Morph 真旋转**：转盘本体、分格线、标签、图标、轴孔在每一页都同名（`roulette-gear` / `roulette-spoke-k` / `roulette-label-k` / `chapter-icon-k` / `roulette-hub`），PowerPoint 的 Morph 按对象名补间，所以是"真的在转"，不是换图。
- **落位公式**：第 i 页整盘 `offset = i*step`；标签 k 的极角 = `offset - k*step`，于是当前章 k=i 恒为 0°（3 点钟）；分格线 k 的极角 = `offset - (k+0.5)*step`（与标签错开半步）。`step = ±360/N`，ccw 取负、cw 取正。
- **分格线必须是「一端在轮心的细矩形 + shape.rotation」**，画普通直线 Morph 不会转。
- **标签永远横排**（rotation 恒 0），只随页改位置、不改角度。
- **图标接力**：图标要加在轴孔之前；非当前章缩到 `icon_hidden_in` 藏在轮心，被不透明轴孔盖住；当前章放大到轮心右 `icon_offset × r` 处（标签左侧）。每章都给 icon 才启用，否则自动关闭。
- **Morph 切换 XML**：`mc:AlternateContent`（Choice 走 p159:morph byObject，Fallback 普通切换），`p14:dur` 毫秒。
- **背景默认不设**（PowerPoint 白底）。要背景传 `background` 图片路径；技能不附带任何图片。

## 快速生成

```bash
python scripts/zhuanpan_deck.py -o out.pptx                          # 内置演示（4 章 + 封面）
python scripts/zhuanpan_deck.py -o out.pptx --config deck.json
python scripts/zhuanpan_deck.py -o out.pptx --config examples/gear9_5.json
python scripts/zhuanpan_deck.py -o out.pptx --config deck.json --wheel gear9 --direction cw
```

依赖仅 `python-pptx`。配置里的相对路径（icon / background）以配置文件所在目录为基准。

## 配置 JSON

```json
{
  "cover": {"eyebrow": "顶部小字", "title": "主标题", "claim": "一句话主张",
            "target": "副标题", "name_line": "署名行", "teacher": ""},
  "pages": [
    {"part": "PART 01", "section": "第一章", "title": "章标题",
     "points": ["要点1", "要点2"], "icon": "icons/a.png", "num": "01"}
  ],
  "footer_left": "活动名", "footer_name": "姓名",
  "morph_ms": 1600, "direction": "ccw", "wheel": "disc",
  "colors": {"wheel_fill": "E6E9ED", "accent": "5A7186"}
}
```

字段（括号内为默认值）：

| 组 | 字段 | 说明 |
| --- | --- | --- |
| 结构 | `pages` / `cover`(null) / `footer_left`("") / `footer_name`("") | 章列表 2~12 个；cover 给对象才出封面 |
| 转盘 | `wheel`(disc) / `wheel_size`(1.12) / `wheel_cx`(0.06) / `wheel_cy`(0.50) / `wheel_shadow`(null) / `wheel_line_pt`(1.25) / `spoke_pt`(1.4) / `spoke_len`(0.94) / `hub_ratio`(0.20) | wheel 取 disc 或 gear9；wheel_size 是直径占页高的倍数；wheel_cx 0.06 → 只露右半边 |
| 旋转 | `direction`(ccw) / `morph_ms`(1600) / `auto_advance_ms`(null) | ccw 逆时针、cw 顺时针 |
| 标签 | `label_prefix`(part) / `label_radius`(0.70) / `label_size`(40) / `label_num_size`(20) / `label_w_in`(1.5) / `label_h_in`(1.1) | prefix 写 `""` 则只显示编号 |
| 图标 | `icon_size_in`(0.85) / `icon_hidden_in`(0.10) / `icon_offset`(0.40) | 每章 `pages[].icon` 给图片路径才启用接力 |
| 内容区 | `content_left`(0.535) / `content_width`(0.425) / `content_title_size`(28) / `content_point_size`(14) | 右侧固定不动的标题 + 要点区 |
| 其他 | `background`(null) / `font`(微软雅黑) / `footer_left_x`(null) | background 传图片路径；footer_left_x 为 null 时自动避开转盘 |
| 配色 | `colors` 对象 | 键见下；6 位十六进制，写 `""` 表示该项不画 |

`colors` 可用键：`wheel_fill`、`wheel_fill2`（渐变第二色，`""` = 纯色）、`wheel_line`、`spoke`、`hub_fill`、`hub_line`、`label`、`label_num`、`title`、`text`、`accent`、`dim`。
默认是白底可读的中性浅灰盘；`wheel: "gear9"` 时自动换成金属感渐变（仍可被 colors 覆盖）。

## 边界

- 章数 2~12。N>6 时标签开始变挤：调小 `label_size` / `label_w_in`，或调大 `label_radius`。
- 未知配置项、非法颜色、缺图标/背景文件都会直接 `ERROR:` 报错退出，不会静默出错。
- 文字都是真文本框，生成后可以在 PowerPoint 里改字（这点和 ppt-loukong-skill 的透明字孔不同）。
- 第一页（封面）不带进入切换；内容页全部带 Morph。
- 想要圆盘/齿轮以外的造型：改 `add_wheel()` 里的 `MSO_SHAPE` 或加一个分支，其余机制不用动。

## QA 检查单

1. python-pptx 重开无异常；内容页 `roulette-gear` 的 rotation 逐页递增 `±360/N`；每页有 Morph 的 `mc:AlternateContent` 且 `p14:dur` == `morph_ms`。
2. 当前章标签中心与轮心同高、位于轮心右侧（3 点钟），且 `rotation == 0`；ccw 时上一章在正上方、下一章在正下方。
3. 用真实 PowerPoint 渲染逐页目检：pywin32 晚期绑定 `win32com.client.Dispatch('PowerPoint.Application')`，`Slides.Item(i).Export(path,'PNG',1600,900)`（.NET/PowerShell COM 在此机器报 TYPE_E_CANTLOADLIBRARY）。
4. 提醒用户：Morph 需 PowerPoint 2019/365 放映；WPS/旧版回退普通切换，文件不会坏。
