---
name: ppt-zhuanpan-skill
description: 生成翻页旋转的大转盘 PPT（Morph 真旋转，风格无关）。左侧大转盘只露右半边，每翻一页转 360/N（N=章数，2~12），当前章标签永远落在右缘 3 点钟方向正对右侧内容区。三种造型：中性圆盘、9 齿齿轮、内外双盘（外环逆时针 + 内盘顺时针反向转，同一张照片填进两个盘，标签沿圆周切向）。造型、章数、配色、照片、背景图、图标、换页音效全部可配，技能不附带任何素材。当用户要"转盘/轮盘/章节转盘 PPT""内外双盘反向旋转""把照片填进转盘""翻页时响一声/换页音效"、点名 ppt-zhuanpan-skill，或要把目录页/章节过渡页做成翻页旋转大转盘时使用。
---

# PPT 翻页旋转转盘（风格无关）

这个技能只负责一件事：把一个"翻页会转的大转盘"放进页面左侧。造型、配色、章数、照片、背景、图标都由配置决定，默认中性无风格、白底可读。已用 python-pptx 生成并经真实 PowerPoint 逐页渲染核验。

## 机制（改脚本前必读）

- **Morph 真旋转**：转盘本体、内盘、分格线、标签、图标、轴孔在每一页都同名（`roulette-gear` / `roulette-inner` / `roulette-spoke-k` / `roulette-label-k` / `chapter-icon-k` / `roulette-hub`），PowerPoint 的 Morph 按对象名补间，所以是"真的在转"，不是换图。
- **落位公式**：第 i 页外环 `offset = i*step`；标签 k 的极角 = `offset - k*step`，于是当前章 k=i 恒为 0°（3 点钟）；分格线 k 的极角 = `offset - (k+0.5)*step`（与标签错开半步）。`step = ±360/N`，ccw 取负、cw 取正。
- **分格线必须是「细矩形 + shape.rotation」**，画普通直线 Morph 不会转。单盘时一端在轮心；双盘时只在环带上（内盘外缘 → 外盘内缘）。
- **标签两种排法**（`label_mode`）：`flat` 永远横排（rotation 恒 0，只随页改位置）；`tangential` 沿圆周切向，`rotation = 极角 + 90`，转到右缘是竖的、转到底下会倒过来（参考视频里的排法）。double 默认 tangential，其余默认 flat。
- **双盘反转**（`wheel: "double"`）：外环 `roulette-gear` 转 `offset`，内盘 `roulette-inner` 转 `-offset × inner_spin`，方向永远相反、幅值默认相同（每页各 360/N）。两个盘同心，内盘半径 = 外盘 × `inner_ratio`。默认不写 `wheel_cx`，轮心由 `inner_edge` 反推（内盘左缘越出页面左边界一点点，被裁掉一线，和参考视频一样）；配置里显式给了 `wheel_cx` 就以它为准。
- **照片填充**：`set_picture_fill(shape, 图片, alpha, crop_frac)` 往 spPr 写 `a:blipFill`，`rotWithShape="1"` 是关键——形状被 Morph 旋转时照片跟着一起转，才是"照片在盘里转"。外环用 `ring_alpha` 压淡，内盘 `inner_alpha` 保持清晰；`photo_match` 给内盘写 `a:srcRect` 中央裁切（比例 = inner_ratio），照片尺度与外环完全一致，看上去是同一张图的延续。
- **图标接力**：图标要加在轴孔/内盘之前；非当前章缩到 `icon_hidden_in` 藏在轮心（单盘被不透明轴孔盖住，双盘被不透明内盘盖住）；当前章放大到轮心右 `icon_offset × r`。每章都给 `icon` 才启用，否则自动关闭。
- **Morph 切换 XML**：`mc:AlternateContent`（Choice 走 p159:morph byObject，Fallback 普通切换），`p14:dur` 毫秒。
- **换页音效**（`sound`）：不是插入音频形状、也不走动画时间轴，而是 PowerPoint 的「切换声音」——在同一个 `p:transition` 里挂 `<p:sndAc><p:stSnd><p:snd r:embed="rIdN" volume="…" name="…"/></p:stSnd></p:sndAc>`，Choice 和 Fallback 各挂一份（这份 XML 形状是照 PowerPoint 自己导出的文件抄的，不是猜的）。音频作为 `RT.AUDIO` 关系嵌进 `/ppt/media/audioN.wav`，同一个文件全篇只存一份（`_audio_part` 按绝对路径缓存）。所以音效天然跟"翻页"这个动作对齐，不需要 `auto_advance_ms`；封面不挂（翻到第一章时盘还没转）。
- **背景默认不设**（PowerPoint 白底）。要背景传 `background` 图片路径；技能不附带任何图片。

## 快速生成

```bash
python scripts/zhuanpan_deck.py -o out.pptx                          # 内置演示（4 章 + 封面）
python scripts/zhuanpan_deck.py -o out.pptx --config deck.json
python scripts/zhuanpan_deck.py -o out.pptx --config examples/double4.json   # 内外双盘 + 照片
python scripts/zhuanpan_deck.py -o out.pptx --config examples/gear9_5.json   # 9 齿齿轮
python scripts/zhuanpan_deck.py -o out.pptx --config deck.json --wheel double --direction cw
python scripts/make_demo_photos.py    # 只为 examples 生成抽象演示照片（正式使用请换成自己的照片）
python scripts/make_demo_sound.py     # 只为 examples 合成机械咔哒换页音效（正式使用请换成自己的音频）
```

依赖 `python-pptx`（make_demo_photos.py 另需 Pillow，make_demo_sound.py 另需 numpy）。配置里的相对路径（photo / icon / background / sound）以配置文件所在目录为基准。

## 配置 JSON

```json
{
  "wheel": "double",
  "direction": "ccw",
  "morph_ms": 1200,
  "inner_ratio": 0.44,
  "ring_alpha": 55,
  "sound": "sounds/click.wav",
  "sound_volume": 60,
  "cover": {"eyebrow": "顶部小字", "title": "主标题", "claim": "一句话主张",
            "target": "副标题", "name_line": "署名行", "teacher": ""},
  "pages": [
    {"part": "PART 01", "section": "第一章", "title": "章标题",
     "points": ["要点1", "要点2"], "photo": "photos/a1.png",
     "icon": "icons/a.png", "num": "01"}
  ],
  "footer_left": "活动名", "footer_name": "姓名",
  "colors": {"label": "FFFFFF", "spoke": "FFFFFF", "title": "1F2429"}
}
```

字段（括号内为默认值；`wheel: "double"` 时带 ★ 的字段自动改用 DOUBLE_DEFAULTS 里的值）：

| 组 | 字段 | 说明 |
| --- | --- | --- |
| 结构 | `pages` / `cover`(null) / `footer_left`("") / `footer_name`("") | 章列表 2~12 个；cover 给对象才出封面 |
| 转盘 | `wheel`(disc) / `wheel_size`★(1.12) / `wheel_cx`★(0.06) / `wheel_cy`(0.50) / `wheel_shadow`(null) / `wheel_line_pt`★(1.25) / `spoke_pt`★(1.4) / `spoke_len`(0.94) / `hub_ratio`★(0.20) | wheel 取 `disc` / `gear9` / `double`；wheel_size 是直径占页高的倍数；double 默认 size 1.17、hub_ratio 0（不画轴孔）、`wheel_cx` 留空由 `inner_edge` 反推轮心 |
| 双盘 | `inner_ratio`(0.44) / `inner_edge`(-0.012) / `inner_spin`(1.0) / `ring_alpha`(55) / `inner_alpha`(100) / `photo_match`(true) / `ring_line_pt`(1.0) / `inner_line_pt`(1.0) / `inner_shadow`(null) | inner_ratio = 内盘半径占外盘比例；inner_edge = 内盘左缘落在页宽的哪一处（负 = 越出页面左边界被裁掉一线；写 0 就正好贴边），仅在没写 `wheel_cx` 时生效；inner_spin 写 2 → 内盘每页转两倍（齿轮咬合感）；alpha 是照片不透明度 % |
| 照片 | 顶层 `photo`(null) / 每章 `pages[].photo` | 照片填充用；double 必须有照片，缺文件直接报错。每章各一张，或只在顶层给一张全篇共用 |
| 旋转 | `direction`(ccw) / `morph_ms`★(1600) / `auto_advance_ms`(null) | ccw 逆时针（外环）、cw 顺时针；内盘永远反向 |
| 音效 | `sound`(null) / `sound_volume`(null) | 换页时响的音频文件路径（wav / mp3 / m4a / mid / midi / aiff）；sound_volume 是 0~100 的百分比，留空按 PowerPoint 默认。只挂在内容页，封面静音 |
| 标签 | `label_mode`(null→自动) / `label_prefix`★(part) / `label_radius`(0.70) / `label_size`★(40) / `label_num_size`(20) / `label_w_in`(1.5) / `label_h_in`(1.1) / `label_tan_w_in`(3.2) / `label_tan_h_in`(0.7) | label_mode 取 `flat` / `tangential`；prefix 写 `""` 只显示编号；tangential 用 tan 那两个尺寸且只有一行「前缀 编号」 |
| 图标 | `icon_size_in`(0.85) / `icon_hidden_in`(0.10) / `icon_offset`★(0.40) | 每章 `pages[].icon` 给图片路径才启用接力；double 默认 0.62（落在环带上） |
| 内容区 | `content_left`★(0.535) / `content_width`★(0.425) / `content_title_size`(28) / `content_point_size`(14) | 右侧固定不动的标题 + 要点区；double 默认 0.49 / 0.46 |
| 其他 | `background`(null) / `font`(微软雅黑) / `footer_left_x`(null) | background 传图片路径；footer_left_x 为 null 时自动避开转盘 |
| 配色 | `colors` 对象 | 键见下；6 位十六进制，写 `""` 表示该项不画 |

`colors` 可用键：`wheel_fill`、`wheel_fill2`（渐变第二色，`""` = 纯色）、`wheel_line`、`spoke`、`hub_fill`、`hub_line`、`ring_line`（双盘外环描边）、`inner_line`（双盘内盘描边）、`label`、`label_num`、`title`、`text`、`accent`、`dim`。
默认是白底可读的中性浅灰盘；`gear9` 自动叠金属渐变；`double` 自动把 `label / spoke / ring_line / inner_line` 改成白色（照片上白字最稳）、不画轴孔——都仍可被 `colors` 覆盖。

## 边界

- 章数 2~12。N>6 时标签开始变挤：调小 `label_size` / `label_w_in`（或 tangential 的 `label_tan_w_in`），或调大 `label_radius`。
- `wheel: "double"` 必须有照片：每章 `photo` 或顶层 `photo`，缺文件直接 `ERROR:` 退出。照片本身由用户提供，技能不附带。
- 照片是"填充"不是"素材贴图"：形状旋转时照片一起转（rotWithShape），所以内盘里的照片方向会随页变化；这是设计效果，不是 bug。
- `tangential` 标签转到页面下方会倒过来（和参考视频一致）。要一直好读就写 `"label_mode": "flat"`。
- 深色 `background` 或暗照片上，记得同时把 `colors` 的 `title` / `text` / `dim` 调亮，否则右侧正文看不清。
- 想整体左右挪盘：在配置里写 `wheel_cx`（占页宽比例）即可覆盖自动反推；只想调"内盘被左边界裁多少"就改 `inner_edge`（`0` = 正好相切，`-0.03` = 切掉更多）。
- 未知配置项、非法颜色、非法 `label_mode`、`inner_ratio` 越界、`inner_edge` 越界、`sound_volume` 越界、缺照片/图标/背景/音效文件或音效扩展名不支持都会直接 `ERROR:` 报错退出，不会静默出错。
- 音效只在**放映**时响（切换声音本来就是放映属性），编辑视图里翻页面是静音的；WPS / 旧版 PowerPoint 可能忽略这段 XML，文件不会坏。想听效果就用 PowerPoint 放映。
- 技能不附带任何真实音频；`examples/sounds/*.wav` 由 `make_demo_sound.py` 用 numpy 现合成（`click.wav` 一下落位、`ratchet.wav` 一串由密到疏的咬合），可重建、无版权。别把别人视频/歌里的音频塞进配置再公开分发。
- 文字都是真文本框，生成后可以在 PowerPoint 里改字（这点和 ppt-loukong-skill 的透明字孔不同）。
- 第一页（封面）不带进入切换；内容页全部带 Morph。
- 想要这三种以外的造型：改 `add_wheel()` 里的 `MSO_SHAPE` 或加一个分支，其余机制不用动。

## QA 检查单

1. python-pptx 重开无异常；内容页 `roulette-gear` 的 rotation 逐页递增 `±360/N`；double 时 `roulette-inner` 逐页递增 ∓（与外环符号相反）；每页有 Morph 的 `mc:AlternateContent` 且 `p14:dur` == `morph_ms`。
2. 当前章标签中心与轮心同高、位于轮心右侧（3 点钟）；flat 时 `rotation == 0`，tangential 时 `rotation == (极角 + 90) % 360`；ccw 时上一章在正上方、下一章在正下方。
3. 双盘再核：两盘同心、内盘半径 = 外盘 × `inner_ratio`；两个盘的 `a:blipFill` 都带 `rotWithShape="1"` 且 `r:embed` 指向同一张图片；外环有 `a:alphaModFix`、内盘有 `a:srcRect`；分格线中点半径 > 内盘半径（没伸进内盘）。
4. 用真实 PowerPoint 渲染逐页目检：pywin32 晚期绑定 `win32com.client.Dispatch('PowerPoint.Application')`，`Slides.Item(i).Export(path,'PNG',1600,900)`（.NET/PowerShell COM 在此机器报 TYPE_E_CANTLOADLIBRARY）。
5. 带 `sound` 时再核：包里有 `/ppt/media/audioN.ext` 且字节与源文件一致；每个内容页有 1 条 `RT.AUDIO` 关系、`p:sndAc` 在 Choice 与 Fallback 各一份且 `r:embed` 指到该关系；封面 0 份。用 COM 打开应无修复提示，`Slides(i).SlideShowTransition.SoundEffect.Name` 读回文件名（未设置时该属性直接报错，属正常）。python-pptx 把 wav 写成 `[Content_Types].xml` 的 `Override`（PowerPoint 自己写 `Default`），实测 PowerPoint 照读、另存时自动改成 Default。
6. 提醒用户：Morph 需 PowerPoint 2019/365 放映；WPS/旧版回退普通切换，文件不会坏。音效同样只在放映时响。
