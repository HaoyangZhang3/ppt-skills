---
name: ppt-loukong-skill
description: 生成照片填充的镂空字 PPT 开场和电影式谢幕，使用上下独立透明字孔遮罩、底图放大与 Morph 平滑切换。用户要求“镂空字/文字镂空/照片透字/字孔遮罩 PPT”“BE MYSELF 式开场”“THANKS 式谢幕”，或点名 ppt-loukong-skill 时使用。支持 3 页开场（黑场→字孔自两侧滑入居中照片镂空标题→上下退场揭晓）与 2 页谢幕（THANKS→上下滑出→底图放大及结束语），可配置标题、底图、字体、文案和时长。
---

# PPT 镂空字开场与谢幕

用自包含脚本复现两种视频版式。字孔不是文字填充，而是黑色遮罩中的透明区域；移动遮罩后，孔洞透出白色画布或底层照片。

## 快速生成

在技能目录执行：

```bash
python scripts/loukong_deck.py -o opening.pptx --config examples/opening.json
python scripts/loukong_deck.py -o closing.pptx --config examples/closing.json
```

不传配置时生成内置 `BE MYSELF` 开场；`--mode closing` 生成 `THANKS` 谢幕。依赖 `python-pptx` 与 Pillow，不需要 PowerPoint、COM 或外部技能。正式作品将 `image` 换为用户底图；自带 `assets/demo-bg.png` 是程序生成的抽象山景，不是照片。

## 固定动画结构

- **opening，3 页**：①黑场开页——遮罩比页面两侧各宽出一段余量，上条字孔整孔停在页面右缘之外、下条停在左缘之外，页面全黑；②字孔随 Morph 从两侧滑入页面中央，标题孔洞露出照片；③上条向上、下条向下完全退出，底图放大。第 2、3 页设 Morph。
- **closing，2 页**：①`THANKS` 字孔居中并露出照片；②两块遮罩上下退出，底图放大并出现结束语。第 2 页设 Morph。
- 默认切换 **2000ms**，底图末页放大至 **1.10**。跨页底图与两张遮罩使用相同图像内容及 `!!loukong-bg`、`!!loukong-mask-top`、`!!loukong-mask-bottom` 固定名称，确保 Morph 按对象补间。
- 标题默认使用 Impact（220pt 开场、295pt 谢幕），自动缩小以避免超出版面——`text_x` 偏离中线时字宽上限随之收紧，标题始终完整落在页面内。中文或 Impact 不可用时回退到系统粗体，也可用 `font_path` 指定字体文件。

## 配置

JSON 根节点为对象；`image` 和 `font_path` 的相对路径以配置文件所在目录为基准。常用字段：

```json
{
  "mode": "opening",
  "title": "BE MYSELF",
  "image": "photo.jpg",
  "font_path": "C:/Windows/Fonts/impact.ttf",
  "font_size_pt": 220,
  "text_x": 0.5,
  "text_y": 0.5,
  "mask_color": "#000000",
  "mask_opacity": 100,
  "canvas_color": "#FFFFFF",
  "initial_offset": 0.06,
  "image_scale": 1.1,
  "morph_ms": 2000,
  "closing_message": "愿你我更上一层楼",
  "closing_subtitle": "感谢各位的观看和支持，您的关注是我持续更新的动力",
  "closing_text_font": "微软雅黑",
  "closing_message_size": 34,
  "closing_subtitle_size": 14,
  "scrim_opacity": 28
}
```

`mode` 必填语义为 `opening` 或 `closing`。`closing_message` / `closing_subtitle` 设为空字符串可隐藏对应文案；`scrim_opacity` 控制谢幕结束页文字后的黑色压暗层，0 为关闭。标题支持 `\n` 多行。`initial_offset` 是字孔的页外余量（占页宽比例，0～1，默认 0.06）：首页把上条字孔推到右缘外、下条推到左缘外，页面因此全黑，字母只随 Morph 滑入。

## 实现边界与修改要点

- 字孔采用 2× 超采样 Pillow RGBA PNG，生成后嵌入 PPTX；**标题在 PowerPoint 中不是可编辑文字**。修改标题、字体或字孔位置时改配置并重跑，不要手动拖拽单页遮罩代替重生成。
- 遮罩是宽幅图（约 3 倍页宽），首页之后位置恒为 x=-单侧余量；首页全黑是设计使然——若首页能看到字母（哪怕错位），说明遮罩被手动挪动或用旧版生成，应重跑脚本而不是继续改。
- 两半遮罩必须保持独立；不要把它们合并。末页必须只改变相同对象的位置/大小，不能替换底图或遮罩内容。
- 不要给第一页添加“进入本页”的 Morph；它应设置在后续页面。旧版 PowerPoint / WPS 可能回退为普通切换。
- 自带素材需要重建时运行 `python scripts/make_demo_asset.py`。

## QA

```bash
python scripts/verify_deck.py opening.pptx --mode opening
python scripts/verify_deck.py closing.pptx --mode closing
```

核验脚本检查页数、透明字孔、跨页图像哈希、固定对象名、开场首页字孔整孔在页面之外且遮罩盖满页面、居中/退场位置、底图放大、Morph 类型与 2000ms 时长，以及谢幕结束语。交付前还应用真实 PowerPoint 导出各页 PNG 目检，重点看首页是否全黑、标题是否完整、照片是否只从字孔透出、结束语是否清晰；静态核验不能证明动画播放丝滑度。
