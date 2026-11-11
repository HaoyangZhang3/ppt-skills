---
name: ppt-xqmsaf-skill
description: 生成《喜鹊谋杀案》式悬疑暗场风 PPT（黑底+满页红蒙版压出深酒红、淡黄灯源与柔光锥、光锥照亮图片、开灯擦除动画、文字透明度 74%→0% 的聚光灯转移、跨页平滑 Morph 与"大图移出画外"的丝滑衔接）。当用户要"喜鹊谋杀案风/悬疑暗场/酒红复古/灯光 PPT"、点名 ppt-xqmsaf-skill，或要把演示做成暖光暗调的侦探小说质感时使用。
---

# PPT 喜鹊谋杀案风（暗场·酒红·暖光）

复刻自一段抖音教程的三大手法：**灯光**（灯源+柔光锥+红蒙版+光照图+开灯动画+聚光灯转移）、**图片处理**（AI 生成同风格图，叠红蒙版统一色调）、**页面衔接**（全页 Morph + 跨页同名对象 + 大图移出画外）。生成后可直接在 PowerPoint 播放：翻页平滑滑入、灯"啪"地扫亮、聚光灯跨页游走。

## 机制（改脚本前必读）

- **图层次序**（每页固定）：黑底 `xq-bg` → 底图 `xq-hero` → 红蒙版 `xq-wash` → 光照图 `xq-lit` → 光锥 `xq-cone` → 光晕 `xq-glow` → 灯源 `xq-lamp-a/b` → 文字。红蒙版盖在底图上，所以图片天然吃暗调；光照图在蒙版之上，所以只有光锥内是亮的。
- **红蒙版**：满页矩形，填充红 `A50004`、透明度 `wash_alpha`(47%)，黑底上合成近似 **#520002** 深酒红（视频成品实测 #450001，想更深把 `wash_alpha` 调到 55+ 或换更深红）。
- **光锥**：`PIE` 扇形（视频里"被咬了一口的饼形"），`adjustments = ±spread/2`、`rotation = dir`；填充淡黄 `FFF2CC`、不透明度 `100-cone_alpha`(55%)；`a:softEdge` ≈ 半径 6% = 视频的"把角磨平"。默认几何：`cover` 灯在右上、锥朝左下；`content` 灯在左上、锥朝右照图。
- **灯源**：一大一小两个淡黄圆叠放（大圆 62% 不透明+柔边，小圆 96% 偏移左上），对应视频"插入圆形→淡黄→无线条→复制→调整大小"。
- **光晕**：压扁椭圆（宽 4.6r、高 1.15r），`rotation` 对齐光锥方向，低温淡黄+大柔边 = "开灯一闪"的扁圆。
- **光照图（关键）**：视频用"图片+光锥→合并形状→相交"。这里用 Pillow 等价实现——按灯心/方向/张角在图片上抠一个**羽化扇形 alpha 蒙版**，生成 RGBA PNG 覆盖在原图同一位置，肉眼等价且不依赖 PowerPoint UI。底图与光照图用同一 cover 裁剪（`crop_left/...` 与 Pillow cover-fit），保证像素对齐。
- **开灯动画**：光锥+光晕（有图时含光照图）加 **擦除·自顶部**（`wipe(down)`，presetID 22 / subtype 4）、`wipe_ms`(0.5s)、同组同时播放。timing 组用 `onBegin tn=2` 条件 → **翻到该页自动开灯**，不用点。
- **随机线条**：`randombar(horizontal)`，presetID 14 / subtype 10，给标题加"报幕"出场（页配置 `"title_reveal": "bars"`），延迟 `wipe_ms` 后自动播。
- **聚光灯转移**：跨页同名对象（灯/锥/光晕/文字）配 Morph 自动补间；文字用 `text_transparency`（PowerPoint 语义：**74 = 74% 透明**）——上一页 74、下一页 0，翻页时灯移过去、文字亮起来（视频原话"把文字透明度改成 74%→调成 0%"）。
- **Morph 衔接**：除第 1 页外每页挂 `p159:morph byObject`（`morph_ms`），跨页同名对象自动补间。三招配方（大图移出画外、灯光藏进灯源、聚光灯转移）见下方「页面衔接（Morph）」一节。
- **图片处理路线**：视频用"百度搭子"把照片转成同风格（暗调、酒红、版画/油画质感）。本技能接收任意图片；没有 AI 时用 `make_demo_assets.py` 生成抽象演示图，正式使用请换成 AI 生成或实拍处理的图。

## 快速生成

```bash
python scripts/build_xqmsaf.py -o out.pptx                                # 内置演示（无图片）
python scripts/build_xqmsaf.py -o out.pptx --config examples/demo.json    # 带图完整演示
python scripts/make_demo_assets.py                                        # 只为 examples 生成抽象演示图
python scripts/verify_deck.py out.pptx                                    # 静态核验
```

依赖 `python-pptx`、`Pillow`。配置里的相对路径（image）以配置文件所在目录为基准。生成 PPT 只有第 1 页没有进入切换，其余页全部 Morph。

## 配置 JSON

```json
{
  "morph_ms": 1400, "wipe_ms": 500, "bars_ms": 800,
  "wash_alpha": 47, "cone_alpha": 45, "feather": 0.035,
  "font_title": "STZhongsong", "font_body": "宋体", "font_en": "Georgia",
  "palette": {"bg": "000000", "wash": "A50004", "lamp": "FFF2CC",
              "cream": "E8D2CF", "dim": "8E6A62", "accent": "B2554A", "red": "9E0007"},
  "pages": [
    {"type": "cover", "kicker": "PRESENTATION", "title": "标题", "en": "ENGLISH",
     "sub": "副标题", "image": "photos/cover.jpg", "image_box": [0, 0, 1, 1]},
    {"type": "section", "num": "Ⅰ", "kicker": "PART ONE", "title": "章节", "en": "EN"},
    {"type": "toc", "kicker": "Ⅱ · 目录", "title": "目录",
     "body": ["一行一条"], "image": "photos/a.jpg"},
    {"type": "content", "title": "正文页", "body": ["..."],
     "image": "photos/a.jpg", "image_box": [0.52, 0.085, 0.42, 0.83],
     "lit": true, "text_transparency": 74},
    {"type": "end", "title": "第一幕 · 终", "sub": "一句话", "en": "END"}
  ]
}
```

页字段（未知字段直接 `ERROR:`）：

| 字段 | 说明 |
| --- | --- |
| `type` | `cover` / `section` / `content` / `toc` / `end` |
| `kicker` / `title` / `en` / `sub` / `num` | 小标签 / 主标题（`\n` 分行）/ 英文行 / 副标题 / 章节罗马数字 |
| `body` | content 页正文，字符串数组，一行一条 |
| `image` / `image_box` | 图片路径；页面占比 `[x, y, w, h]`（0~1）。默认 content `[0.52, 0.085, 0.42, 0.83]` |
| `lit` | `true` = 生成"光锥照亮图片"的光照图（需要 image） |
| `lamp` | `[x, y, r]`，灯心占页宽/页高比例，r 相对页高。默认见下表 |
| `cone` | `{dir, spread, length}`：方向角（0=右，顺时针）、张角、长度（占页宽）。`length: 0.02` = 灯光藏进灯源 |
| `glow` | 是否显示光晕（默认 cover/section/content 开，end 关；关掉也保留同名小形状给 Morph） |
| `light_on` | 是否加开灯擦除动画（默认 cover/section 开） |
| `title_reveal` | `"bars"` = 标题加随机线条出场 |
| `text_transparency` | 文字透明度 %（74 = 深暗，0 = 全亮），content 页生效，配合跨页做聚光灯转移 |
| `title_size` / `body_size` | 覆盖默认字号 |

页型默认灯光几何（lamp 可整体覆盖）：

| type | lamp | cone（dir/spread/length） | light_on |
| --- | --- | --- | --- |
| cover | 0.705, 0.255, 0.078 | 118° / 66° / 1.55 | ✓ |
| section | 0.80, 0.285, 0.062 | 148° / 60° / 1.50 | ✓ |
| content | 0.285, 0.265, 0.055 | 15° / 50° / 1.00 | ✗ |
| toc | 0.285, 0.265, 0.055 | 15° / 50° / 0.02（灯先藏起，等翻页展开） | ✗ |
| end | 0.50, 0.40, 0.048 | 90° / 52° / 0.02（藏灯） | ✗ |

## 目录页（TOC）

`type: "toc"` 就是"目录"页型：文字布局同 content（kicker 写"目录"、`body` 一行一个条目、字号自动大一号），但它**默认就是衔接的起点**——光锥缩进灯源（`cone.length: 0.02`）、大图默认放在画外（`image_box: [0.90, 0.085, 0.42, 0.83]`）。视频里"目录部分的灯光和衔接怎么弄"的完整做法：

1. 做好目录页（文字列表 + 灯源，灯光先藏在灯里）
2. **复制目录页**当下一张内容页的底稿
3. 新页：**拉大图片**（`image_box` 摆正到 `[0.52, …]`）、**文字字号调大**（`title_size`）、**加上内容**（`body`）、光锥恢复正常长度
4. **回到目录页**：把同一张大图**粘贴到画面外**（toc 的默认位置）
5. 两页之间**平滑切换**：播放时大图从画外滑入、灯从灯源里张开

图层里"选中要讲解的部分置于顶层、其余置于底层"在生成器里已由固定图层顺序实现（底图 → 红蒙版 → 光照图 → 光锥 → 灯 → 文字；内容页 `lit` + 灯光对准即"点亮要讲解的部分"），配置里不需要手工调层。

```json
{"type": "toc", "kicker": "Ⅱ · 目录", "title": "三座塔",
 "body": ["Ⅰ 燃灯塔", "Ⅱ 铁塔", "Ⅲ 瑞光塔"], "image": "photos/tower.jpg"},
{"type": "content", "title": "燃灯塔", "body": ["…"],
 "image": "photos/tower.jpg", "lit": true, "text_transparency": 74}
```

目录页不点灯、不出光照图；灯的亮相留给下一页 Morph 展开。

## 页面衔接（Morph）

技能默认就带着"页面衔接"：除第 1 页外每页都挂**平滑切换**（`p159:morph byObject`，`morph_ms` 默认 1400ms）。跨页自动补间的是**同名 `xq-*` 对象**——上一页的灯、光锥、光晕、图片、文字到下一页会"滑"过去而不是换图。衔接不需要额外开关，把两页参数写成"同一组物件、不同位置/大小/透明度"即可。三招配方：

**① 大图移出画外 → 滑入（目录页 → 内容页，视频的"丝滑衔接"）**

两页引用**同一张图片**（同路径），目录页（`toc`）默认把图放在画外，内容页摆正即可：

```json
{"type": "toc", "title": "目录", "image": "photos/tower.jpg"},
{"type": "content", "title": "燃灯塔", "image": "photos/tower.jpg",
 "image_box": [0.52, 0.085, 0.42, 0.83], "lit": true}
```

播放时大图从右侧滑入占位，灯同步从灯源里张出来。**要点：两页都要有 image**，缺一页就退化成淡入淡出，不是滑动；不用 toc 页型时，上一页手写 `image_box: [0.90, 0.085, 0.42, 0.83]` 效果相同。

**② 把灯光藏进灯源 / 灯游走**

上一页 `cone: {"length": 0.02}`（扇形缩到灯后，看起来收进灯里），下一页恢复正常 `length` 并给新的 `lamp` / `cone.dir`，Morph 会把光锥"拉出来"、灯"走"过去：

```json
{"type": "content", "lamp": [0.285, 0.265, 0.055], "cone": {"dir": 15}},
{"type": "content", "lamp": [0.30, 0.62, 0.055], "cone": {"dir": 345}, "lit": true}
```

**③ 聚光灯转移（文字透明度 74 → 0）**

同一组文字跨页，上一页 `text_transparency: 74`（暗）、下一页 `0`（亮），同时挪灯；有图的页开 `lit`，光移过去照亮图片：

```json
{"type": "content", "title": "燃灯塔", "text_transparency": 74, "lit": true},
{"type": "content", "title": "铁塔",   "text_transparency": 0,
 "lamp": [0.30, 0.62, 0.055], "cone": {"dir": 345}, "lit": true}
```

其它可跨页补间的：`title_size`/`body_size` 变化（文字缩放，对应视频"字号调大"）、标题/英文行改字（内容交叉淡化）、`lamp` 半径变化（灯变大变小）。注意 Morph 需 PowerPoint 2019/365 放映，WPS/旧版回退普通切换；`verify_deck.py` 会检查第 2 页起的 Morph 是否齐全。

## 边界与坑

- **Morph 需要 PowerPoint 2019/365** 放映；WPS/旧版回退普通切换，文件不会坏。擦除动画所有版本都支持。
- 光照图的羽化扇形是"合并形状→相交"的离线等价：**要在 PowerPoint 里手动改光形请直接改图片或重跑脚本**，不要幻想改光锥形状能带着光照图走。
- 图片建议长边 ≥ 1200px；`image_box` 与图片比例差太多会 cover 裁切（与底图一致，不会错位）。视频的做法是先用 AI 把照片转成同风格再放进 PPT，本技能只负责"放进去之后"的暗场处理。
- 文字是普通文本框可直接改；`text_transparency` 是 PowerPoint"文字透明度"语义，别和动画透明度混淆。
- 中文衬线首选 `STZhongsong`（华文中宋，Win 自带），没有就换 `宋体`/`Source Han Serif`；英文默认 Georgia。
- 页数 1~12。长标题会挤：调小 `title_size` 或拆 `\n`；正文超过 6 行请拆页。
- 深色页上别用纯黑字；`palette` 里 `cream/dim/accent` 都可改，但先跑 `verify_deck.py` 再渲染看效果。

## 参考资料

- `references/creative-method.md` —— 视频原版创作方法全记录（PowerPoint 原生操作步骤、风格量化档案、与本技能的参数对照表）。想手工在 PowerPoint 里复刻、或想改机制前先读它。

## QA 检查单

1. `python scripts/verify_deck.py out.pptx` → PASS：图层齐全、wash alpha = (100-47)%、cone alpha = 55%、光锥/光晕有柔边、第 2 页起有 Morph（1400ms）、有 timing 的页含 `wipe(down)`。
2. 真实 PowerPoint 打开无修复提示；`Slides(i).Export(...,'PNG',1600,900)` 导出逐页看：
   - 成品页平均亮度约 **25~65/255**（暗场），最亮点应接近 **238+**（灯心）；
   - 内容页"亮区"应对应光锥方向（如灯在左上、图在右，则右侧更亮）；
   - 灯/锥/光晕不得压到标题文字；正文不溢出页面。
3. 放映检查：第 2 页起翻页是 Morph 滑入；cover/section 进入后自动"开灯"（擦除 0.5s）；聚光灯对页（74→0）灯位移动、文字变亮；目录页→内容页大图从画外滑入。
4. 提醒用户：动画和 Morph 只在真实 PowerPoint 放映时可见，静态缩略图看不出。