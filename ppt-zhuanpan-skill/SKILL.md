---
name: ppt-zhuanpan-skill
description: 生成机械风大齿轮转盘 PPT（Morph 平滑旋转）。左侧大齿轮只露右半边，每翻一页逆时针转 90°，当前章横排「part+数字」标签永远落在右缘 3 点钟方向正对右侧内容区，标签旁工具图标翻页时缩回轮心/放大伸出接力，暗灰钢色、常驻转动齿轮 GIF。当用户要求"转盘/轮盘/齿轮转盘 PPT"、点名用 ppt-zhuanpan-skill、或要把目录/章节页做成翻页旋转的大转盘效果时使用。适合目录页、章节过渡页、4 章结构汇报（职业目标/成长实践/未来发展/结尾展望等）。
---

# PPT 大齿轮转盘（机械风 Morph 轮盘）

从实战汇报 PPT 沉淀的完整配方，已用 python-pptx 验证并逐页渲染检查。

## 效果规格（默认即最终定稿样式，勿随意改）

- **大轮盘只露右半边**：GEAR_9 金属渐变齿轮，直径 1.12 倍页高，圆心在 0.06SW / 0.50SH，可见部分约 1/4 版面；中心有轴孔、4 根过圆心放射线与标签错开 45°。
- **每页逆时针转 90°**：4 页齿轮 rotation 依次为 0°→270°→180°→90°。
- **当前章标签永远落在右缘 3 点钟方向**：横排两段式（小写 part 40pt 白色 + 数字 20pt 浅灰、段前 6pt），全部水平不旋转；上一章在正上方、下一章在下方，均横排。标签基准角 k*90°、半径 0.70r。
- **工具图标接力动画**：每章一个钢色线性图标（pen/wrench/pliers/paper），当前章图标 0.85in 在标签左侧（cx+0.40r），其余章缩成 0.10in 藏在轮心轴孔下。Morph 时旧图标缩小收回轮心、新图标自轮心放大伸出。
- **右侧内容区固定不动**：part·章节小标签（钢蓝）+ 28pt 大标题 + 钢色短横线 + 3 条要点。
- **常驻转动齿轮 GIF**：封面一对大齿轮（0.80SW/0.40SH），内容页右上一对小齿轮（0.885SW/0.15SH）。
- 暗灰机械风背景（网格+刻度尺+齿轮暗纹+四角螺丝），页脚避开齿轮（内容页左页脚从 0.30SW 起）。
- 全部页面带 Morph（平滑）切换，默认 1600ms。

## 快速生成

```bash
python scripts/zhuanpan_deck.py -o out.pptx                  # 内置演示（封面+4章）
python scripts/zhuanpan_deck.py -o out.pptx --config deck.json
```

依赖仅 `python-pptx`；素材（背景 PNG、齿轮 GIF、4 个图标 PNG）已随技能附带在 `assets/`。

## 配置 JSON

```json
{
  "footer_left": "赛事/活动名称 · 赛道（组别）",
  "footer_name": "姓名",
  "morph_ms": 1600,
  "direction": "ccw",
  "cover": {
    "eyebrow": "顶部小字", "title": "生涯发展展示",
    "claim": "一句话核心主张。", "target": "目标职业：……",
    "name_line": "姓名 · 学校 · 专业班级", "teacher": "指导教师：……"
  },
  "pages": [
    {"part": "PART 01", "section": "职业目标", "title": "从好奇心，到一张图纸",
     "points": ["要点1", "要点2", "要点3"], "icon": "pen"},
    {"part": "PART 02", "section": "成长实践", "title": "……", "points": ["……"], "icon": "wrench"},
    {"part": "PART 03", "section": "未来发展", "title": "……", "points": ["……"], "icon": "pliers"},
    {"part": "PART 04", "section": "结尾展望", "title": "……", "points": ["……"], "icon": "paper"}
  ]
}
```

- 固定 **4 章**（90° 步长 × 4 扇区）；`cover` 设为 `null` 可去掉封面。
- `icon` 可选 `pen/wrench/pliers/paper`；4 页都给了才启用图标接力，否则自动关闭。
- `direction: "cw"` 可改顺时针（标签落点不变，转动方向相反）。
- 文案原则：**少字大字**——每页 1 句核心主张 + 2~3 条短要点。

## 修改脚本时的关键实现点

- **Morph 真旋转**：放射线必须用过圆心细矩形 + `shape.rotation`，齿轮/放射线/标签/图标跨页同名（roulette-gear、roulette-spoke-k、roulette-label-k、chapter-icon-k），Morph 按对象名匹配。
- **标签落位公式**：标签 k 基准角 = k*90°，第 i 页 offset = −i*90°，则当前章 k=i 恒在 0°（3 点钟）。标签文本框 rotation 恒为 0（横排）。
- **图标藏轮心**：图标加在齿轮之后、轴孔之前，轴孔（实心深色圆）盖住缩小的非当前图标。
- Morph 切换 XML：`mc:AlternateContent`（Choice Requires p159:morph byObject + Fallback 普通 transition），`p14:dur` 毫秒。
- 背景图填充：`p:bg/p:bgPr/a:blipFill` 写在 `p:cSld` 第一个子元素。
- 重新生成素材（需 Pillow）：`python scripts/make_assets.py`（背景+齿轮 GIF）、`python scripts/make_icons.py`（4 图标）。GIF 透明底做法：统一 ADAPTIVE 调色板 + `paste(255, alpha<128)` + `transparency=255, disposal=2`。

## QA 检查单

1. python-pptx 重新打开无异常；内容页齿轮 rotation 逐页 −90°；每页有 `p:bg` 和 Morph 的 `mc:AlternateContent`。
2. 用真实 PowerPoint 渲染逐页目检（pywin32 晚期绑定 `win32com.client.Dispatch('PowerPoint.Application')`，`Slides.Item(i).Export(path,'PNG',1600,900)`；.NET/PowerShell COM 在此机器会报 TYPE_E_CANTLOADLIBRARY）。
3. 重点看：当前章标签在右缘横排不裁字、图标在标签左侧不重叠、顶部/底部标签完整、页脚不压齿轮。
4. 提醒用户：Morph 需 PowerPoint 2019/365 放映；WPS/旧版回退淡入淡出（文件不会坏）。
