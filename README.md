# ppt-skills

两个自包含的 PowerPoint 效果技能（Agent Skill 格式：`SKILL.md` + `scripts/`），用 python-pptx 程序化生成带 Morph 平滑切换的特效页。

## 包含的技能

| 技能 | 效果 | 依赖 |
| --- | --- | --- |
| `ppt-zhuanpan-skill` | 机械风大齿轮转盘：左侧大齿轮只露右半边，每翻页逆时针转 90°，当前章横排标签落在右缘 3 点钟方向，工具图标接力伸缩，暗灰钢色 + 常驻转动齿轮 GIF | python-pptx |
| `ppt-loukong-skill` | 照片填充镂空字：3 页开场（上下遮罩左右错位 → 居中照片透字 → 上下退场揭晓）与 2 页电影式谢幕（THANKS → 上下滑出 → 底图放大 + 结束语） | python-pptx + Pillow |

## 安装

把任一技能文件夹 junction / 软链 / 复制到你的 agent 技能目录（任何读取 `SKILL.md` 的 agent 均可，如 Qoder、WorkBuddy、DeepSeek Harness 的 skills 目录）：

```bash
# Windows junction 示例
mklink /J "%USERPROFILE%\.qoder-cn\skills\ppt-loukong-skill" "<clone>\ppt-loukong-skill"
```

多技能克隆建议放在统一仓库目录下，再按技能名分别链接。

## 快速使用

```bash
# 转盘：内置演示或自定义配置
python ppt-zhuanpan-skill/scripts/zhuanpan_deck.py -o out.pptx
python ppt-zhuanpan-skill/scripts/zhuanpan_deck.py -o out.pptx --config deck.json

# 镂空字：开场 / 谢幕
python ppt-loukong-skill/scripts/loukong_deck.py -o opening.pptx --config ppt-loukong-skill/examples/opening.json
python ppt-loukong-skill/scripts/loukong_deck.py -o closing.pptx --config ppt-loukong-skill/examples/closing.json

# 镂空字静态核验（页数/透明字孔/对象身份/Morph 时长）
python ppt-loukong-skill/scripts/verify_deck.py opening.pptx --mode opening
```

配置字段见各技能 `SKILL.md`。

## 注意

- Morph（平滑）切换需要 PowerPoint 2019 / 365 放映；WPS 或旧版会回退为普通切换，文件不会损坏。
- `ppt-loukong-skill` 的标题是透明 PNG 字孔，生成后不能在 PowerPoint 里当文字编辑；改标题/字体请改配置重跑。
- 仓库内 `assets/demo-bg.png` 为程序生成的抽象山景演示底图（`scripts/make_demo_asset.py` 可重建），非照片、无版权风险；正式使用请替换为自己的底图。

## 许可

MIT，见 [LICENSE](LICENSE)。
