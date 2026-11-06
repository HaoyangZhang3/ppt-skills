# ppt-skills

两个自包含的 PowerPoint 效果技能（Agent Skill 格式：`SKILL.md` + `scripts/`），用 python-pptx 程序化生成带 Morph 平滑切换的特效页。

## 包含的技能

| 技能 | 效果 | 依赖 |
| --- | --- | --- |
| `ppt-zhuanpan-skill` | 翻页旋转大转盘：左侧转盘只露右半边，每翻页整盘转 360/N（N=章数），当前章标签落在右缘 3 点钟方向，图标接力伸缩；三种造型——中性圆盘 / 9 齿齿轮 / 内外双盘（外环逆时针、内盘顺时针反向转，照片填进两个盘，标签沿圆周切向）；章数、配色、背景、图标、换页音效全可配，默认中性、不附带素材 | python-pptx |
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
python ppt-zhuanpan-skill/scripts/zhuanpan_deck.py -o out.pptx --config ppt-zhuanpan-skill/examples/double4.json   # 内外双盘 + 照片填充
python ppt-zhuanpan-skill/scripts/zhuanpan_deck.py -o out.pptx --config ppt-zhuanpan-skill/examples/gear9_5.json   # 9 齿齿轮
# 换页音效（PowerPoint 切换声音）：配置里写 "sound": "sounds/click.wav"，先合成演示音效
python ppt-zhuanpan-skill/scripts/make_demo_sound.py

# 镂空字：开场 / 谢幕
python ppt-loukong-skill/scripts/loukong_deck.py -o opening.pptx --config ppt-loukong-skill/examples/opening.json
python ppt-loukong-skill/scripts/loukong_deck.py -o closing.pptx --config ppt-loukong-skill/examples/closing.json

# 镂空字静态核验（页数/透明字孔/对象身份/Morph 时长）
python ppt-loukong-skill/scripts/verify_deck.py opening.pptx --mode opening
```

配置字段见各技能 `SKILL.md`。

## 注意

- Morph（平滑）切换需要 PowerPoint 2019 / 365 放映；WPS 或旧版会回退为普通切换，文件不会损坏。
- `ppt-zhuanpan-skill` 不附带任何图片/音频素材：照片填充（`wheel: "double"` 必需）、图标接力、背景图与换页音效（`sound`）都要自己在配置里给路径；默认是白底中性圆盘、静音。
- `ppt-zhuanpan-skill/examples/photos/demo-*.png` 是为让 examples 开箱能跑而程序生成的抽象演示图（`scripts/make_demo_photos.py` 可重建），非真实照片、无版权风险；正式使用请替换为自己的照片。
- `ppt-zhuanpan-skill/examples/sounds/*.wav` 同样由 `scripts/make_demo_sound.py` 用 numpy 现合成（阻尼谐振 + 高频噪声瞬态），非录音、无版权风险。换页音效是 PowerPoint 的「切换声音」，只在放映时响。
- `ppt-loukong-skill` 的标题是透明 PNG 字孔，生成后不能在 PowerPoint 里当文字编辑；改标题/字体请改配置重跑。
- 仓库内 `assets/demo-bg.png` 为程序生成的抽象山景演示底图（`scripts/make_demo_asset.py` 可重建），非照片、无版权风险；正式使用请替换为自己的底图。

## 维护（改完即推送）

修改任一技能内容后立刻同步，保持本地与远程一致：

```bash
git add -A
git commit -m "简述本次修改"
git push origin main
```

凭据经 `gh auth setup-git` 绑定后可直接 push。若 push 长时间无响应，先确认本机到 github.com 的网络连通（部分网络需加速），可加 `GIT_TERMINAL_PROMPT=0` 让凭据问题快速报错而不是挂起。

## 许可

MIT，见 [LICENSE](LICENSE)。
