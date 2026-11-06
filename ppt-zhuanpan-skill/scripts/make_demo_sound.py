# -*- coding: utf-8 -*-
"""合成 examples 用的机械"咔哒"换页音效（纯 numpy 生成，无外部素材、无版权）。

技能本身不附带任何音频素材；这个脚本只是为了让 examples 里带 sound 的配置开箱能跑。
正式使用把自己准备好的 wav/mp3/m4a 路径写进配置的 sound 即可。

用法：
  python scripts/make_demo_sound.py                  # 输出到 examples/sounds/
  python scripts/make_demo_sound.py -o my/sounds     # 输出到指定目录
"""
import argparse
import os
import wave

import numpy as np

SR = 44100
PEAK = 0.62                     # 归一化峰值：留足余量，切换音效不会爆


def damped(freq, decay, n, phase=0.0):
    """一个阻尼谐振：freq 主频，decay 时间常数（秒），n 样本数。"""
    t = np.arange(n) / SR
    return np.exp(-t / decay) * np.sin(2 * np.pi * freq * t + phase)


def noise_burst(n, hp_hz=1500.0):
    """短促的高频噪声（咔哒的"脆"来自这里），一阶高通。"""
    x = np.random.RandomState(7).randn(n)
    a = np.exp(-2 * np.pi * hp_hz / SR)
    y = np.empty_like(x)
    y[0] = x[0]
    for i in range(1, n):
        y[i] = a * y[i - 1] + (1 - a) * (x[i] - x[i - 1])
    return y


def click(at, gain=1.0, tail=True):
    """一次咔哒：噪声瞬态 + 两个金属谐振 + 一下低频"咚" + 牙尖回弹。"""
    n = int(0.12 * SR)
    out = np.zeros(n)
    nb = noise_burst(int(0.004 * SR))
    out[:nb.size] += nb * 0.9
    out += damped(2600, 0.010, n) * 0.55
    out += damped(5300, 0.006, n) * 0.30
    out += damped(150, 0.035, n) * 0.35
    if tail:                                    # 齿轮牙尖回弹的第二下
        m = int(0.045 * SR)
        out[m:] += damped(3100, 0.007, n - m) * 0.18
    return at, out * gain


def render(events, total, gap=0):
    """按事件时刻叠加，头尾各做 3ms 淡入淡出，再归一化到 PEAK。"""
    y = np.zeros(int(total * SR) + gap)
    for at, seg in events:
        i = int(at * SR) + gap
        y[i:i + seg.size] += seg
    y = y[gap:]
    ramp = int(0.003 * SR)
    y[:ramp] *= np.linspace(0, 1, ramp)
    y[-ramp:] *= np.linspace(1, 0, ramp)
    peak = np.abs(y).max()
    return y / peak * PEAK if peak > 0 else y


def save(path, y):
    with wave.open(path, 'wb') as f:
        f.setnchannels(1)
        f.setsampwidth(2)
        f.setframerate(SR)
        f.writeframes((np.clip(y, -1, 1) * 32767).astype('<i2').tobytes())


def build():
    """返回 {文件名: 波形}。click = 一下落位；ratchet = 一串由密到疏的咬合。"""
    out = {}
    out['click.wav'] = render([click(0.0, 1.0)], 0.22, gap=int(0.01 * SR))

    events, at, step, k = [], 0.0, 0.055, 0
    while at < 0.62:                            # 步进越来越宽：像转盘减速落位
        events.append(click(at, gain=0.85 - 0.055 * k + 0.06 * (k % 2)))
        step *= 1.13
        at += step
        k += 1
    out['ratchet.wav'] = render(events, 0.95, gap=int(0.01 * SR))
    return out


def main(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    for name, y in sorted(build().items()):
        path = os.path.join(out_dir, name)
        save(path, y)
        print('%-12s %.3fs  peak=%.2f  rms=%.3f  %d bytes  -> %s'
              % (name, y.size / float(SR), np.abs(y).max(),
                 float(np.sqrt((y ** 2).mean())), os.path.getsize(path), path))


if __name__ == '__main__':
    ap = argparse.ArgumentParser(description='合成机械咔哒换页音效（纯 numpy）')
    ap.add_argument('-o', '--out', default=None,
                    help='输出目录（默认 ../examples/sounds）')
    a = ap.parse_args()
    here = os.path.dirname(os.path.abspath(__file__))
    main(a.out or os.path.join(os.path.dirname(here), 'examples', 'sounds'))
