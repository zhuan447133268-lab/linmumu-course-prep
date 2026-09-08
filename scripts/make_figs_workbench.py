# -*- coding: utf-8 -*-
"""工作台课程配图：资料窗 只读 vs 可写 对比图（对应学生最常犯的错）。
工具/流程类课程实例（默认文科配色）。换主题=复制本文件，改绘图内容。
配色 canonical 来源 = deck_lib.PALETTES（改色只改那一处）。
用法：python make_figs_workbench.py [学科族] [输出目录]   （默认 文科 ./_figs）
"""
import os
import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle
import matplotlib.font_manager as fm

_FIG_THEME = sys.argv[1] if len(sys.argv) > 1 else '文科'
FIGS = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(os.getcwd(), '_figs')
os.makedirs(FIGS, exist_ok=True)

# 配色：从 deck_lib（canonical）取当前学科族，转 hex 字符串
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import deck_lib as _DL
_P = _DL.PALETTES.get(_FIG_THEME, _DL.PALETTES['文科'])
_hx = lambda v: '#{0:06X}'.format(v)
PRIMARY = _hx(_P['primary'])
SECOND = _hx(_P['secondary'])
DARK = _hx(_P['dark'])
GREEN = '#3C8C5A'   # 图内专用语义色（正确/安全），不随学科换
RED = '#C0392B'     # 图内专用语义色（错误/危险），不随学科换
INK = '#3A2A20'

# CJK 字体
for cand in ['Microsoft YaHei', 'SimHei', 'PingFang SC', 'Noto Sans CJK SC']:
    try:
        fm.findfont(cand, fallback_to_default=False)
        plt.rcParams['font.sans-serif'] = [cand]
        break
    except Exception:
        continue
plt.rcParams['axes.unicode_minus'] = False


def window(ax, x, y, w, h, line_c, title, lines, accent):
    box = FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.02,rounding_size=0.04',
                         linewidth=2, edgecolor=line_c, facecolor='white', zorder=2)
    ax.add_patch(box)
    # 标题栏
    bar = FancyBboxPatch((x, y + h - 0.55), w, 0.55,
                         boxstyle='round,pad=0.0,rounding_size=0.04',
                         linewidth=0, edgecolor='none', facecolor=accent, zorder=3)
    ax.add_patch(bar)
    ax.text(x + w / 2, y + h - 0.28, title, color='white', fontsize=15,
            ha='center', va='center', fontweight='bold', zorder=4)
    # 内容行
    for i, ln in enumerate(lines):
        yy = y + h - 1.05 - i * 0.62
        ax.text(x + 0.45, yy, '• ' + ln, color=INK, fontsize=13.5,
                ha='left', va='center', zorder=4)


def badge(ax, cx, cy, color, symbol):
    c = Circle((cx, cy), 0.42, color=color, zorder=5)
    ax.add_patch(c)
    ax.text(cx, cy, symbol, color='white', fontsize=26, ha='center',
            va='center', fontweight='bold', zorder=6)


def fig_readonly():
    fig, ax = plt.subplots(figsize=(11.8, 4.0), dpi=160)
    ax.set_xlim(0, 12); ax.set_ylim(0, 4); ax.axis('off')

    # 左：只读（正确）
    window(ax, 0.4, 0.5, 5.3, 3.0, GREEN,
           '资料窗 · 只读模式（正确）',
           ['只能看你的文献/笔记', '不能改、不能删原文件', 'AI 整理时不会碰源数据'], GREEN)
    badge(ax, 5.55, 3.45, GREEN, '对')

    # 右：可写（危险）
    window(ax, 6.3, 0.5, 5.3, 3.0, RED,
           '资料窗 · 可写模式（危险）',
           ['AI 能改、能删原文件', '一句话就可能清掉你的资料', '误删 = 不可逆'], RED)
    badge(ax, 11.45, 3.45, RED, '危')

    ax.annotate('', xy=(6.2, 2.0), xytext=(5.8, 2.0),
                arrowprops=dict(arrowstyle='->', color=RED, lw=2.2))
    ax.text(6.05, 2.45, '别让\nAI 拿到\n写权限', color=RED, fontsize=12.5,
            ha='center', va='center', fontweight='bold')

    ax.text(6.0, 0.18, '一句话守住安全：把资料窗说成「只能看、不能改我的原文件」',
            color=DARK, fontsize=13.5, ha='center', va='center', fontweight='bold')
    fig.tight_layout(pad=0.2)
    fig.savefig(os.path.join(FIGS, 'fig_readonly.png'), transparent=True,
                bbox_inches='tight', dpi=160)
    plt.close(fig)
    print('fig_readonly.png done')


if __name__ == '__main__':
    fig_readonly()
