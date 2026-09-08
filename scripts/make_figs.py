# -*- coding: utf-8 -*-
"""生成课件配图（透明底 PNG）。这是工科实例模板：换主题=复制本文件，改绘图内容。
用法：python make_figs.py [学科族] [输出目录]   （输出目录默认 ./_figs，相对当前目录）
"""
import os
import sys
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['mathtext.fontset'] = 'dejavusans'

# ---------- 配色（canonical 来源 = deck_lib.PALETTES，改色只改那一处） ----------
_FIG_THEME = sys.argv[1] if len(sys.argv) > 1 else '工科'
OUT = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(os.getcwd(), '_figs')
os.makedirs(OUT, exist_ok=True)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import deck_lib as _DL
_P = _DL.PALETTES.get(_FIG_THEME, _DL.PALETTES['工科'])
_hx = lambda v: '#{0:06X}'.format(v)

INK = _hx(_P['dark'])
SECOND = _hx(_P['secondary'])   # 第二强调：θ₂ / 第二连杆 / 工作空间环
PRIMARY = _hx(_P['primary'])    # 主色：θ₁ / 第一连杆 / 基准线以外的强调
GRAY = '#8A93A0'
LIGHT = '#C9D2DC'

L1, L2 = 400.0, 300.0


def setup(ax, xlim, ylim, aspect=True):
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    if aspect:
        ax.set_aspect('equal')
    ax.axis('off')


def axes_arrows(ax, xmax=780, ymax=760):
    ax.annotate('', xy=(xmax, 0), xytext=(-90, 0),
                arrowprops=dict(arrowstyle='->', color=GRAY, lw=1.4))
    ax.annotate('', xy=(0, ymax), xytext=(0, -90),
                arrowprops=dict(arrowstyle='->', color=GRAY, lw=1.4))
    ax.text(xmax + 18, -18, 'x', color=GRAY, fontsize=15)
    ax.text(-26, ymax + 4, 'y', color=GRAY, fontsize=15)
    ax.plot([-70, xmax], [0, 0], color=LIGHT, lw=0.9, zorder=0)
    ax.plot([0, 0], [-70, ymax], color=LIGHT, lw=0.9, zorder=0)


def base(ax, s=1.0):
    """基座：三角 + 地面斜线。"""
    w, h = 70 * s, 34 * s
    ax.add_patch(plt.Polygon([(-w, 0), (w, 0), (0, h)], color=INK, zorder=5))
    for x in np.linspace(-w, w, 9):
        ax.plot([x, x - 13 * s], [0, -15 * s], color=GRAY, lw=1.0, zorder=4)


def joint(ax, p, c=INK, r=9, z=6):
    ax.add_patch(plt.Circle(p, r, color=c, zorder=z))
    ax.add_patch(plt.Circle(p, r, color='white', zorder=z + 1,
                            fill=False, lw=2.2, ec=c))


def link(ax, a, b, c, lw=9):
    ax.plot([a[0], b[0]], [a[1], b[1]], color=c, lw=lw,
            solid_capstyle='round', zorder=4)


def arc(ax, center, r, a0, a1, c, lw=2.0):
    t = np.linspace(np.radians(a0), np.radians(a1), 80)
    ax.plot(center[0] + r * np.cos(t), center[1] + r * np.sin(t),
            color=c, lw=lw, zorder=5)


def ang_text(ax, center, r, deg, txt, c):
    a = np.radians(deg)
    ax.text(center[0] + r * np.cos(a), center[1] + r * np.sin(a), txt,
            color=c, fontsize=16, ha='center', va='center',
            bbox=dict(boxstyle='round,pad=0.22', fc='white', ec=c, lw=1.3),
            zorder=7)


def mid_text(ax, a, b, txt, c, off=(0, 0), leader=False):
    m = ((a[0] + b[0]) / 2 + off[0], (a[1] + b[1]) / 2 + off[1])
    if leader:
        ax.plot([(a[0] + b[0]) / 2, m[0]], [(a[1] + b[1]) / 2, m[1]],
                color=c, lw=1.0, ls=(0, (2, 2)), zorder=5)
    ax.text(m[0], m[1], txt, color=c, fontsize=17, ha='center', va='center',
            bbox=dict(boxstyle='round,pad=0.25', fc='white', ec='none',
                      alpha=0.92), zorder=7)


# ---------- 图 1：构型与坐标系 ----------
def fig_arm():
    t1, t2 = 50.0, 25.0
    O = (0.0, 0.0)
    P1 = (L1 * np.cos(np.radians(t1)), L1 * np.sin(np.radians(t1)))
    P = (P1[0] + L2 * np.cos(np.radians(t1 + t2)),
         P1[1] + L2 * np.sin(np.radians(t1 + t2)))

    fig, ax = plt.subplots(figsize=(6.4, 5.8), dpi=200)
    setup(ax, (-160, 830), (-140, 840))
    axes_arrows(ax)
    base(ax)

    arc(ax, O, 130, 0, t1, PRIMARY)
    ang_text(ax, O, 175, t1 / 2 + 4, r'$\theta_1$', PRIMARY)
    arc(ax, P1, 92, t1, t1 + t2, SECOND)
    ang_text(ax, P1, 145, t1 + t2 / 2, r'$\theta_2$', SECOND)

    link(ax, O, P1, INK, 10)
    link(ax, P1, P, INK, 10)
    joint(ax, O)
    joint(ax, P1, c=SECOND)
    ax.add_patch(plt.Circle(P, 10, color=SECOND, zorder=6))

    mid_text(ax, O, P1, r'$L_1=400$', INK, off=(-118, -6))
    mid_text(ax, P1, P, r'$L_2=300$', INK, off=(150, 65), leader=True)
    ax.text(P[0] + 26, P[1] + 16, '末端 (x, y)', color=SECOND,
            fontsize=15, ha='left')

    ax.plot([0, 200], [0, 0], color=PRIMARY, lw=1.6, ls=(0, (5, 3)))
    ax.text(230, -46, '水平基准', color=PRIMARY, fontsize=13, ha='center')

    fig.savefig(os.path.join(OUT, 'fig_arm.png'), transparent=True,
                bbox_inches='tight', pad_inches=0.12)
    plt.close(fig)


# ---------- 图 2：θ₂ 基准对比 ----------
def fig_theta2():
    t1, t2 = 50.0, 25.0
    O = (0.0, 0.0)
    P1 = (L1 * np.cos(np.radians(t1)), L1 * np.sin(np.radians(t1)))
    P_ok = (P1[0] + L2 * np.cos(np.radians(t1 + t2)),
            P1[1] + L2 * np.sin(np.radians(t1 + t2)))
    P_bad = (P1[0] + L2 * np.cos(np.radians(t2)),
             P1[1] + L2 * np.sin(np.radians(t2)))

    fig, ax = plt.subplots(figsize=(6.6, 5.8), dpi=200)
    setup(ax, (-150, 760), (-140, 800))
    axes_arrows(ax, 700, 700)
    base(ax)

    # 错误
    link(ax, P1, P_bad, SECOND, 8)
    joint(ax, P_bad, c=SECOND, r=8)
    arc(ax, P1, 130, 0, t2, SECOND, 1.6)
    ang_text(ax, P1, 186, t2 / 2 + 4, r'$25^\circ$', SECOND)

    # 正确
    link(ax, O, P1, INK, 10)
    link(ax, P1, P_ok, SECOND, 9)
    joint(ax, O)
    joint(ax, P1, c=INK)
    ax.add_patch(plt.Circle(P_ok, 10, color=SECOND, zorder=6))
    arc(ax, P1, 92, t1, t1 + t2, SECOND, 1.8)

    ax.plot([P_ok[0], P_bad[0]], [P_ok[1], P_bad[1]],
            color=SECOND, lw=2.2, ls=(0, (6, 4)), zorder=5)
    ax.text((P_ok[0] + P_bad[0]) / 2 + 18, (P_ok[1] + P_bad[1]) / 2,
            '差了 253.6 mm', color=SECOND, fontsize=15, ha='left')

    ax.text(P_ok[0] + 20, P_ok[1] + 14, '正确', color=SECOND, fontsize=15)
    ax.text(P_bad[0] + 20, P_bad[1] - 6, '错误', color=SECOND, fontsize=15)

    fig.savefig(os.path.join(OUT, 'fig_theta2.png'), transparent=True,
                bbox_inches='tight', pad_inches=0.12)
    plt.close(fig)


# ---------- 图 3：工作空间圆环 ----------
def fig_workspace(points=None, labels=None, name='fig_workspace.png'):
    fig, ax = plt.subplots(figsize=(6.8, 6.0), dpi=200)
    setup(ax, (-820, 820), (-760, 860))
    axes_arrows(ax, 800, 820)

    th = np.linspace(0, 2 * np.pi, 400)
    ax.fill((L1 + L2) * np.cos(th), (L1 + L2) * np.sin(th),
            color=SECOND, alpha=0.13, zorder=1)
    ax.fill(100 * np.cos(th), 100 * np.sin(th),
            color='white', zorder=2)
    ax.plot((L1 + L2) * np.cos(th), (L1 + L2) * np.sin(th),
            color=SECOND, lw=2.6, zorder=3)
    ax.plot(100 * np.cos(th), 100 * np.sin(th),
            color=SECOND, lw=2.6, zorder=3)

    ax.annotate('', xy=((L1 + L2) * np.cos(np.radians(25)),
                        (L1 + L2) * np.sin(np.radians(25))), xytext=(0, 0),
                arrowprops=dict(arrowstyle='->', color=SECOND, lw=1.8))
    ax.text(695, 350, '700', color=SECOND, fontsize=18, ha='center')
    ax.annotate('', xy=(100 * np.cos(np.radians(-30)),
                        100 * np.sin(np.radians(-30))), xytext=(0, 0),
                arrowprops=dict(arrowstyle='->', color=SECOND, lw=1.8))
    ax.text(160, -100, '100', color=SECOND, fontsize=18, ha='center')

    base(ax, 0.85)

    if points:
        for (px, py), lb, (dx, dy), ha in zip(
                points, labels, [(-14, 62), (34, 30)], ['right', 'left']):
            ax.add_patch(plt.Circle((px, py), 20, color=INK, zorder=6))
            ax.add_patch(plt.Circle((px, py), 20, color='none',
                                    ec='white', lw=2.5, zorder=7))
            ax.text(px + dx, py + dy, lb, color=INK, fontsize=16, ha=ha,
                    bbox=dict(boxstyle='round,pad=0.24', fc='white',
                              ec=INK, lw=1.2), zorder=8)

    fig.savefig(os.path.join(OUT, name), transparent=True,
                bbox_inches='tight', pad_inches=0.12)
    plt.close(fig)


# ---------- 图 4：两段累加 ----------
def fig_accum():
    t1, t2 = 50.0, 25.0
    O = (0.0, 0.0)
    P1 = (L1 * np.cos(np.radians(t1)), L1 * np.sin(np.radians(t1)))
    P = (P1[0] + L2 * np.cos(np.radians(t1 + t2)),
         P1[1] + L2 * np.sin(np.radians(t1 + t2)))

    fig, ax = plt.subplots(figsize=(6.6, 5.8), dpi=200)
    setup(ax, (-150, 800), (-140, 820))
    axes_arrows(ax, 740, 740)
    base(ax)

    # 第一段投影
    ax.plot([O[0], P1[0]], [O[1], O[1]], color=PRIMARY, lw=0, zorder=1)
    ax.annotate('', xy=(P1[0], 0), xytext=(0, 0),
                arrowprops=dict(arrowstyle='->', color=PRIMARY, lw=2.0))
    ax.annotate('', xy=(P1[0], P1[1]), xytext=(P1[0], 0),
                arrowprops=dict(arrowstyle='->', color=PRIMARY, lw=2.0))
    ax.plot([P1[0], P1[0]], [0, P1[1]], color=PRIMARY, lw=1.2, ls=(0, (4, 3)))
    ax.text(P1[0] / 2, -46, r'$L_1\cos\theta_1$', color=PRIMARY, fontsize=16,
            ha='center', bbox=dict(boxstyle='round,pad=0.2', fc='white',
                                   ec='none', alpha=0.95))
    ax.text(P1[0] + 26, P1[1] / 2, r'$L_1\sin\theta_1$', color=PRIMARY,
            fontsize=16, ha='left', bbox=dict(boxstyle='round,pad=0.2',
                                              fc='white', ec='none',
                                              alpha=0.95))

    # 第二段投影
    ax.annotate('', xy=(P[0], P1[1]), xytext=(P1[0], P1[1]),
                arrowprops=dict(arrowstyle='->', color=SECOND, lw=2.0))
    ax.annotate('', xy=(P[0], P[1]), xytext=(P[0], P1[1]),
                arrowprops=dict(arrowstyle='->', color=SECOND, lw=2.0))
    ax.plot([P[0], P[0]], [P1[1], P[1]], color=SECOND, lw=1.2, ls=(0, (4, 3)))
    ax.text((P1[0] + P[0]) / 2 + 8, P1[1] - 108,
            r'$L_2\cos(\theta_1+\theta_2)$', color=SECOND, fontsize=16,
            ha='center', bbox=dict(boxstyle='round,pad=0.2', fc='white',
                                   ec='none', alpha=0.95))
    ax.text(P[0] + 30, (P1[1] + P[1]) / 2,
            r'$L_2\sin(\theta_1+\theta_2)$', color=SECOND, fontsize=16,
            ha='left', bbox=dict(boxstyle='round,pad=0.2', fc='white',
                                 ec='none', alpha=0.95))

    link(ax, O, P1, INK, 10)
    link(ax, P1, P, INK, 10)
    joint(ax, O)
    joint(ax, P1, c=SECOND)
    ax.add_patch(plt.Circle(P, 10, color=SECOND, zorder=6))

    fig.savefig(os.path.join(OUT, 'fig_accum.png'), transparent=True,
                bbox_inches='tight', pad_inches=0.12)
    plt.close(fig)


if __name__ == '__main__':
    fig_arm()
    fig_theta2()
    fig_workspace()                      # 空白圆环
    fig_workspace(points=[(500, 200), (-100, 550)],
                  labels=['A (500, 200)', 'B (−100, 550)'],
                  name='fig_workspace_pts.png')
    fig_accum()
    print('figs ->', OUT)
    for f in sorted(os.listdir(OUT)):
        print(' ', f, os.path.getsize(os.path.join(OUT, f)))
