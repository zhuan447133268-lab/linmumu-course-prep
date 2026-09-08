# -*- coding: utf-8 -*-
"""可复用的 PPTX 设计系统（linmumu-course-prep 用）。
按专业换肤：import deck_lib as D; D.set_theme('文科'); D.set_course('课程名')
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- palette（按专业换肤） ----------
PALETTES = {
    '工科': dict(primary=0xFF6B35, primary_soft=0xFFECE3, secondary=0x00A08A, secondary_soft=0xDDF3EF, dark=0x0B1F3A, dark_soft=0x163052, bg=0xF7F4EE),
    '理科': dict(primary=0x2A9D8F, primary_soft=0xDDF3EF, secondary=0x457B9D, secondary_soft=0xE1EBF2, dark=0x1D3557, dark_soft=0x2A4A6B, bg=0xF7F9FB),
    '文科': dict(primary=0xE76F51, primary_soft=0xFBEBE4, secondary=0xC1683A, secondary_soft=0xF3E4D8, dark=0x6E2F1C, dark_soft=0x8A4330, bg=0xFBF7F0),
    '经管': dict(primary=0xB8893B, primary_soft=0xF3EAD6, secondary=0x2F4858, secondary_soft=0xE2E7EA, dark=0x1E2D36, dark_soft=0x334A55, bg=0xFAF6F0),
    '医学': dict(primary=0x2C6E8F, primary_soft=0xE2EEF3, secondary=0x4C9A82, secondary_soft=0xE3F1EC, dark=0x173A4A, dark_soft=0x28566A, bg=0xF4F8F6),
    '艺术': dict(primary=0x6D597A, primary_soft=0xECE6EF, secondary=0xB56576, secondary_soft=0xF2E4E8, dark=0x3A2E42, dark_soft=0x554459, bg=0xFBF6F2),
    '农林': dict(primary=0x5A7D3C, primary_soft=0xEBF0E2, secondary=0x8A5A2B, secondary_soft=0xF0E7DA, dark=0x2C3A1E, dark_soft=0x44512F, bg=0xF7F6EE),
}

FONT = 'Microsoft YaHei'
THEME = '工科'
COURSE = '课件'
TOTAL = 15
_n = [0]


def _rgb(h):
    return RGBColor((h >> 16) & 0xFF, (h >> 8) & 0xFF, h & 0xFF)


def set_theme(name):
    global THEME
    THEME = name if name in PALETTES else '工科'


def set_course(label):
    global COURSE
    COURSE = label


def set_total(n):
    global TOTAL
    TOTAL = n


# 颜色（每次 set_theme 后调用 refresh 重算）
PAL = PALETTES[THEME]
INK = _rgb(PAL['dark']); INK_SOFT = _rgb(PAL['dark_soft'])
ORANGE = _rgb(PAL['primary']); ORANGE_SOFT = _rgb(PAL['primary_soft'])
TEAL = _rgb(PAL['secondary']); TEAL_SOFT = _rgb(PAL['secondary_soft'])
BLUE = _rgb(PAL['secondary'])
CREAM = _rgb(PAL['bg']); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xE5, 0xDF, 0xD2); GRAY = RGBColor(0x6B, 0x72, 0x80)
FAINT = RGBColor(0xB9, 0xA8, 0x95); TRACK = RGBColor(0xE9, 0xE4, 0xD9)


def refresh():
    global INK, INK_SOFT, ORANGE, ORANGE_SOFT, TEAL, TEAL_SOFT, BLUE, CREAM
    PAL = PALETTES[THEME]
    INK = _rgb(PAL['dark']); INK_SOFT = _rgb(PAL['dark_soft'])
    ORANGE = _rgb(PAL['primary']); ORANGE_SOFT = _rgb(PAL['primary_soft'])
    TEAL = _rgb(PAL['secondary']); TEAL_SOFT = _rgb(PAL['secondary_soft'])
    BLUE = _rgb(PAL['secondary']); CREAM = _rgb(PAL['bg'])


prs = None
W, H = 13.333, 7.5


def new_prs(theme='工科', course='课件', total=15):
    global prs, W, H
    set_theme(theme); set_course(course); set_total(total); refresh()
    prs = Presentation()
    prs.slide_width = Inches(W); prs.slide_height = Inches(H)
    return prs


def _set_run(r, s, sz, b, c):
    r.text = s
    r.font.size = Pt(sz); r.font.bold = b; r.font.color.rgb = c
    r.font.name = FONT
    rPr = r._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set('typeface', FONT)


def rect(s, l, t, w, h, fill, line=None, lw=1.0,
         shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = s.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if radius is not None:
        try:
            sh.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    return sh


def oval(s, l, t, w, h, fill=None, line=None, lw=1.0):
    return rect(s, l, t, w, h, fill, line, lw, shape=MSO_SHAPE.OVAL)


def text(s, l, t, w, h, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT,
         spacing=1.0, space_after=6):
    tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, item in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = spacing; p.space_after = Pt(space_after)
        if isinstance(item, tuple):
            runs = [item]
        elif isinstance(item, list) and item and not isinstance(item[0], tuple):
            runs = [tuple(item)]          # 扁平 [str,sz,b,c] 当成单 run
        else:
            runs = item
        for (s_, sz, b, c) in runs:
            _set_run(p.add_run(), s_, sz, b, c)
    return tb


def pic(s, path, l, t, w=None, h=None):
    return s.shapes.add_picture(path, Inches(l), Inches(t),
                                Inches(w) if w else None, Inches(h) if h else None)


def card(s, l, t, w, h, fill=WHITE, line=LINE, lw=1.2, radius=0.055):
    return rect(s, l, t, w, h, fill, line, lw,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)


def footer(s):
    n = _n[0]
    text(s, 0.9, 7.08, 8.5, 0.3, [(COURSE, 10.5, False, GRAY)])
    text(s, 11.4, 7.08, 1.03, 0.3, [(f'{n} / {TOTAL}', 11, True, GRAY)],
         align=PP_ALIGN.RIGHT)
    rect(s, 0, 7.42, W, 0.08, TRACK)
    rect(s, 0, 7.42, W * n / TOTAL, 0.08, ORANGE)


def footer_dark(s):
    n = _n[0]
    text(s, 0.9, 7.08, 8.5, 0.3,
         [(COURSE, 10.5, False, RGBColor(0x9A, 0x8C, 0x7C))])
    text(s, 11.4, 7.08, 1.03, 0.3,
         [(f'{n} / {TOTAL}', 11, True, RGBColor(0x9A, 0x8C, 0x7C))],
         align=PP_ALIGN.RIGHT)
    rect(s, 0, 7.42, W, 0.08, INK_SOFT)
    rect(s, 0, 7.42, W * n / TOTAL, 0.08, ORANGE)


def header(s, kicker, title, title_size=30):
    rect(s, 0, 0, W, 0.075, ORANGE)
    text(s, 0.9, 0.42, 11.5, 0.32, [(kicker, 13, True, ORANGE)])
    text(s, 0.9, 0.76, 11.9, 0.75, [(title, title_size, True, INK)])


def new_page(kicker='', title='', title_size=30):
    global prs
    _n[0] += 1
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, CREAM)
    if title:
        header(s, kicker, title, title_size)
    footer(s)
    return s


def section_page(num, title, sub):
    global prs
    _n[0] += 1
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, INK)
    oval(s, 9.2, -2.6, 6.4, 6.4, None, INK_SOFT, 2.2)
    oval(s, 10.2, -1.6, 4.4, 4.4, None, INK_SOFT, 2.2)
    oval(s, 11.0, -0.8, 2.8, 2.8, None, ORANGE, 1.6)
    text(s, 0.9, 1.75, 4.4, 2.2, [(num, 120, True, ORANGE)])
    rect(s, 4.75, 2.25, 0.045, 2.0, ORANGE)
    text(s, 5.25, 2.45, 7.2, 1.9, [(title, 36, True, WHITE), (sub, 16, False, FAINT)],
         spacing=1.15, space_after=10)
    footer_dark(s)
    return s


def conclusion_bar(s, txt_runs, y=6.28, h=0.62):
    rect(s, 0.9, y, W - 1.8, h, INK, None,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    text(s, 1.3, y, W - 2.6, h, [txt_runs], anchor=MSO_ANCHOR.MIDDLE,
         align=PP_ALIGN.CENTER)


def tag(s, l, t, txt, fill=ORANGE_SOFT, fg=ORANGE, w=None):
    w = w or (0.32 + len(txt) * 0.115)
    rect(s, l, t, w, 0.34, fill, None, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
         radius=0.5)
    text(s, l, t - 0.012, w, 0.36, [(txt, 12.5, True, fg)],
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    return w


def demo_link(s, l, t, target='05_交互演示.html', label='▶ 看演示'):
    """视觉提示 + 文件名脚注（不挂可点击超链接，避免预览器报错）。"""
    w = 0.35 + len(label) * 0.19
    card(s, l, t, w, 0.42, ORANGE_SOFT, None, radius=0.5)
    text(s, l, t - 0.012, w, 0.44, [(label, 12.5, True, ORANGE)],
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    text(s, l + w + 0.22, t - 0.012, 7.0, 0.44,
         [('配套演示：' + target + '（与本课件同文件夹，双击即开）', 12, False, GRAY)],
         anchor=MSO_ANCHOR.MIDDLE)
    return w
