#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""交付前三查巡检（linmumu-course-prep skill · 与 export_shots 平级）

用法:
    python check_deliverables.py <课程节目录> [学科族] [--no-write]

查1 数字复算   运行课程节目录内的 _recalc.py（每节生成时必须随讲稿一起写，
               把讲稿/课件里每个计算数字 assert 一遍）。缺文件 → WARN。
               另输出 01/03/04/05 的「数字+单位」台账，供人工扫一眼跨文件不一致。
查2 配色 diff  抽取 05_交互演示.html 与 02_课件.pptx（slide XML）里全部 hex，
               与 deck_lib.PALETTES[学科族] 比对：精确命中或 RGB 距离<=30 的浅调
               视为通过，其余列出待人工判断（语义扩展色如工位色需在 00 说明）。
               附项：05 的 SVG 取景检查（viewBox 必须紧凑，任一方向内容覆盖率
               <50% 即 FAIL——防"沿用公式坐标系导致整页放大"）
查3 引用存在性 提取 00/01/03/04/06 里的「第 X 页 / PX」引用和「N 页」总页数声明，
               与 02 课件实际页数比对。推荐引用一律用页面标题（kicker），天然不烂。

结果默认写进 00_待你确认.md 末尾的「附：交付前三查」区（--no-write 关闭）。
退出码: 0=全过 1=有 FAIL 2=只有 WARN（CI 可按需区分）。
"""
import argparse
import datetime
import os
import re
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

NEUTRALS = {
    'FFFFFF', 'F5F6F8', 'FBFCFD', 'F7F7F8', 'E3E6EA', 'ECEFF2', 'E2E8EF',
    'C9D3DD', 'C9D2DC', '9AA5B1', '8A93A0', '6B7280', '1F2328',
    # deck_lib 中性装饰常量（LINE / TRACK / FAINT / 页脚弱化字）——设计系统自身的一部分
    'E5DFD2', 'E9E4D9', 'B9A895', '9A8C7C', '8A7A6E',  # 05 常用中性暖灰
    '9AA3AD',  # 05 错误态灰虚线专用（对错对比的"灰"，非学科色）
}

FALLBACK_PALETTES = {  # 与 deck_lib.PALETTES 保持同步；import 失败时降级用
    '工科': ('FF6B35', 'FFECE3', '00A08A', 'DDF3EF', '0B1F3A', '163052', 'F7F4EE'),
    '理科': ('2A9D8F', 'DDF3EF', '457B9D', 'E1EBF2', '1D3557', '2A4A6B', 'F7F9FB'),
    '文科': ('E76F51', 'FBEBE4', 'C1683A', 'F3E4D8', '6E2F1C', '8A4330', 'FBF7F0'),
    '经管': ('B8893B', 'F3EAD6', '2F4858', 'E2E7EA', '1E2D36', '334A55', 'FAF6F0'),
    '医学': ('2C6E8F', 'E2EEF3', '4C9A82', 'E3F1EC', '173A4A', '28566A', 'F4F8F6'),
    '艺术': ('6D597A', 'ECE6EF', 'B56576', 'F2E4E8', '3A2E42', '554459', 'FBF6F2'),
    '农林': ('5A7D3C', 'EBF0E2', '8A5A2B', 'F0E7DA', '2C3A1E', '44512F', 'F7F6EE'),
}


def load_palette(theme):
    try:
        from deck_lib import PALETTES
        p = PALETTES[theme]
        return ['{0:06X}'.format(v & 0xFFFFFF) for v in p.values()]
    except Exception:
        return list(FALLBACK_PALETTES.get(theme, FALLBACK_PALETTES['工科']))


def norm_hex(h):
    h = h.upper().lstrip('#')
    if len(h) == 3:
        h = ''.join(c * 2 for c in h)
    return h


def rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def dist(a, b):
    return sum((x - y) ** 2 for x, y in zip(rgb(a), rgb(b))) ** 0.5


def read(p):
    """读文本。.docx 从 zip 抽 word/document.xml 按段落拼回纯文本（不依赖 python-docx，
    巡检自身保持零额外依赖）；其余按 UTF-8 读。"""
    import io
    if p.lower().endswith('.docx'):
        try:
            with zipfile.ZipFile(p) as z:
                xml = z.read('word/document.xml').decode('utf-8', 'ignore')
        except Exception:
            return ''
        paras = re.split(r'</w:p>', xml)
        return '\n'.join(re.sub(r'<[^>]+>', '', x) for x in paras)
    return io.open(p, encoding='utf-8', errors='ignore').read()


def pick(lesson, base):
    """按基名取文件：03/04 自 v1.2 起 docx 直出，md 为旧版兜底。"""
    for ext in ('.docx', '.md'):
        p = os.path.join(lesson, base + ext)
        if os.path.exists(p):
            return p
    return None


# ---------------- 查1 ----------------
def check_numbers(lesson, results):
    recalc = os.path.join(lesson, '_recalc.py')
    if os.path.exists(recalc):
        # 子进程强制 UTF-8：agent 写的 _recalc.py 可能打印 ✔/− 等非 GBK 字符，
        # 中文 Windows 控制台默认 GBK 会因 UnicodeEncodeError 假 FAIL（断言其实全过）
        r = subprocess.run([sys.executable, '-X', 'utf8', recalc],
                           capture_output=True, text=True,
                           encoding='utf-8', errors='replace')
        if r.returncode == 0:
            results.append(('查1 数字复算', 'PASS', '_recalc.py 全部断言通过'))
        else:
            results.append(('查1 数字复算', 'FAIL',
                            '_recalc.py 复算失败:\n' + (r.stdout + r.stderr).strip()[:800]))
    else:
        results.append(('查1 数字复算', 'WARN',
                        '缺 _recalc.py（本节数字未做机器复算）。按铁律 9，随讲稿一起补：'
                        '把本节每个计算数字 assert 一遍，命名 _recalc.py 放本目录。'))
    # 数字台账（人工扫描跨文件不一致，如 253.6 只在某一个文件出现）
    inv = {}
    candidates = [pick(lesson, b) for b in ('01_这一节的设计', '03_讲稿', '04_测验')]
    candidates.append(os.path.join(lesson, '05_交互演示.html'))
    for p in candidates:
        if not p or not os.path.exists(p):
            continue
        tag = os.path.basename(p)[:2]
        for m in re.finditer(r'(\d+(?:\.\d+)?)\s*(mm|度|°|分钟|学时|Hz|kHz|kg|%|页)', read(p)):
            inv.setdefault(m.group(2), {}).setdefault(m.group(1), set()).add(tag)
    lines = []
    for unit in sorted(inv):
        items = sorted(inv[unit].items(), key=lambda kv: float(kv[0]))
        lines.append(unit + ': ' + '、'.join(
            v + '(' + ','.join(sorted(fs)) + ')' for v, fs in items))
    return lines


# ---------------- 查2 ----------------
def check_colors(lesson, theme, results):
    palette = load_palette(theme)
    found = {}

    html = os.path.join(lesson, '05_交互演示.html')
    if os.path.exists(html):
        raw = read(html)
        # 先剥非颜色上下文，防止 id 引用被当 hex 色误判
        # （实测案例：注释里写 g#bad → 被当成 #BBAADD 报"距配色色 74"）
        raw = re.sub(r'["\']#[A-Za-z_][\w-]*["\']', '', raw)    # JS/CSS 字符串选择器 '#id'
        raw = re.sub(r'url\(#[^)]+\)', '', raw)                 # SVG paint-server 引用 url(#id)
        raw = re.sub(r'(?:xlink:)?href="#[^"]+"', '', raw)      # <use href="#id">
        raw = re.sub(r'#[A-Za-z_][\w-]*(?=\s*[{,.:])', '', raw)  # CSS 选择器 #id{ / #id, / #id. / #id:
        # 颜色里的 # 前面绝不会紧跟字母/连字符（g#bad 这类元素前缀引用直接排除）
        hexes = set(re.findall(r'(?<![\w-])#[0-9A-Fa-f]{3,6}\b', raw))
        found['05_交互演示.html'] = {norm_hex(h) for h in hexes}

    pptx = os.path.join(lesson, '02_课件.pptx')
    if os.path.exists(pptx):
        hexes = set()
        with zipfile.ZipFile(pptx) as z:
            for n in z.namelist():
                if re.match(r'ppt/slides/slide\d+\.xml$', n):
                    hexes |= set(re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"',
                                            z.read(n).decode('utf-8', 'ignore')))
        found['02_课件.pptx'] = {norm_hex(h) for h in hexes}

    unknown = []
    for src, hexes in sorted(found.items()):
        for h in sorted(hexes):
            if h in NEUTRALS or h in palette:
                continue
            d = min(dist(h, p) for p in palette)
            if d > 30:
                unknown.append('%s: #%s（距最近配色色 %.0f）' % (src, h, d))
    if unknown:
        results.append(('查2 配色 diff', 'WARN',
                        '以下颜色不属于「%s」配色表及其浅调，逐个判断：主视觉色必须换；'
                        '确属语义扩展色（警示/工位等）的在 00 附注里说明。\n' % theme
                        + '\n'.join(unknown)))
    else:
        n = sum(len(v) for v in found.values())
        results.append(('查2 配色 diff', 'PASS',
                        '%d 个颜色全部命中「%s」配色表、其浅调或中性色' % (n, theme)))
    check_svg_frame(lesson, results)


def check_svg_frame(lesson, results):
    """查2 附项：05 的 SVG 取景检查。失败模式（实测踩坑）：viewBox 直接沿用
    公式坐标系 → 画框大而内容缩在角落，页面被放大、大片空白。
    合格线：内容包围盒在 viewBox 的横/纵两个方向覆盖率都 ≥50%。
    修法：viewBox 手工紧凑取到紧贴图形，svg 加 max-width 限宽居中。"""
    html = os.path.join(lesson, '05_交互演示.html')
    if not os.path.exists(html):
        return
    raw = read(html)
    m = re.search(r'<svg[^>]*viewBox="([^"]+)"', raw)
    if not m:
        results.append(('查2 SVG 取景', 'WARN', '05 里找不到带 viewBox 的 <svg>，跳过取景检查'))
        return
    nums = [float(x) for x in re.split(r'[ ,]+', m.group(1).strip())]
    if len(nums) != 4 or nums[2] <= 0 or nums[3] <= 0:
        return
    vx, vy, vw, vh = nums
    xs, ys = [], []
    for tag in re.finditer(r'<(line|circle|ellipse|rect|polygon|text)\b([^>]*)>',
                           raw):
        t, attrs = tag.group(1), tag.group(2)
        d = dict(re.findall(r'([\w:]+)="(-?[\d.]+)"', attrs))
        try:
            if t == 'line':
                xs += [float(d['x1']), float(d['x2'])]
                ys += [float(d['y1']), float(d['y2'])]
            elif t in ('circle', 'ellipse'):
                cx, cy = float(d.get('cx', 0)), float(d.get('cy', 0))
                r = float(d.get('r', d.get('rx', 0)))
                xs += [cx - r, cx + r]
                ys += [cy - r, cy + r]
            elif t == 'rect':
                x, y = float(d.get('x', 0)), float(d.get('y', 0))
                w_, h_ = float(d.get('width', 0)), float(d.get('height', 0))
                xs += [x, x + w_]
                ys += [y, y + h_]
            elif t == 'polygon':
                for px, py in re.findall(r'(-?[\d.]+)[, ]+(-?[\d.]+)', attrs):
                    xs.append(float(px))
                    ys.append(float(py))
            elif t == 'text' and 'x' in d and 'y' in d:
                xs.append(float(d['x']))
                ys.append(float(d['y']))
        except (KeyError, ValueError):
            continue
    if not xs:
        return
    cover_x = (max(xs) - min(xs)) / vw
    cover_y = (max(ys) - min(ys)) / vh
    worst = min(cover_x, cover_y)
    if worst < 0.5:
        results.append(('查2 SVG 取景', 'FAIL',
                        '05 的 SVG 取景不紧凑：内容占画框横向 %.0f%% / 纵向 %.0f%%（任一方向 <50%% 即不合格），'
                        '渲染出来会整页放大、大片空白。修法：viewBox 手工紧凑取到紧贴图形'
                        '（勿直接沿用公式坐标系），svg 加 max-width 限宽居中。'
                        % (cover_x * 100, cover_y * 100)))
    else:
        results.append(('查2 SVG 取景', 'PASS',
                        '取景紧凑（内容占画框横向 %.0f%% / 纵向 %.0f%%）'
                        % (cover_x * 100, cover_y * 100)))


# ---------------- 查3 ----------------
def slide_count(pptx):
    with zipfile.ZipFile(pptx) as z:
        return sum(1 for n in z.namelist()
                   if re.match(r'ppt/slides/slide\d+\.xml$', n))


def _ref_candidates(lesson):
    for name in ('00_待你确认.md', '01_这一节的设计.md'):
        p = os.path.join(lesson, name)
        if os.path.exists(p):
            yield p
    for base in ('03_讲稿', '04_测验', '06_挑刺报告'):
        p = pick(lesson, base)
        if p:
            yield p


def check_refs(lesson, results):
    pptx = os.path.join(lesson, '02_课件.pptx')
    if not os.path.exists(pptx):
        results.append(('查3 引用存在性', 'WARN', '找不到 02_课件.pptx，跳过'))
        return
    total = slide_count(pptx)
    issues = []
    for p in _ref_candidates(lesson):
        name = os.path.basename(p)
        text = read(p)
        for m in re.finditer(r'第\s*(\d+)\s*页|(?<![A-Za-z0-9])P(\d+)\b', text):
            no = int(m.group(1) or m.group(2))
            ctx = text[max(0, m.start() - 14):m.end() + 14].replace('\n', ' ')
            if no > total:
                issues.append('%s: 引用第 %d 页超出实际 %d 页 ｜ …%s…' % (name, no, total, ctx))
        for m in re.finditer(r'(?<![第每])0*(\d{1,2})\s*页', text):
            no = int(m.group(1))
            ctx = text[max(0, m.start() - 12):m.end() + 12].replace('\n', ' ')
            if m.start() > 0 and text[m.start() - 1] in '.．':
                continue  # 小数尾巴（如 7.08 页脚线）不是页数声明
            if '每页' in ctx:
                continue
            if re.search(r'初版|旧版|原.{0,4}版', ctx):  # 历史版本说明不是总页数声明
                continue
            if re.match(r'^\d', ctx.lstrip('，。、；：')[:1]) or no == total:
                continue  # 数字开头多为数据，不是页数声明
            issues.append('%s: 「%s 页」与实际 %d 页不符 ｜ …%s…' % (name, no, total, ctx))
    if issues:
        results.append(('查3 引用存在性', 'FAIL',
                        '课件共 %d 页。以下引用/声明对不上（引用推荐一律用页面标题，'
                        '不用页码；文件重出后引用方必须同步重出）:\n' % total
                        + '\n'.join(issues)))
    else:
        results.append(('查3 引用存在性', 'PASS',
                        '页引用与总页数声明均与当前 %d 页课件一致（或未用页码引用）' % total))


# ---------------- 汇总 ----------------
def report(lesson, theme, write=True):
    results = []
    ledger = check_numbers(lesson, results)
    check_colors(lesson, theme, results)
    check_refs(lesson, results)

    lines = ['=' * 62, '交付前三查 · %s · 学科族=%s' % (os.path.basename(lesson), theme), '=' * 62]
    for name, verdict, detail in results:
        lines.append('[%s] %s — %s' % (verdict, name, detail))
    if ledger:
        lines.append('---- 查1 数字台账（人工扫一眼跨文件不一致）----')
        lines.extend(ledger)
    fails = sum(1 for r in results if r[1] == 'FAIL')
    warns = sum(1 for r in results if r[1] == 'WARN')
    print('\n'.join(lines))
    print('结论: %d FAIL / %d WARN' % (fails, warns))

    if write:
        p = os.path.join(lesson, '00_待你确认.md')
        if os.path.exists(p):
            text = read(p)
            marker = '## 附：交付前三查'
            if marker in text:
                text = text[:text.index(marker)].rstrip() + '\n'
            stamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            sec = ['---', '', '## 附：交付前三查（check_deliverables.py 自动生成 %s）' % stamp, '']
            for name, verdict, detail in results:
                sec.append('- **%s：%s** %s' % (name, verdict,
                           detail.replace('\n', '；')[:200]))
            if ledger:
                sec.append('- 数字台账 %d 行见运行输出（stdout），本附注不重复' % len(ledger))
            io = __import__('io')
            io.open(p, 'w', encoding='utf-8').write(text + '\n'.join(sec) + '\n')
            print('已写入 00 附注区:', p)
    sys.exit(1 if fails else (2 if warns else 0))


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('lesson')
    ap.add_argument('theme', nargs='?', default='工科')
    ap.add_argument('--no-write', action='store_true')
    a = ap.parse_args()
    report(os.path.abspath(a.lesson), a.theme, write=not a.no_write)
