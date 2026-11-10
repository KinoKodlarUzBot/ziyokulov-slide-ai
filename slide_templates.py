# -*- coding: utf-8 -*-
"""
slide_templates.py - 30 ta premium PowerPoint uslubi (python-pptx).

Ishlatish (main.py ichida):
    import slide_templates as st
    if st.has_style(template_id):
        st.build_presentation(prs, topic, slide_data, template_id, subtitle, tag)
"""
import math
from lxml import etree
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn

SW, SH = 13.333, 7.5
R, O, RR = MSO_SHAPE.RECTANGLE, MSO_SHAPE.OVAL, MSO_SHAPE.ROUNDED_RECTANGLE
TRI, HEX, DIA = MSO_SHAPE.ISOSCELES_TRIANGLE, MSO_SHAPE.HEXAGON, MSO_SHAPE.DIAMOND
ARCH, TEAR, STAR = MSO_SHAPE.ROUND_2_SAME_RECTANGLE, MSO_SHAPE.TEAR, MSO_SHAPE.STAR_5_POINT


# ---------------------------------------------------------------- ranglar
def rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _lum(h):
    r, g, b = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    return 0.299 * r + 0.587 * g + 0.114 * b


def ontext(bg):
    return "1B1B24" if _lum(bg) > 0.62 else "FFFFFF"


def mix(c1, c2, t):
    """c1 ga t ulushda c2 aralashtiriladi (0..1)."""
    a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join("%02X" % round(a[i] * (1 - t) + b[i] * t) for i in range(3))


class Pal(dict):
    __getattr__ = dict.get


# ---------------------------------------------------------------- shakllar
def _paint(sh, fill, line, lw, alpha, grad, dash):
    if grad:
        c1, c2, ang = grad
        sh.fill.gradient()
        sh.fill.gradient_angle = ang
        stops = sh.fill.gradient_stops
        stops[0].color.rgb = rgb(c1)
        stops[0].position = 0
        stops[1].color.rgb = rgb(c2)
        stops[1].position = 1
    elif fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
        if alpha is not None:
            clr = sh._element.spPr.find(qn("a:solidFill"))[0]
            a = etree.SubElement(clr, qn("a:alpha"))
            a.set("val", str(int(alpha * 1000)))
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(lw)
        if dash:
            sh.line.dash_style = MSO_LINE.DASH
    else:
        sh.line.fill.background()


def box(s, kind, x, y, w, h, fill=None, line=None, lw=0.75, alpha=None, grad=None,
        rot=0, rad=None, dash=False, adj=None):
    sh = s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    _paint(sh, fill, line, lw, alpha, grad, dash)
    if rot:
        sh.rotation = rot
    if rad is not None and kind == RR:
        sh.adjustments[0] = max(0.0, min(0.5, rad / min(w, h)))
    if adj:
        for i, v in enumerate(adj):
            sh.adjustments[i] = v
    return sh


def poly(s, pts, fill=None, line=None, lw=1.0, alpha=None, grad=None, close=True):
    e = [(int(Inches(x)), int(Inches(y))) for x, y in pts]
    fb = s.shapes.build_freeform(e[0][0], e[0][1])
    fb.add_line_segments(e[1:], close=close)
    sh = fb.convert_to_shape()
    sh.shadow.inherit = False
    _paint(sh, fill, line, lw, alpha, grad, False)
    return sh


def ring(s, cx, cy, r, col, lw=0.75):
    return box(s, O, cx - r, cy - r, 2 * r, 2 * r, line=col, lw=lw)


def blob_pts(cx, cy, r, k=3, a=0.16, ph=0.0, n=56):
    pts = []
    for i in range(n):
        t = 2 * math.pi * i / n
        rr = r * (1 + a * math.sin(k * t + ph))
        pts.append((cx + rr * math.cos(t), cy + rr * math.sin(t)))
    return pts


def wave(s, base, amp, period, phase, fill, alpha=None, bottom=SH):
    pts, x = [], -0.2
    while x <= SW + 0.25:
        pts.append((x, base + amp * math.sin((x / period) * 2 * math.pi + phase)))
        x += 0.2
    pts += [(SW + 0.25, bottom + 0.1), (-0.2, bottom + 0.1)]
    return poly(s, pts, fill=fill, alpha=alpha)


def bez(p0, p1, p2, p3, n=36):
    out = []
    for i in range(n + 1):
        t = i / n
        u = 1 - t
        out.append((u**3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t**3 * p3[0],
                    u**3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t**3 * p3[1]))
    return out


def brackets(s, col, x0, y0, x1, y1, ln=0.8, t=0.035):
    for (cx, cy, sx, sy) in ((x0, y0, 1, 1), (x1, y0, -1, 1), (x0, y1, 1, -1), (x1, y1, -1, -1)):
        box(s, R, cx if sx > 0 else cx - ln, cy if sy > 0 else cy - t, ln, t, fill=col)
        box(s, R, cx if sx > 0 else cx - t, cy if sy > 0 else cy - ln, t, ln, fill=col)


def grid(s, col, step, w=0.008):
    x = 0.0
    while x <= SW:
        box(s, R, x, 0, w, SH, fill=col)
        x += step
    y = 0.0
    while y <= SH:
        box(s, R, 0, y, SW, w, fill=col)
        y += step


def bg(s, color):
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = rgb(color)


def text(s, x, y, w, h, t, size, color, font="Calibri", bold=False, italic=False,
         al="l", anchor="t", sp=None, ls=None, outline=None):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    p = tf.paragraphs[0]
    p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[al]
    if ls:
        p.line_spacing = ls
    r = p.add_run()
    r.text = t
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font
    rpr = r._r.get_or_add_rPr()
    if outline:
        ln = etree.Element(qn("a:ln"))
        ln.set("w", "15875")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", outline)
        rpr.insert(0, ln)
        rpr.insert(1, etree.Element(qn("a:noFill")))
    else:
        f.color.rgb = rgb(color)
    if sp:
        rpr.set("spc", str(int(sp * 100)))
    return tb


# ---------------------------------------------------------------- matn o'lchami
def fit_title(t, w, h, mx, mn=20):
    for sz in range(mx, mn - 1, -1):
        cpl = max(1, int(w * 72 / (sz * 0.57)))
        lines = math.ceil(len(t) * 1.1 / cpl)
        if lines * sz * 1.17 / 72 <= h:
            return sz
    return mn


def fit_each(texts, w, h, mx=22, mn=12):
    """Har bir matn (w x h) qutiga sig'adigan eng katta shrift."""
    if not texts:
        return mx
    for sz in range(mx, mn - 1, -1):
        cpl = max(1, int(w * 72 / (sz * 0.54)))
        if all(math.ceil(len(t) * 1.06 / cpl) * sz * 1.24 / 72 <= h for t in texts):
            return sz
    return mn


# ---------------------------------------------------------------- kartochka va raqam
def card(s, P, x, y, w, h, fill=None, alpha=None):
    ck = P.ck or {}
    rad = ck.get("rad")
    kind = RR if rad else R
    off = ck.get("shadow")
    if off:
        box(s, kind, x + off, y + off, w, h, fill=ck.get("shc", "000000"), rad=rad)
    a = alpha if alpha is not None else ck.get("alpha")
    return box(s, kind, x, y, w, h, fill=fill or P.card, line=ck.get("line"),
               lw=ck.get("lw", 0.75), alpha=a, rad=rad, dash=ck.get("dash", False))


def _ncol(P, i):
    cyc = P.cycle
    return cyc[i % len(cyc)] if cyc else P.accent


def badge(s, P, x, cy, d, i):
    """Raqam belgisi. x - chap chekka, cy - markaz. Egallagan kenglikni qaytaradi."""
    col = _ncol(P, i)
    nk = P.nk or "circle"
    if nk == "plain":
        text(s, x, cy - d / 2, d * 1.5, d, "%02d" % (i + 1), int(d * 54), col, P.head, bold=True, anchor="m")
        return d * 1.5
    fg = P.nbfg or ontext(col)
    kind = {"circle": O, "square": RR, "hex": HEX}.get(nk, O)
    box(s, kind, x, cy - d / 2, d, d, fill=col, rad=0.08 if kind == RR else None)
    text(s, x, cy - d / 2, d, d, str(i + 1), int(d * 32), fg, P.head, bold=True, al="c", anchor="m")
    return d


# ---------------------------------------------------------------- joylashuvlar
def lay_rows(s, P, items, area, start=0):
    ax, ay, aw, ah = area
    n = len(items)
    g = 0.2
    rh = min((ah - g * (n - 1)) / n, 1.3)
    total = rh * n + g * (n - 1)
    y0 = ay + (ah - total) / 2
    d = min(0.62, rh * 0.62)
    tx0 = ax + 0.3
    bw = d * (1.5 if (P.nk == "plain") else 1)
    tw = aw - (0.3 + bw + 0.3) - 0.3
    size = fit_each(items, tw, rh - 0.2, 28 if n <= 2 else 22, 12)
    for i, it in enumerate(items):
        y = y0 + i * (rh + g)
        card(s, P, ax, y, aw, rh)
        badge(s, P, tx0, y + rh / 2, d, start + i)
        text(s, tx0 + bw + 0.3, y + 0.1, tw, rh - 0.2, it, size, P.text, P.body, anchor="m")


def lay_cols(s, P, items, area, start=0):
    ax, ay, aw, ah = area
    n = len(items)
    gap = 0.32
    cw = (aw - gap * (n - 1)) / n
    ch = min(ah, 4.5)
    y = ay + (ah - ch) / 2
    d = 0.6
    size = fit_each(items, cw - 0.6, ch - 1.75, 22, 12)
    for i, it in enumerate(items):
        x = ax + i * (cw + gap)
        card(s, P, x, y, cw, ch)
        bw = badge(s, P, x + 0.3, y + 0.3 + d / 2, d, start + i)
        box(s, R, x + 0.3, y + 1.12, 0.6, 0.05, fill=_ncol(P, start + i))
        text(s, x + 0.3, y + 1.4, cw - 0.6, ch - 1.65, it, size, P.text, P.body)


def lay_grid(s, P, items, area, start=0):
    ax, ay, aw, ah = area
    n = len(items)
    cols = 2 if n <= 4 else 3
    rows = math.ceil(n / cols)
    g = 0.25
    cw = (aw - g * (cols - 1)) / cols
    ch = min((ah - g * (rows - 1)) / rows, 2.3)
    total = ch * rows + g * (rows - 1)
    y0 = ay + (ah - total) / 2
    d = min(0.56, ch * 0.4)
    bw = d * (1.5 if P.nk == "plain" else 1)
    tw = cw - 0.25 - bw - 0.25 - 0.25
    size = fit_each(items, tw, ch - 0.3, 20, 12)
    for i, it in enumerate(items):
        x = ax + (i % cols) * (cw + g)
        y = y0 + (i // cols) * (ch + g)
        card(s, P, x, y, cw, ch)
        badge(s, P, x + 0.25, y + ch / 2, d, start + i)
        text(s, x + 0.25 + bw + 0.25, y + 0.15, tw, ch - 0.3, it, size, P.text, P.body, anchor="m")


def lay_hero(s, P, items, area, start=0):
    ax, ay, aw, ah = area
    lw_ = aw * 0.5
    col = _ncol(P, start)
    box(s, RR if (P.ck or {}).get("rad") else R, ax, ay, lw_, ah, fill=col,
        rad=(P.ck or {}).get("rad"), line=(P.ck or {}).get("line"), lw=(P.ck or {}).get("lw", 0.75))
    fg = P.nbfg or ontext(col)
    size = fit_each([items[0]], lw_ - 0.9, ah - 1.4, 30, 15)
    text(s, ax + 0.45, ay + 0.35, 1.5, 0.6, "01", 20, fg, P.head, bold=True)
    box(s, R, ax + 0.45, ay + 0.95, 0.7, 0.05, fill=fg)
    text(s, ax + 0.45, ay + 1.2, lw_ - 0.9, ah - 1.5, items[0], size, fg, P.body, bold=True)
    lay_rows(s, P, items[1:], (ax + lw_ + 0.3, ay, aw - lw_ - 0.3, ah), start + 1)


def lay_timeline(s, P, items, area, start=0):
    ax, ay, aw, ah = area
    n = len(items)
    sp = aw / n
    mid = ay + ah / 2
    box(s, R, ax + 0.2, mid - 0.02, aw - 0.4, 0.04, fill=P.muted)
    tw = min(sp * 1.75, 3.7)
    th = ah / 2 - 0.55
    size = fit_each(items, tw, th, 18, 12)
    for i, it in enumerate(items):
        cx = ax + sp * i + sp / 2
        badge(s, P, cx - 0.3, mid, 0.6, start + i)
        if i % 2 == 0:
            text(s, cx - tw / 2, ay, tw, th, it, size, P.text, P.body, al="c", anchor="b")
        else:
            text(s, cx - tw / 2, mid + 0.5, tw, th, it, size, P.text, P.body, al="c", anchor="t")


def _simple_split(items, area, s, P, start):
    lay_rows(s, P, items, area, start)


LAYOUTS = {"rows": lay_rows, "cols": lay_cols, "grid": lay_grid, "hero": lay_hero, "timeline": lay_timeline}


def pick_layout(name, n):
    if n <= 0:
        return None
    if name == "cols" and n > 3:
        name = "grid"
    if name == "timeline" and (n > 5 or n < 3):
        name = "rows"
    if name == "hero" and n < 3:
        name = "rows"
    if name == "grid" and n < 3:
        name = "rows"
    return LAYOUTS[name]


# ================================================================ REYESTR
STYLES, STYLE_NAMES = {}, {}
STYLE_GROUPS = {"uni": [], "lit": [], "teach": [], "trend": []}
ALL_LAYOUTS = ["rows", "cols", "hero", "grid", "timeline"]


def reg(sid, name, groups, pal, cover, frame, ck=None, nk="circle", layouts=None, **extra):
    p = Pal(head="Arial", body="Calibri", ck=ck or {}, nk=nk, layouts=layouts or ALL_LAYOUTS)
    p.update(pal)
    p.update(extra)
    STYLES[sid] = dict(pal=p, cover=cover, frame=frame)
    STYLE_NAMES[sid] = name
    for g in groups:
        STYLE_GROUPS[g].append(sid)


# ---------------------------------------------------------------- 1 Aurora
def cv_aurora(s, P):
    box(s, O, 8.4, -1.8, 6.6, 6.6, fill=P.accent, alpha=35)
    box(s, O, 10.0, 3.5, 4.8, 4.8, fill=P.accent2, alpha=30)
    box(s, O, 7.3, 4.4, 2.0, 2.0, fill="FFFFFF", alpha=7)
    return dict(t=(0.9, 2.45, 7.3, 2.7), tag=(0.9, 1.75, 6), sub=(0.9, 5.4, 6.8), tg=P.accent2)


def fr_aurora(s, P, pg):
    box(s, O, 10.6, -1.6, 4.2, 4.2, fill=P.accent, alpha=18)
    box(s, O, -1.4, 5.9, 3.6, 3.6, fill=P.accent2, alpha=14)
    box(s, R, 0.8, 1.78, 1.2, 0.06, fill=P.accent2)
    return dict(t=(0.8, 0.55, 10.2, 1.1), area=(0.8, 2.15, 11.73, 4.65))


reg("aurora", "Aurora", ["trend"],
    dict(bg="0A1022", ink="FFFFFF", text="C3CCE6", accent="7C5CFF", accent2="1DD3B0", card="141B36", muted="5A6690"),
    cv_aurora, fr_aurora, ck=dict(rad=0.18, line="2A3466"), layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 2 Bento
def cv_bento(s, P):
    box(s, RR, 7.7, 0.8, 4.9, 3.0, fill=P.accent, rad=0.3)
    box(s, RR, 7.7, 4.0, 2.35, 2.7, fill=P.accent2, rad=0.3)
    box(s, RR, 10.25, 4.0, 2.35, 2.7, fill=P.ink, rad=0.3)
    box(s, O, 11.1, 1.3, 1.1, 1.1, fill="FFFFFF", alpha=30)
    return dict(t=(0.9, 2.3, 6.5, 2.8), tag=(0.9, 1.7, 5), sub=(0.9, 5.35, 6.3))


def fr_bento(s, P, pg):
    for i, c in enumerate([P.accent, P.accent2, P.ink]):
        box(s, RR, 11.2 + i * 0.55, 0.62, 0.42, 0.42, fill=c, rad=0.1)
    return dict(t=(0.7, 0.5, 10.2, 1.0), area=(0.7, 1.85, 11.93, 4.95))


reg("bento", "Bento", ["trend"],
    dict(bg="F2F2F7", ink="1C1C28", text="4B4B5E", accent="5B6CFF", accent2="FFD166", card="FFFFFF", muted="A0A0B2",
         cycle=["5B6CFF", "FFD166", "1C1C28", "FF7A8A"]),
    cv_bento, fr_bento, ck=dict(rad=0.22), nk="square", layouts=["grid", "rows", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 3 Brutal
def cv_brutal(s, P):
    box(s, R, 0.92, 1.42, 8.6, 5.1, fill="000000")
    box(s, R, 0.8, 1.3, 8.6, 5.1, fill="FFFFFF", line="000000", lw=3)
    box(s, STAR, 10.0, 0.9, 2.0, 2.0, fill=P.accent, line="000000", lw=3)
    box(s, O, 10.6, 4.6, 2.2, 2.2, fill=P.accent2, line="000000", lw=3)
    box(s, R, 10.0, 3.3, 1.2, 0.5, fill="FFFFFF", line="000000", lw=3)
    return dict(t=(1.3, 2.2, 7.7, 2.7), tag=(1.3, 1.65, 6), sub=(1.3, 5.2, 7.7), ink="000000", tg="000000", tx="222222", tm=42)


def fr_brutal(s, P, pg):
    box(s, R, 0.82, 0.57, 9.9, 1.15, fill="000000")
    box(s, R, 0.7, 0.45, 9.9, 1.15, fill="FFFFFF", line="000000", lw=3)
    box(s, O, 11.5, 0.5, 1.0, 1.0, fill=P.accent2, line="000000", lw=3)
    return dict(t=(0.95, 0.5, 9.4, 1.05), tm=28, area=(0.7, 2.1, 11.93, 4.7), ink="000000", pn=(SW - 1.9, 7.02, 1.3, "r", "000000"))


reg("brutal", "Brutal", ["lit"],
    dict(bg="FFE14D", ink="000000", text="111111", accent="FF5D8F", accent2="4D9DFF", card="FFFFFF", muted="000000",
         cycle=["FF5D8F", "4D9DFF", "7BE05A"]),
    cv_brutal, fr_brutal, ck=dict(line="000000", lw=2.5, shadow=0.09), layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 4 Tungi oltin
def cv_gold(s, P):
    ring(s, 10.7, 2.8, 3.2, mix(P.accent, P.bg, 0.68))
    ring(s, 10.7, 2.8, 2.2, mix(P.accent, P.bg, 0.4), 1)
    box(s, O, 10.58, 2.68, 0.24, 0.24, fill=P.accent)
    box(s, R, 1.0, 1.5, 0.9, 0.04, fill=P.accent)
    box(s, R, 1.0, 5.2, 6.5, 0.012, fill=mix(P.muted, P.bg, 0.3))
    return dict(t=(1.0, 2.35, 8.3, 2.6), tag=(1.0, 1.75, 6), sub=(1.0, 5.4, 8.2), tm=46)


def fr_gold(s, P, pg):
    box(s, R, 0.9, 1.68, 0.9, 0.045, fill=P.accent)
    box(s, R, 0.9, 6.98, 11.53, 0.01, fill=mix(P.muted, P.bg, 0.3))
    return dict(t=(0.9, 0.6, 11.5, 1.0), area=(0.9, 2.1, 11.53, 4.6), pn=(SW - 2.2, 7.08, 1.3, "r", P.muted))


reg("midnight_gold", "Tungi oltin", ["uni"],
    dict(bg="0B1B2B", ink="F5F1E8", text="B4C2D0", accent="D4A24C", accent2="8FA3B8", card="12283D", muted="5E7388",
         head="Cambria", tb=False),
    cv_gold, fr_gold, ck=dict(rad=0.06, line="24415C"), nk="plain", layouts=["rows", "cols", "hero", "grid", "timeline"])


# ---------------------------------------------------------------- 5 Pastel
def cv_pastel(s, P):
    box(s, O, 8.6, -1.6, 6.0, 6.0, fill=P.c1)
    box(s, O, -1.4, 4.9, 4.6, 4.6, fill=P.c2)
    box(s, O, 9.6, 4.6, 2.8, 2.8, fill=P.c3)
    return dict(t=(0.9, 2.35, 7.6, 2.5), tag=(0.9, 1.7, 6), sub=(0.9, 5.2, 7.2))


def fr_pastel(s, P, pg):
    box(s, O, 11.3, -1.3, 3.2, 3.2, fill=P.c1)
    box(s, O, -1.3, 6.3, 2.9, 2.9, fill=P.c2)
    box(s, O, 12.3, 6.5, 1.8, 1.8, fill=P.c3, alpha=70)
    return dict(t=(0.8, 0.5, 10.2, 1.1), area=(0.8, 2.0, 11.73, 4.75))


reg("pastel", "Pastel", ["lit"],
    dict(bg="FFF3F0", ink="4A3F6B", text="6B5E8A", accent="8E6BBF", accent2="F28DB2", card="FFFFFF", muted="B9A6D4",
         c1="FFC8DD", c2="BDE0FE", c3="CDB4DB", cycle=["F28DB2", "8E6BBF", "6FA8DC"]),
    cv_pastel, fr_pastel, ck=dict(rad=0.3), layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 6 Shveysariya
def cv_swiss(s, P):
    box(s, R, 0, 0, 4.6, SH, fill=P.accent)
    text(s, 0.45, 4.7, 4, 2.4, "01", 130, "FFFFFF", P.head, bold=True, anchor="b")
    box(s, R, 5.3, 5.15, 7.2, 0.025, fill=P.ink)
    return dict(t=(5.3, 1.9, 7.3, 3.0), tag=(5.3, 1.35, 6), sub=(5.3, 5.35, 7.2), tm=44)


def fr_swiss(s, P, pg):
    box(s, R, 0, 0, 0.38, SH, fill=P.accent)
    box(s, R, 1.0, 1.78, 11.53, 0.025, fill=P.ink)
    return dict(t=(1.0, 0.6, 11.4, 1.05), area=(1.0, 2.15, 11.53, 4.65))


reg("swiss", "Shveysariya", ["uni"],
    dict(bg="FFFFFF", ink="111111", text="3A3A3A", accent="E63946", accent2="111111", card="F3F3F3", muted="9A9A9A"),
    cv_swiss, fr_swiss, ck=dict(), nk="plain", layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 7 Neon
def cv_neon(s, P):
    brackets(s, P.accent, 0.5, 0.5, SW - 0.5, SH - 0.5, 0.9, 0.04)
    box(s, O, 9.4, 0.8, 4.2, 4.2, fill=P.accent, alpha=6)
    box(s, O, 10.6, 3.6, 3.0, 3.0, fill=P.accent2, alpha=6)
    return dict(t=(1.3, 2.35, 10.4, 2.6), tag=(1.3, 1.7, 8), sub=(1.3, 5.3, 10), tg=P.accent2, tm=44)


def fr_neon(s, P, pg):
    brackets(s, P.accent, 0.45, 0.45, SW - 0.45, SH - 0.45, 0.5, 0.035)
    box(s, R, 0.9, 1.72, 11.53, 0.02, fill=P.muted)
    return dict(t=(0.9, 0.6, 11.5, 1.0), area=(0.9, 2.1, 11.53, 4.7))


reg("neon", "Neon", ["trend"],
    dict(bg="05070A", ink="39FF88", text="A9DEC0", accent="39FF88", accent2="00E5FF", card="0B1410", muted="2B6B47",
         head="Courier New"),
    cv_neon, fr_neon, ck=dict(rad=0.05, line="39FF88", lw=1), nk="square", layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 8 Qum
def cv_sand(s, P):
    box(s, ARCH, 8.7, 2.2, 3.7, 6.0, fill=P.accent, adj=[0.5, 0])
    box(s, ARCH, 9.3, 2.9, 2.5, 5.4, fill=P.accent2, adj=[0.5, 0])
    return dict(t=(0.9, 2.3, 7.4, 2.6), tag=(0.9, 1.7, 6), sub=(0.9, 5.3, 7.2), tm=44)


def fr_sand(s, P, pg):
    box(s, ARCH, 11.9, 6.6, 1.0, 1.4, fill=P.accent, adj=[0.5, 0])
    box(s, R, 0.9, 1.68, 0.9, 0.045, fill=P.accent)
    return dict(t=(0.9, 0.6, 10.4, 1.0), area=(0.9, 2.0, 11.53, 4.5), pn=(SW - 3.2, 7.02, 1.3, "r", P.muted))


reg("sand", "Qum", ["teach"],
    dict(bg="F3E9DC", ink="5A3A2A", text="6E5040", accent="C8643C", accent2="E8A87C", card="FBF4EA", muted="B79A82",
         head="Cambria", tb=False),
    cv_sand, fr_sand, ck=dict(rad=0.14), layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 9 Okean
def cv_ocean(s, P):
    wave(s, 5.0, 0.3, 4.6, 0.0, P.accent2)
    wave(s, 5.6, 0.25, 5.4, 1.5, P.accent)
    return dict(t=(0.9, 1.4, 10.8, 2.6), tag=(0.9, 0.85, 8), sub=(0.9, 4.05, 10.5), tm=46)


def fr_ocean(s, P, pg):
    wave(s, 6.95, 0.1, 3.2, 0.0, P.accent2)
    wave(s, 7.15, 0.08, 2.8, 1.0, P.accent)
    return dict(t=(0.8, 0.5, 11.7, 1.05), area=(0.8, 1.9, 11.73, 4.55), pn=(SW - 2.0, 6.6, 1.3, "r", P.muted))


reg("ocean", "Okean", ["teach"],
    dict(bg="FFFFFF", ink="023E8A", text="35546F", accent="0077B6", accent2="00B4D8", card="EAF6FB", muted="7FBBD6",
         cycle=["0077B6", "00B4D8", "48CAE4"]),
    cv_ocean, fr_ocean, ck=dict(rad=0.16), layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 10 O'rmon
def cv_forest(s, P):
    box(s, TEAR, 9.6, 0.7, 3.4, 3.4, fill=P.accent2, rot=45)
    box(s, TEAR, 10.4, 3.7, 2.6, 2.6, fill=mix(P.accent2, P.bg, 0.5), rot=20)
    ring(s, 9.3, 5.7, 1.1, mix(P.accent, P.bg, 0.55))
    box(s, R, 1.0, 5.15, 0.9, 0.04, fill=P.accent)
    return dict(t=(1.0, 2.3, 7.9, 2.6), tag=(1.0, 1.7, 6), sub=(1.0, 5.4, 7.6), tm=46)


def fr_forest(s, P, pg):
    box(s, TEAR, 11.9, -0.7, 2.2, 2.2, fill=P.accent2, rot=45)
    box(s, R, 0.9, 1.68, 0.9, 0.045, fill=P.accent)
    return dict(t=(0.9, 0.6, 10.4, 1.0), area=(0.9, 2.1, 11.53, 4.6))


reg("forest", "O'rmon", ["teach"],
    dict(bg="14342B", ink="F1EFE3", text="BCD2C4", accent="7FD1AE", accent2="2D6A4F", card="1C4438", muted="5C9580",
         head="Cambria", tb=False),
    cv_forest, fr_forest, ck=dict(rad=0.12, line="2E6B58"), layouts=["rows", "grid", "cols", "hero", "timeline"])



# ---------------------------------------------------------------- 11 Akademik
def _frame2(s, P):
    box(s, R, 0.45, 0.45, SW - 0.9, SH - 0.9, line=P.ink, lw=2)
    box(s, R, 0.62, 0.62, SW - 1.24, SH - 1.24, line=P.ink, lw=0.5)


def cv_academic(s, P):
    _frame2(s, P)
    box(s, R, 6.17, 5.0, 1.0, 0.03, fill=P.accent)
    return dict(t=(1.6, 2.2, 10.13, 2.6), al="c", tag=(1.6, 1.55, 10.13), sub=(1.6, 5.25, 10.13), tm=46)


def fr_academic(s, P, pg):
    _frame2(s, P)
    box(s, R, 6.17, 1.95, 1.0, 0.03, fill=P.accent)
    return dict(t=(1.2, 0.85, 10.93, 1.0), al="c", tm=30, area=(1.2, 2.3, 10.93, 4.2), pn=(6.0, 6.5, 1.3, "c", P.muted))


reg("academic", "Akademik", ["uni"],
    dict(bg="FBF7EE", ink="7B1E3A", text="4A4036", accent="B8892D", accent2="7B1E3A", card="FFFFFF", muted="8A7A5A",
         head="Cambria", tb=False),
    cv_academic, fr_academic, ck=dict(line="D9CDB3"), nk="plain", layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 12 Diagonal
def cv_diag(s, P):
    poly(s, [(0, 0), (8.6, 0), (5.4, SH), (0, SH)], fill=P.ink)
    poly(s, [(8.6, 0), (9.8, 0), (6.6, SH), (5.4, SH)], fill=P.accent)
    return dict(t=(0.9, 2.2, 6.0, 2.7), tag=(0.9, 1.6, 5), sub=(0.9, 5.2, 5.4), ink="FFFFFF", tx="C5CCD8", tm=40)


def fr_diag(s, P, pg):
    poly(s, [(11.2, 0), (SW, 0), (SW, 2.1)], fill=P.accent)
    poly(s, [(0, SH), (0, 6.0), (1.5, SH)], fill=P.ink)
    return dict(t=(0.8, 0.55, 10.2, 1.1), area=(0.8, 2.05, 11.73, 4.7), pn=(SW - 3.0, 7.02, 1.3, "r", P.muted))


reg("diagonal", "Diagonal", ["uni"],
    dict(bg="FFFFFF", ink="12284C", text="3F4E66", accent="FF7A00", accent2="12284C", card="F2F5FA", muted="9AA7BC",
         cycle=["FF7A00", "12284C"]),
    cv_diag, fr_diag, ck=dict(rad=0.1), nk="square", layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 13 Gradient
def cv_grad(s, P):
    box(s, R, 0, 0, SW, SH, grad=("7B2FF7", "2D6BFF", 45))
    box(s, O, 9.3, 0.5, 2.9, 2.9, fill="FFFFFF", alpha=14)
    box(s, O, 11.6, 2.9, 1.5, 1.5, fill="FFFFFF", alpha=14)
    box(s, O, 8.4, 4.7, 3.6, 3.6, fill="FFFFFF", alpha=10)
    return dict(t=(1.0, 2.35, 8.0, 2.7), tag=(1.0, 1.7, 6), sub=(1.0, 5.3, 7.5), ink="FFFFFF", tx="E1DBFF", tg="FFFFFF", tm=48)


def fr_grad(s, P, pg):
    box(s, R, 0, 0, SW, 1.75, grad=("7B2FF7", "2D6BFF", 0))
    box(s, O, 11.6, -0.9, 2.4, 2.4, fill="FFFFFF", alpha=12)
    return dict(t=(0.8, 0.35, 10.7, 1.05), ink="FFFFFF", area=(0.8, 2.15, 11.73, 4.6))


reg("gradient", "Gradient", ["trend"],
    dict(bg="F7F5FF", ink="2A1B6B", text="3B3A5C", accent="7B2FF7", accent2="2D6BFF", card="FFFFFF", muted="A7A3C9",
         cycle=["7B2FF7", "2D6BFF"]),
    cv_grad, fr_grad, ck=dict(rad=0.18, line="E4DEFA"), layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 14 Organik
def cv_blob(s, P):
    poly(s, blob_pts(10.6, 3.4, 3.4, 3, 0.16, 0.4), fill=P.accent2)
    poly(s, blob_pts(11.4, 4.0, 1.9, 2, 0.14, 1.2), fill=P.accent, alpha=80)
    poly(s, blob_pts(8.6, 6.4, 1.1, 3, 0.15, 2.0), fill=P.accent2, alpha=60)
    return dict(t=(0.9, 2.35, 7.6, 2.6), tag=(0.9, 1.7, 6), sub=(0.9, 5.2, 7.2))


def fr_blob(s, P, pg):
    poly(s, blob_pts(12.9, 0.3, 1.9, 3, 0.15, 0.3), fill=P.accent2, alpha=70)
    poly(s, blob_pts(0.2, 7.4, 1.7, 3, 0.15, 1.0), fill=P.accent2, alpha=50)
    return dict(t=(0.8, 0.55, 10.3, 1.0), area=(0.8, 2.0, 11.73, 4.7))


reg("blob", "Organik", ["lit"],
    dict(bg="E9FBF4", ink="0B4F3A", text="3C6E5C", accent="1FA97A", accent2="7DE0B5", card="FFFFFF", muted="8CCFB2",
         cycle=["1FA97A", "34C38F", "0F8A62"]),
    cv_blob, fr_blob, ck=dict(rad=0.28), layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 15 Memphis
def _zig(s, x, y, n, col, lw=5):
    poly(s, [(x + i * 0.3, y + (0.16 if i % 2 else 0)) for i in range(n)], line=col, lw=lw, close=False)


def cv_memphis(s, P):
    box(s, O, 10.8, 0.8, 1.5, 1.5, fill=P.accent)
    box(s, TRI, 11.0, 4.9, 2.0, 1.8, fill=P.accent2)
    box(s, R, 9.6, 3.2, 0.9, 0.9, fill=P.ink, rot=20)
    _zig(s, 0.9, 6.2, 10, "FFD93D")
    for i in range(5):
        box(s, O, 9.0 + i * 0.32, 6.7, 0.16, 0.16, fill=P.accent)
    return dict(t=(0.9, 2.2, 8.2, 2.7), tag=(0.9, 1.6, 6), sub=(0.9, 5.15, 7.8))


def fr_memphis(s, P, pg):
    box(s, O, 12.0, 0.4, 0.7, 0.7, fill=P.accent)
    box(s, TRI, 12.2, 6.55, 0.8, 0.7, fill=P.accent2)
    _zig(s, 0.8, 7.05, 8, "FFD93D", 4)
    return dict(t=(0.8, 0.55, 10.8, 1.05), area=(0.8, 2.0, 11.73, 4.75))


reg("memphis", "Memphis", ["lit"],
    dict(bg="FFFFFF", ink="1A535C", text="3D6B73", accent="FF6B6B", accent2="4ECDC4", card="FFF9E6", muted="9DBFC2",
         cycle=["FF6B6B", "4ECDC4", "FFC93C"]),
    cv_memphis, fr_memphis, ck=dict(rad=0.14, line="1A535C", lw=1.5), layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 16 Doska
def _wood(s, P):
    box(s, R, 0, 0, SW, SH, line="8B5E3C", lw=22)


def cv_chalk(s, P):
    _wood(s, P)
    box(s, O, 9.7, 1.0, 2.4, 2.4, line=P.accent2, lw=1.5, dash=True)
    text(s, 8.9, 4.3, 4.0, 1.0, "E = mc\u00b2", 32, P.accent2, "Cambria", italic=True, al="c")
    box(s, R, 1.0, 5.0, 5.5, 0.025, fill=P.ink)
    return dict(t=(1.0, 2.2, 7.9, 2.7), tag=(1.0, 1.65, 6), sub=(1.0, 5.25, 7.5), tg=P.accent, tm=46)


def fr_chalk(s, P, pg):
    _wood(s, P)
    box(s, R, 0.9, 1.75, 3.6, 0.025, fill=P.ink)
    return dict(t=(0.9, 0.65, 11.5, 1.0), area=(0.9, 2.1, 11.53, 4.6))


reg("chalkboard", "Doska", ["teach"],
    dict(bg="1F3B34", ink="F2F2EA", text="D5E6DB", accent="F4D35E", accent2="9BD1B8", card="27483F", muted="7FA597",
         head="Cambria", tb=False, ti=True, cycle=["F4D35E", "9BD1B8", "F29E8E"]),
    cv_chalk, fr_chalk, ck=dict(rad=0.08, line="5C8A7B", dash=True), layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 17 Daftar
def _ruled(s, P):
    y = 1.0
    while y < 7.2:
        box(s, R, 0, y, SW, 0.007, fill="CFE0F5")
        y += 0.42
    box(s, R, 1.4, 0, 0.012, SH, fill="F0A0A0")
    for yy in (1.2, 3.7, 6.2):
        box(s, O, 0.45, yy, 0.26, 0.26, fill="E8E8E2", line="D6D6CE")


def cv_notebook(s, P):
    _ruled(s, P)
    box(s, R, 1.9, 4.95, 3.6, 0.035, fill=P.accent)
    return dict(t=(1.9, 2.15, 10.2, 2.7), tag=(1.9, 1.6, 6), sub=(1.9, 5.2, 9.5), tm=46)


def fr_notebook(s, P, pg):
    _ruled(s, P)
    return dict(t=(1.9, 0.6, 10.6, 1.0), area=(1.9, 2.0, 10.63, 4.8))


reg("notebook", "Daftar", ["teach"],
    dict(bg="FFFEF5", ink="2B3A67", text="3B4B7A", accent="E4572E", accent2="2B3A67", card="FFFDF0", muted="A9B4D0",
         head="Cambria", tb=False, ti=True),
    cv_notebook, fr_notebook, ck=dict(line="CFE0F5"), nk="plain", layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 18 Monoxrom
def cv_mono(s, P):
    box(s, O, 11.9, 0.9, 0.5, 0.5, fill=P.accent)
    box(s, R, 0.9, 5.55, 11.53, 0.012, fill=P.ink)
    return dict(t=(0.9, 2.9, 11.2, 2.4), tag=(0.9, 2.3, 6), sub=(0.9, 5.75, 11), tg=P.ink, tm=54)


def fr_mono(s, P, pg):
    box(s, O, 12.3, 0.72, 0.22, 0.22, fill=P.accent)
    box(s, R, 0.9, 1.85, 11.53, 0.012, fill=P.ink)
    return dict(t=(0.9, 0.65, 10.8, 1.05), area=(0.9, 2.2, 11.53, 4.55))


reg("mono", "Monoxrom", ["uni"],
    dict(bg="FFFFFF", ink="111111", text="444444", accent="E5322D", accent2="111111", card="F4F4F4", muted="AAAAAA"),
    cv_mono, fr_mono, ck=dict(), nk="plain", layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 19 Shisha
def cv_glass(s, P):
    box(s, R, 0, 0, SW, SH, grad=("00C6A7", "1E4AE9", 35))
    box(s, O, 0.3, 0.3, 3.2, 3.2, fill="FFFFFF", alpha=15)
    box(s, O, 10.2, 4.6, 4.2, 4.2, fill="FFFFFF", alpha=14)
    box(s, RR, 0.9, 1.2, 9.7, 5.1, fill="FFFFFF", alpha=20, line="D6EEFF", lw=0.75, rad=0.35)
    return dict(t=(1.5, 2.15, 8.5, 2.6), tag=(1.5, 1.7, 6), sub=(1.5, 5.1, 8.5), tg="FFFFFF", tm=44)


def fr_glass(s, P, pg):
    box(s, R, 0, 0, SW, SH, grad=("00C6A7", "1E4AE9", 35))
    box(s, O, -0.8, -0.8, 3.0, 3.0, fill="FFFFFF", alpha=14)
    box(s, O, 11.2, 5.4, 3.6, 3.6, fill="FFFFFF", alpha=12)
    return dict(t=(0.9, 0.55, 11.5, 1.0), area=(0.9, 1.95, 11.53, 4.85))


reg("glass", "Shisha", ["trend"],
    dict(bg="0F5FD6", ink="FFFFFF", text="FFFFFF", accent="FFFFFF", accent2="A8F0E4", card="FFFFFF", muted="D3EAF8",
         nbfg="1E4AE9"),
    cv_glass, fr_glass, ck=dict(rad=0.22, line="D6EEFF", alpha=20), layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 20 Quyosh botishi
def cv_sunset(s, P):
    box(s, R, 0, 0, SW, SH, grad=("FF7A59", "C13584", 90))
    box(s, O, 4.8, 3.7, 3.9, 3.9, fill="FFE29A", alpha=92)
    poly(s, [(0, SH), (0, 5.7), (2.6, 4.4), (5.2, 6.2), (8.6, 3.9), (SW, 6.3), (SW, SH)], fill="3D1E6D")
    return dict(t=(0.9, 0.85, 11.5, 2.0), tag=None, sub=(0.9, 3.0, 11), ink="FFFFFF", tx="FFF1DD", tm=50)


def fr_sunset(s, P, pg):
    box(s, R, 0, 0, SW, 1.7, grad=("FF7A59", "C13584", 0))
    poly(s, [(0, SH), (0, 7.0), (2, 6.7), (4, 7.1), (7, 6.6), (10, 7.0), (SW, 6.7), (SW, SH)], fill="3D1E6D")
    return dict(t=(0.8, 0.35, 11.7, 1.0), ink="FFFFFF", area=(0.8, 2.0, 11.73, 4.5), pn=(SW - 1.9, 7.14, 1.3, "r", "FFFFFF"))


reg("sunset", "Quyosh botishi", ["trend"],
    dict(bg="FFF6F0", ink="4A2D5E", text="5A3D6B", accent="C13584", accent2="FF7A59", card="FFFFFF", muted="C9A9B8",
         cycle=["C13584", "FF7A59", "6D3FA0"]),
    cv_sunset, fr_sunset, ck=dict(rad=0.18, line="F3D8E6"), layouts=["rows", "cols", "grid", "hero", "timeline"])



# ---------------------------------------------------------------- 21 Tex-setka
def _node(s, col, x, y, r=0.08):
    box(s, O, x - r, y - r, 2 * r, 2 * r, fill=col)


def cv_tech(s, P):
    grid(s, "0F2A5C", 0.9)
    for pts, col in (([(9.6, 1.2), (11.2, 1.2), (11.2, 2.6), (12.6, 2.6)], P.accent),
                     ([(9.0, 5.8), (10.4, 5.8), (10.4, 4.6), (12.6, 4.6)], P.accent2)):
        poly(s, pts, line=col, lw=1.5, close=False)
        _node(s, col, *pts[0])
        _node(s, col, *pts[-1])
    return dict(t=(0.9, 2.35, 8.6, 2.7), tag=(0.9, 1.7, 6), sub=(0.9, 5.3, 8), tm=44)


def fr_tech(s, P, pg):
    grid(s, "0C2250", 1.0)
    pts = [(11.4, 0.5), (12.2, 0.5), (12.2, 1.05), (12.9, 1.05)]
    poly(s, pts, line=P.accent, lw=1.5, close=False)
    _node(s, P.accent, *pts[0])
    _node(s, P.accent, *pts[-1])
    return dict(t=(0.9, 0.6, 10.2, 1.0), area=(0.9, 2.1, 11.53, 4.7))


reg("techgrid", "Tex-setka", ["uni"],
    dict(bg="07142E", ink="E6F4FF", text="B0C6E2", accent="4CC9F0", accent2="7B8CFF", card="0D1F45", muted="2A5C9E",
         cycle=["4CC9F0", "7B8CFF"]),
    cv_tech, fr_tech, ck=dict(rad=0.08, line="1E4A8A"), nk="square", layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 22 Atirgul oltin
def cv_rose(s, P):
    ring(s, 9.8, 3.75, 2.9, P.accent)
    ring(s, 9.8, 3.75, 2.5, mix(P.accent, P.bg, 0.5), 0.5)
    ring(s, 9.8, 3.75, 1.2, P.accent2, 1.25)
    box(s, O, 9.7, 3.65, 0.2, 0.2, fill=P.accent)
    box(s, R, 0.9, 1.5, 1.2, 0.02, fill=P.accent)
    return dict(t=(0.9, 2.3, 7.5, 2.7), tag=(0.9, 1.75, 6), sub=(0.9, 5.25, 7.2), tm=46)


def fr_rose(s, P, pg):
    ring(s, 12.4, 0.3, 1.5, P.accent)
    ring(s, 12.4, 0.3, 1.15, P.accent2, 1)
    box(s, R, 0.9, 1.68, 1.2, 0.02, fill=P.accent)
    return dict(t=(0.9, 0.6, 10.2, 1.0), area=(0.9, 2.05, 11.53, 4.7))


reg("rosegold", "Atirgul oltin", ["uni"],
    dict(bg="FBEFEA", ink="8A4B3C", text="85615A", accent="C98B7B", accent2="E3B7A8", card="FFF8F5", muted="D2AFA3",
         head="Cambria", tb=False, ti=True),
    cv_rose, fr_rose, ck=dict(rad=0.2, line="E8CFC6"), layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 23 Chiziqli
STRIPES = ["3A86FF", "8338EC", "FF006E", "FB5607"]


def cv_stripe(s, P):
    for i, c in enumerate(STRIPES):
        box(s, R, i * 0.6, 0, 0.6, SH, fill=c)
    return dict(t=(3.3, 2.3, 9.0, 2.7), tag=(3.3, 1.7, 6), sub=(3.3, 5.25, 8.5), tm=46)


def fr_stripe(s, P, pg):
    for i, c in enumerate(STRIPES):
        box(s, R, i * 0.2, 0, 0.2, SH, fill=c)
    return dict(t=(1.4, 0.6, 11.1, 1.05), area=(1.4, 2.0, 11.13, 4.8))


reg("stripe", "Chiziqli", ["lit"],
    dict(bg="FFFFFF", ink="111111", text="4A4A4A", accent="8338EC", accent2="FF006E", card="F6F6F9", muted="A0A0A0",
         cycle=STRIPES),
    cv_stripe, fr_stripe, ck=dict(rad=0.1), layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 24 Olti burchak
def cv_hex(s, P):
    box(s, HEX, 9.2, 0.8, 2.8, 2.4, fill=P.accent)
    box(s, HEX, 11.4, 2.7, 2.0, 1.7, fill=P.accent2)
    box(s, HEX, 8.6, 3.4, 1.6, 1.4, fill=P.accent, alpha=50)
    box(s, HEX, 10.0, 4.9, 2.6, 2.2, fill=P.accent2, alpha=70)
    return dict(t=(0.9, 2.3, 7.6, 2.7), tag=(0.9, 1.7, 6), sub=(0.9, 5.25, 7.2))


def fr_hex(s, P, pg):
    box(s, HEX, 11.7, 0.3, 1.2, 1.04, fill=P.accent)
    box(s, HEX, 12.5, 1.15, 0.9, 0.78, fill=P.accent2, alpha=80)
    box(s, HEX, -0.3, 6.6, 1.0, 0.87, fill=P.accent2, alpha=60)
    return dict(t=(0.8, 0.55, 10.5, 1.0), area=(0.8, 2.0, 11.73, 4.75))


reg("hexagon", "Olti burchak", ["lit"],
    dict(bg="E8F7F5", ink="124E48", text="3F6E69", accent="2A9D8F", accent2="76C7BC", card="FFFFFF", muted="9BD3CC",
         cycle=["2A9D8F", "1D7A70", "4DB6A8"]),
    cv_hex, fr_hex, ck=dict(rad=0.1), nk="hex", layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 25 Yo'l xaritasi
def cv_road(s, P):
    pts = bez((0.8, 6.5), (3.0, 4.3), (5.0, 7.5), (7.2, 5.7)) + bez((7.2, 5.7), (9.0, 4.2), (11.0, 6.0), (12.5, 4.2))
    poly(s, pts, line=P.accent2, lw=9, close=False)
    for i in (0, 18, 36, 54, 72):
        x, y = pts[min(i, len(pts) - 1)]
        box(s, O, x - 0.18, y - 0.18, 0.36, 0.36, fill=P.accent, line="FFFFFF", lw=2.5)
    return dict(t=(0.9, 0.9, 11.4, 2.0), tag=(0.9, 0.45, 6), sub=(0.9, 3.05, 10.5), tm=46)


def fr_road(s, P, pg):
    for i in range(24):
        box(s, R, 0.8 + i * 0.5, 7.05, 0.28, 0.05, fill=P.accent2)
    box(s, O, 12.1, 0.6, 0.5, 0.5, fill=P.accent)
    return dict(t=(0.8, 0.55, 11.0, 1.05), area=(0.8, 1.9, 11.73, 4.9), pn=(SW - 1.9, 6.7, 1.3, "r", P.muted))


reg("roadmap", "Yo'l xaritasi", ["teach"],
    dict(bg="FFF7E8", ink="6B3E26", text="7F5B45", accent="E76F51", accent2="F4A261", card="FFFFFF", muted="D2B08C",
         cycle=["E76F51", "F4A261", "2A9D8F"]),
    cv_road, fr_road, ck=dict(rad=0.16, line="F3DFC4"), layouts=["timeline", "rows", "cols", "grid", "hero"])


# ---------------------------------------------------------------- 26 Katta raqam
def cv_bignum(s, P):
    text(s, 0.2, 0.3, 7.4, 6.9, "01", 300, P.ink, P.head, bold=True, anchor="m", outline="3A3A4C")
    box(s, R, 6.9, 5.25, 0.9, 0.06, fill=P.accent)
    return dict(t=(6.9, 2.3, 5.9, 2.8), tag=(6.9, 1.7, 5), sub=(6.9, 5.5, 5.9), tm=40)


def fr_bignum(s, P, pg):
    text(s, 8.6, -0.5, 4.9, 3.0, "%02d" % pg, 150, P.ink, P.head, bold=True, al="r", outline="2A2A36")
    return dict(t=(0.9, 0.65, 8.6, 1.1), area=(0.9, 2.1, 11.53, 4.7))


reg("bignum", "Katta raqam", ["trend"],
    dict(bg="0F0F14", ink="FFFFFF", text="B8B8CA", accent="C6FF3D", accent2="7B61FF", card="1A1A22", muted="5B5B6E",
         cycle=["C6FF3D", "7B61FF"]),
    cv_bignum, fr_bignum, ck=dict(rad=0.12, line="2E2E3A"), nk="plain", layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 27 Foto ramka
def cv_photo(s, P):
    box(s, R, 0.7, 0.7, 11.93, 4.1, fill="CBD5E1")
    box(s, O, 9.6, 1.2, 0.9, 0.9, fill="E2E8F0")
    poly(s, [(0.7, 4.8), (0.7, 3.5), (3.4, 1.9), (5.6, 3.3), (7.6, 2.5), (10.2, 3.9), (12.63, 3.0), (12.63, 4.8)], fill="94A3B8")
    poly(s, [(0.7, 4.8), (0.7, 4.2), (2.8, 3.4), (5.0, 4.4), (7.4, 3.8), (12.63, 4.8)], fill="64748B")
    return dict(t=(0.7, 5.0, 8.6, 1.5), tag=(9.7, 5.2, 2.9), sub=(0.7, 6.55, 8.6), tm=32)


def fr_photo(s, P, pg):
    box(s, R, 0.7, 0.62, 0.09, 0.85, fill=P.accent)
    box(s, R, 11.0, 0.5, 1.63, 1.1, fill="CBD5E1")
    poly(s, [(11.0, 1.6), (11.0, 1.2), (11.6, 0.8), (12.0, 1.1), (12.3, 0.95), (12.63, 1.25), (12.63, 1.6)], fill="94A3B8")
    return dict(t=(1.0, 0.55, 9.7, 1.05), area=(0.7, 1.95, 11.93, 4.8))


reg("photoframe", "Foto ramka", ["uni"],
    dict(bg="FFFFFF", ink="0F172A", text="475569", accent="0EA5E9", accent2="94A3B8", card="F1F5F9", muted="94A3B8",
         cycle=["0EA5E9", "6366F1"]),
    cv_photo, fr_photo, ck=dict(rad=0.1, line="E2E8F0"), layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 28 Burchak
def cv_corner(s, P):
    poly(s, [(0, 0), (5.2, 0), (0, 4.4)], fill=P.accent)
    poly(s, [(SW, SH), (8.6, SH), (SW, 3.6)], fill=P.accent2)
    poly(s, [(0, 0), (2.4, 0), (0, 2.0)], fill=mix(P.accent, "000000", 0.25))
    return dict(t=(3.3, 2.3, 7.9, 2.7), tag=(3.3, 1.9, 6), sub=(3.3, 5.25, 7.9), tm=46)


def fr_corner(s, P, pg):
    poly(s, [(0, 0), (2.2, 0), (0, 1.9)], fill=P.accent)
    poly(s, [(SW, SH), (11.4, SH), (SW, 5.8)], fill=P.accent2)
    return dict(t=(1.9, 0.55, 10.5, 1.05), area=(0.9, 2.0, 11.53, 4.7), pn=(9.7, 7.02, 1.3, "r", P.muted))


reg("corner", "Burchak", ["uni"],
    dict(bg="FFFFFF", ink="0F172A", text="475569", accent="1D4ED8", accent2="60A5FA", card="F1F5FF", muted="94A3B8",
         cycle=["1D4ED8", "60A5FA"]),
    cv_corner, fr_corner, ck=dict(rad=0.1, line="DBEAFE"), nk="square", layouts=["rows", "grid", "cols", "hero", "timeline"])


# ---------------------------------------------------------------- 29 Duoton
def cv_duo(s, P):
    box(s, R, 0, 0, SW / 2, SH, fill=P.accent)
    box(s, R, SW / 2, 0, SW / 2, SH, fill=P.accent2)
    box(s, O, 5.2, 4.75, 2.95, 2.95, fill="FFE3EC")
    return dict(t=(1.0, 1.25, 11.33, 2.5), al="c", tag=(1.0, 0.7, 11.33), sub=(1.0, 3.95, 11.33),
                ink="FFFFFF", tg="FFFFFF", tx="FFFFFF", tm=50)


def fr_duo(s, P, pg):
    box(s, R, 0, 0, SW / 2, 0.22, fill=P.accent)
    box(s, R, SW / 2, 0, SW / 2, 0.22, fill=P.accent2)
    box(s, O, 12.1, 0.55, 0.55, 0.55, fill=P.accent2)
    return dict(t=(0.9, 0.6, 10.9, 1.05), area=(0.9, 2.0, 11.53, 4.75))


reg("duotone", "Duoton", ["trend"],
    dict(bg="FFFFFF", ink="2B2D8F", text="3A3B6B", accent="2B2D8F", accent2="FF5C8A", card="FFF1F6", muted="B5B6D6",
         cycle=["2B2D8F", "FF5C8A"]),
    cv_duo, fr_duo, ck=dict(rad=0.12, line="FFD1E0"), layouts=["rows", "cols", "grid", "hero", "timeline"])


# ---------------------------------------------------------------- 30 Qirollik
def cv_royal(s, P):
    ring(s, 6.667, -1.4, 3.7, P.accent, 1)
    ring(s, 6.667, -1.4, 3.2, mix(P.accent, P.bg, 0.5), 0.5)
    ring(s, 6.667, -1.4, 2.7, mix(P.accent, P.bg, 0.7), 0.5)
    box(s, DIA, 6.39, 2.05, 0.55, 0.8, fill=P.accent)
    return dict(t=(1.5, 3.2, 10.33, 2.0), al="c", tag=(1.5, 2.95, 10.33), sub=(1.5, 5.5, 10.33), tm=46)


def fr_royal(s, P, pg):
    box(s, DIA, 6.52, 1.75, 0.3, 0.42, fill=P.accent)
    box(s, R, 4.4, 1.95, 1.9, 0.012, fill=P.accent)
    box(s, R, 7.04, 1.95, 1.9, 0.012, fill=P.accent)
    return dict(t=(1.0, 0.6, 11.33, 1.0), al="c", tm=30, area=(1.0, 2.4, 11.33, 4.2), pn=(6.0, 6.85, 1.3, "c", P.muted))


reg("royal", "Qirollik", ["uni"],
    dict(bg="2A0E4A", ink="F3E3B3", text="D2C2E6", accent="D4A24C", accent2="8C5BD0", card="3A1A62", muted="8A6AB4",
         head="Cambria", tb=False),
    cv_royal, fr_royal, ck=dict(rad=0.1, line="5A3A8C"), nk="plain", layouts=["rows", "cols", "grid", "hero", "timeline"])


# ================================================================ YIG'ISH
def has_style(sid):
    return sid in STYLES


def _slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _cover(prs, sid, topic, subtitle, tag):
    S = STYLES[sid]
    P = S["pal"]
    s = _slide(prs)
    bg(s, P.bg)
    c = S["cover"](s, P)
    al = c.get("al", "l")
    x, y, w, h = c["t"]
    size = fit_title(topic, w, h, c.get("tm", 44), 24)
    text(s, x, y, w, h, topic, size, c.get("ink", P.ink), P.head, bold=P.tb is not False,
         italic=bool(P.ti), al=al, anchor="m", ls=1.0)
    if c.get("tag") and tag:
        tx, ty, tw = c["tag"]
        text(s, tx, ty, tw, 0.4, tag, 13, c.get("tg", P.accent), P.body, bold=True, al=al, sp=3)
    if c.get("sub") and subtitle:
        sx, sy, sw = c["sub"]
        text(s, sx, sy, sw, 0.9, subtitle, 17, c.get("tx", P.text), P.body, al=al)
    return s


def _content(prs, sid, data, page, total):
    S = STYLES[sid]
    P = S["pal"]
    s = _slide(prs)
    bg(s, P.bg)
    f = S["frame"](s, P, page)
    x, y, w, h = f["t"]
    title = str(data.get("title", "")).strip()
    size = fit_title(title, w, h, f.get("tm", 32), 20)
    text(s, x, y, w, h, title, size, f.get("ink", P.ink), P.head, bold=P.tb is not False,
         italic=bool(P.ti), al=f.get("al", "l"), anchor="m", ls=1.0)
    items = [str(b).strip() for b in (data.get("content") or []) if str(b).strip()][:6]
    fn = pick_layout(P.layouts[(page - 2) % len(P.layouts)], len(items))
    if fn:
        fn(s, P, items, f["area"])
    px, py, pw, pal_, pc = f.get("pn", (SW - 1.9, 7.02, 1.3, "r", P.muted))
    text(s, px, py, pw, 0.3, "%d / %d" % (page, total), 11, pc, P.body, al=pal_)
    return s


def _closing(prs, sid, line, tag):
    """Oxirgi slayd: muqova dizayni, markazda qisqa matn."""
    S = STYLES[sid]
    P = S["pal"]
    s = _slide(prs)
    bg(s, P.bg)
    c = S["cover"](s, P)
    al = c.get("al", "l")
    x, y, w, h = c["t"]
    text(s, x, y, w * 0.9, h, line, fit_title(line, w * 0.9, h, 54, 28), c.get("ink", P.ink), P.head,
         bold=P.tb is not False, italic=bool(P.ti), al=al, anchor="m", ls=1.0)
    return s


def build_presentation(prs, topic, slides, style_id, subtitle="", tag="", closing=None):
    """prs - python-pptx Presentation; slides - [{"title":..., "content":[...]}, ...]."""
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    total = len(slides) + 1 + (1 if closing else 0)
    _cover(prs, style_id, topic, subtitle, tag)
    for i, d in enumerate(slides):
        _content(prs, style_id, d, i + 2, total)
    if closing:
        _closing(prs, style_id, closing, tag)
    return prs
