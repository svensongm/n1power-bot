"""Pictures (PNG for Telegram) and the monthly Excel file for the N1 POWER bot.
Layout follows the desk's Excel price sheet: banner, header, state row, hours, Avg and
Peak rows, green-yellow-red colour scale, decimal comma."""
from __future__ import annotations

import io
import math
from decimal import Decimal, ROUND_HALF_UP

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt                      # noqa: E402
from matplotlib.patches import Rectangle             # noqa: E402

GREEN, YELLOW, RED = (99, 190, 123), (255, 235, 132), (248, 105, 107)    # Excel 3-colour scale
BANNER, HEAD, OKFILL, WAITFILL, NAFILL = "#FFF2CC", "#DDEBF7", "#C6EFCE", "#FFEB9C", "#EDEDED"
GRID, GREYTXT, DARK = "#BFBFBF", "#808080", "#1F4E79"
MAX_SIDE = 2500                                      # Telegram keeps photos up to 2560 px sharp
DAYS = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]


def fmt(v, nd: int = 2) -> str:
    """Excel-style rounding (half up) with a decimal comma: 185.005 -> '185,01'."""
    if v is None:
        return ""
    q = Decimal(repr(float(v))).quantize(Decimal(1).scaleb(-nd), rounding=ROUND_HALF_UP)
    if q == 0:
        q = abs(q)
    return f"{q:.{nd}f}".replace(".", ",")


def signed(v, nd: int = 2) -> str:
    return "" if v is None else ("+" if v > 0 and fmt(v, nd) != "0,00" else "") + fmt(v, nd)


def _percentile(vals, p):
    s = sorted(vals)
    k = (len(s) - 1) * p
    f, c = math.floor(k), math.ceil(k)
    return s[f] + (s[c] - s[f]) * (k - f)


class Scale:
    """Green (low) - yellow (middle) - red (high) like Excel; mid=None -> median, mid=0 for spreads."""

    def __init__(self, values, mid=None):
        vals = [v for v in values if v is not None]
        self.ok = bool(vals)
        if self.ok:
            self.lo, self.hi = min(vals), max(vals)
            self.mid = _percentile(vals, 0.5) if mid is None else mid

    def color(self, v):
        if v is None or not self.ok:
            return None
        if v <= self.mid:
            t = 0.0 if self.mid == self.lo else (v - self.lo) / (self.mid - self.lo)
            a, b = GREEN, YELLOW
        else:
            t = 1.0 if self.hi == self.mid else (v - self.mid) / (self.hi - self.mid)
            a, b = YELLOW, RED
        t = max(0.0, min(1.0, t))
        return "#%02X%02X%02X" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


class Canvas:
    def __init__(self, w: float, h: float):
        self.w, self.h = int(math.ceil(w)), int(math.ceil(h))
        self.fig = plt.figure(figsize=(self.w / 100, self.h / 100), dpi=100)
        self.ax = self.fig.add_axes([0, 0, 1, 1])
        self.ax.set_xlim(0, self.w)
        self.ax.set_ylim(self.h, 0)
        self.ax.axis("off")

    def box(self, x, y, w, h, fill="white", edge=GRID, lw=0.6):
        self.ax.add_patch(Rectangle((x, y), w, h, facecolor=fill or "white", edgecolor=edge, linewidth=lw))

    def text(self, x, y, s, size=8.5, bold=False, color="black", ha="center", va="center", italic=False):
        if s not in (None, ""):
            self.ax.text(x, y, s, ha=ha, va=va, fontsize=size, color=color,
                         fontweight="bold" if bold else "normal", fontstyle="italic" if italic else "normal")

    def cell(self, x, y, w, h, s="", fill="white", size=8.5, bold=False, color="black", ha="center", edge=GRID):
        self.box(x, y, w, h, fill, edge)
        tx = x + w / 2 if ha == "center" else (x + 4 if ha == "left" else x + w - 4)
        self.text(tx, y + h / 2, s, size, bold, color, ha)

    def png(self) -> bytes:
        scale = min(2.0, MAX_SIDE / max(self.w, self.h))
        buf = io.BytesIO()
        self.fig.savefig(buf, format="png", dpi=100 * scale)
        plt.close(self.fig)
        return buf.getvalue()


# --------------------------------------------------------------------------- hourly table
RB, RHD, RH, C0 = 22, 22, 18, 46


def hourly_size(t: dict, width: float | None = None):
    n = len(t["headers"])
    cw = (width - C0) / n if width else (62 if n <= 26 else 58)
    return C0 + cw * n, RB + RHD + RH * (1 + len(t["rows"]) + 2) + 20, cw


def draw_hourly(cv: Canvas, x0: float, y0: float, t: dict, width: float | None = None) -> float:
    W, H, cw = hourly_size(t, width)
    scale = Scale([v for row in t["rows"] for v in row])
    y = y0
    cv.cell(x0, y, C0, RB, t["date"], BANNER, 9, True)
    cv.cell(x0 + C0, y, cw * 2, RB, t["weekday"], "white", 9, True)
    cv.cell(x0 + C0 + cw * 2, y, W - C0 - cw * 2, RB, t["banner"], BANNER, 9, True, "#C65911")
    y += RB
    cv.cell(x0, y, C0, RHD, "Hour", HEAD, 9, True)
    for j, h in enumerate(t["headers"]):
        cv.cell(x0 + C0 + cw * j, y, cw, RHD, h, HEAD, 8.5 if len(h) <= 7 else 7.5, True)
    y += RHD
    cv.cell(x0, y, C0, RH, "state", HEAD, 8, True)
    for j, s in enumerate(t["states"]):
        cv.cell(x0 + C0 + cw * j, y, cw, RH, s, {"ACTUAL": OKFILL, "waiting": WAITFILL}.get(s, NAFILL), 7, s == "ACTUAL")
    y += RH
    for i, row in enumerate(t["rows"]):
        cv.cell(x0, y, C0, RH, str(t["hours"][i]), "white", 8.5)
        for j, v in enumerate(row):
            cv.cell(x0 + C0 + cw * j, y, cw, RH, fmt(v), scale.color(v) or "white", 8.5)
        y += RH
    for name, vals, fill in (("Avg", t["avg"], BANNER), ("Peak", t["peak"], "white")):
        cv.cell(x0, y, C0, RH, name, fill, 8.5, True)
        for j, v in enumerate(vals):
            cv.cell(x0 + C0 + cw * j, y, cw, RH, fmt(v), fill, 8.5, True)
        y += RH
    cv.text(x0 + 4, y + 10, t.get("footer", ""), 7, False, GREYTXT, "left")
    return H


def hourly_table_png(t: dict) -> bytes:
    W, H, _ = hourly_size(t)
    cv = Canvas(W, H)
    draw_hourly(cv, 0, 0, t)
    return cv.png()


# --------------------------------------------------------------------------- flags (drawn, no images needed)
from matplotlib.patches import Circle, FancyBboxPatch, Polygon  # noqa: E402

NAMES = {"HU": "Hungary", "HR": "Croatia", "SI": "Slovenia", "RS": "Serbia", "RO": "Romania", "BG": "Bulgaria",
         "GR": "Greece", "ME": "Montenegro", "MK": "North Macedonia", "IT": "Italy · PUN", "DE": "Germany",
         "AT": "Austria", "FR": "France", "ES": "Spain", "CH": "Switzerland", "UA": "Ukraine", "PL": "Poland",
         "CZ": "Czechia", "SK": "Slovakia", "AL": "Albania", "BA": "Bosnia and Herzegovina", "XK": "Kosovo",
         "TR": "Türkiye", "MD": "Moldova", "BE": "Belgium", "NL": "Netherlands", "PT": "Portugal"}


def _star(cx, cy, r, n=5):
    pts = []
    for k in range(2 * n):
        rr = r if k % 2 == 0 else r * 0.45
        a = math.pi / 2 + k * math.pi / n
        pts.append((cx + rr * math.cos(a), cy - rr * math.sin(a)))
    return pts


def draw_flag(ax, code: str, x: float, y: float, w: float, h: float, z: float = 6):
    """Simplified national flag in the box with top-left (x, y), size w x h (y grows downwards)."""
    c = (code or "").upper().replace("-", "_")[:2]

    def rect(xx, yy, ww, hh, col):
        ax.add_patch(Rectangle((xx, yy), ww, hh, facecolor=col, edgecolor="none", zorder=z))

    def poly(pts, col, zz=None):
        ax.add_patch(Polygon(pts, closed=True, facecolor=col, edgecolor="none", zorder=zz or z + 0.1))

    def circ(cx, cy, r, col, edge="none", lw=0.0):
        ax.add_patch(Circle((cx, cy), r, facecolor=col, edgecolor=edge, linewidth=lw, zorder=z + 0.2))

    def hs(cols, weights=None):
        weights = weights or [1] * len(cols)
        yy = y
        for col, wt in zip(cols, weights):
            hh = h * wt / sum(weights)
            rect(x, yy, w, hh, col)
            yy += hh

    def vs(cols, weights=None):
        weights = weights or [1] * len(cols)
        xx = x
        for col, wt in zip(cols, weights):
            ww = w * wt / sum(weights)
            rect(xx, y, ww, h, col)
            xx += ww

    def shield(sx, sy, sw, sh, col, edge="none", lw=0.0):
        pts = [(sx, sy), (sx + sw, sy), (sx + sw, sy + sh * 0.62), (sx + sw / 2, sy + sh), (sx, sy + sh * 0.62)]
        ax.add_patch(Polygon(pts, closed=True, facecolor=col, edgecolor=edge, linewidth=lw, zorder=z + 0.3))

    cx, cy = x + w / 2, y + h / 2
    if c == "HU":
        hs(["#CD2A3E", "#FFFFFF", "#436F4D"])
    elif c == "HR":
        hs(["#FF0000", "#FFFFFF", "#171796"])
        s, n = h * 0.5, 4
        x0, y0 = cx - s / 2, y + h * 0.2
        for i in range(n):
            for j in range(n):
                ax.add_patch(Rectangle((x0 + j * s / n, y0 + i * s / n), s / n, s / n, zorder=z + 0.3,
                                       facecolor="#FF0000" if (i + j) % 2 == 0 else "#FFFFFF", edgecolor="none"))
    elif c == "SI":
        hs(["#FFFFFF", "#005DA4", "#ED1C24"])
        sx, sy, sw, sh = x + w * 0.2, y + h * 0.13, w * 0.2, h * 0.46
        shield(sx, sy, sw, sh, "#005DA4", "#ED1C24", 0.5)
        poly([(sx + sw * 0.15, sy + sh * 0.55), (sx + sw * 0.5, sy + sh * 0.2), (sx + sw * 0.85, sy + sh * 0.55)],
             "#FFFFFF", z + 0.4)
    elif c == "RS":
        hs(["#C6363C", "#0C4076", "#FFFFFF"])
        sx, sy, sw, sh = x + w * 0.22, y + h * 0.18, w * 0.18, h * 0.55
        shield(sx, sy, sw, sh, "#C6363C", "#EDB92E", 0.6)
        rect(sx + sw * 0.42, sy + sh * 0.12, sw * 0.16, sh * 0.6, "#FFFFFF")
        ax.patches[-1].set_zorder(z + 0.4)
        rect(sx + sw * 0.15, sy + sh * 0.3, sw * 0.7, sh * 0.14, "#FFFFFF")
        ax.patches[-1].set_zorder(z + 0.4)
    elif c == "RO":
        vs(["#002B7F", "#FCD116", "#CE1126"])
    elif c == "BG":
        hs(["#FFFFFF", "#00966E", "#D62612"])
    elif c == "GR":
        hs(["#0D5EAF" if i % 2 == 0 else "#FFFFFF" for i in range(9)])
        side, arm = h * 5 / 9, h / 9
        rect(x, y, side, side, "#0D5EAF")
        rect(x + side / 2 - arm / 2, y, arm, side, "#FFFFFF")
        rect(x, y + side / 2 - arm / 2, side, arm, "#FFFFFF")
    elif c == "ME":
        rect(x, y, w, h, "#D4AF3A")
        b = h * 0.1
        rect(x + b, y + b, w - 2 * b, h - 2 * b, "#C40308")
        poly(_star(cx, cy, h * 0.24, 4), "#D4AF3A")
    elif c == "MK":
        rect(x, y, w, h, "#D20000")
        for tx, ty in [(x, y), (cx, y), (x + w, y), (x + w, cy), (x + w, y + h), (cx, y + h), (x, y + h), (x, cy)]:
            dx, dy = tx - cx, ty - cy
            ln = math.hypot(dx, dy) or 1
            px, py = -dy / ln * h * 0.12, dx / ln * h * 0.12
            poly([(cx, cy), (tx + px, ty + py), (tx - px, ty - py)], "#FFE600")
        circ(cx, cy, h * 0.2, "#FFE600", "#D20000", 0.6)
    elif c == "IT":
        vs(["#009246", "#FFFFFF", "#CE2B37"])
    elif c == "DE":
        hs(["#000000", "#DD0000", "#FFCE00"])
    elif c == "AT":
        hs(["#ED2939", "#FFFFFF", "#ED2939"])
    elif c == "FR":
        vs(["#0055A4", "#FFFFFF", "#EF4135"])
    elif c == "ES":
        hs(["#AA151B", "#F1BF00", "#AA151B"], [1, 2, 1])
        rect(x + w * 0.22, y + h * 0.33, w * 0.12, h * 0.34, "#AD1519")
        ax.patches[-1].set_zorder(z + 0.3)
    elif c == "CH":
        rect(x, y, w, h, "#DA291C")
        t, ln = h * 0.2, h * 0.62
        rect(cx - t / 2, cy - ln / 2, t, ln, "#FFFFFF")
        ax.patches[-1].set_zorder(z + 0.3)
        rect(cx - ln / 2, cy - t / 2, ln, t, "#FFFFFF")
        ax.patches[-1].set_zorder(z + 0.3)
    elif c == "UA":
        hs(["#0057B7", "#FFD700"])
    elif c == "PL":
        hs(["#FFFFFF", "#DC143C"])
    elif c == "CZ":
        hs(["#FFFFFF", "#D7141A"])
        poly([(x, y), (x + w * 0.5, cy), (x, y + h)], "#11457E")
    elif c == "SK":
        hs(["#FFFFFF", "#0B4EA2", "#EE1C25"])
        shield(x + w * 0.2, y + h * 0.2, w * 0.22, h * 0.6, "#EE1C25", "#FFFFFF", 0.6)
    elif c == "BE":
        vs(["#000000", "#FDDA24", "#EF3340"])
    elif c == "NL":
        hs(["#AE1C28", "#FFFFFF", "#21468B"])
    elif c == "AL":
        rect(x, y, w, h, "#E41E20")
        poly([(cx, cy - h * .32), (cx + w * .1, cy - h * .18), (cx + w * .24, cy - h * .3), (cx + w * .2, cy + h * .02),
              (cx + w * .08, cy + h * .1), (cx + w * .05, cy + h * .34), (cx, cy + h * .24), (cx - w * .05, cy + h * .34),
              (cx - w * .08, cy + h * .1), (cx - w * .2, cy + h * .02), (cx - w * .24, cy - h * .3),
              (cx - w * .1, cy - h * .18)], "#111111")
    elif c == "BA":
        rect(x, y, w, h, "#002395")
        poly([(x + w * 0.27, y), (x + w * 0.77, y), (x + w * 0.77, y + h)], "#FECB00")
        for k in range(7):
            tt = (k + 0.5) / 7
            circ(x + w * (0.2 + 0.47 * tt), y + h * tt, h * 0.045, "#FFFFFF")
    elif c == "XK":
        rect(x, y, w, h, "#244AA5")
        poly([(cx - w * .15, cy), (cx - w * .05, cy - h * .16), (cx + w * .12, cy - h * .1), (cx + w * .16, cy + h * .1),
              (cx + w * .02, cy + h * .24), (cx - w * .12, cy + h * .14)], "#D0A650")
        for k in range(6):
            circ(cx - w * 0.2 + k * w * 0.08, y + h * 0.18 - math.sin(math.pi * (k + 0.5) / 6) * h * 0.06,
                 h * 0.035, "#FFFFFF")
    elif c == "TR":
        rect(x, y, w, h, "#E30A17")
        circ(x + w * 0.38, cy, h * 0.25, "#FFFFFF")
        circ(x + w * 0.43, cy, h * 0.2, "#E30A17")
        poly(_star(x + w * 0.6, cy, h * 0.11), "#FFFFFF", z + 0.4)
    elif c == "MD":
        vs(["#0046AE", "#FFD200", "#CC092F"])
        circ(cx, cy, h * 0.13, "#B07E09")
    elif c == "PT":
        vs(["#006600", "#FF0000"], [2, 3])
        circ(x + w * 0.4, cy, h * 0.2, "#FFE900")
    else:
        rect(x, y, w, h, "#CBD5E1")
        ax.text(cx, cy, c, ha="center", va="center", fontsize=5.5, color="#334155", zorder=z + 1)
    ax.add_patch(Rectangle((x, y), w, h, facecolor="none", edgecolor="#94A3B8", linewidth=0.5, zorder=z + 1))


# --------------------------------------------------------------------------- stylish calendar cards
INK, MUTED, LINE = "#0F172A", "#64748B", "#E2E8F0"
HEADBAND, CHIP, ACCENT, PAGE = "#1E293B", "#334155", "#2563EB", "#EEF2F7"
SPREAD_BG, SPREAD_INK = "#DBEAFE", "#1D4ED8"
CW, HEADH, DOWH, WH, FOOTH = 74, 38, 20, 60, 28
DOW = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN", "WEEK"]


def _soft(hexcol, k=0.22):
    r, g, b = (int(hexcol[i:i + 2], 16) for i in (1, 3, 5))
    return "#%02X%02X%02X" % (round(r + (255 - r) * k), round(g + (255 - g) * k), round(b + (255 - b) * k))


def _rbox(ax, x, y, w, h, fc, ec="none", lw=0.0, r=6, z=2, alpha=1.0):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}", facecolor=fc,
                                edgecolor=ec, linewidth=lw, zorder=z, alpha=alpha))


def _card(cv: Canvas, x0, y0, w, h, code, title, chip=None):
    """Card with shadow, dark header band with flag and title, optional chip on the right."""
    ax = cv.ax
    _rbox(ax, x0 + 2, y0 + 3, w, h, "#0F172A", r=12, z=0.5, alpha=0.10)
    _rbox(ax, x0, y0, w, h, "#FFFFFF", "#CBD5E1", 0.6, r=12, z=1)
    _rbox(ax, x0, y0, w, HEADH, HEADBAND, r=12, z=1.5)
    ax.add_patch(Rectangle((x0, y0 + HEADH / 2), w, HEADH / 2, facecolor=HEADBAND, edgecolor="none", zorder=1.5))
    if code:
        draw_flag(ax, code, x0 + 14, y0 + 9, 30, 20)
    cv.text(x0 + (54 if code else 14), y0 + HEADH / 2 + 1, title, 10.5, True, "#FFFFFF", "left")
    if chip:
        cw = 9 + 5.4 * len(chip)
        _rbox(ax, x0 + w - 12 - cw, y0 + 9, cw, 20, CHIP, r=10, z=2)
        cv.text(x0 + w - 12 - cw / 2, y0 + 19.5, chip, 7.6, True, "#F8FAFC")


def _calendar_block(cv: Canvas, x0, y0, b: dict, n_weeks: int):
    """Country card: each day B (base), P (peak 08-20), MTD (month average of base to that day);
    week column Mon-Sun B and P and the spread to HU."""
    ax = cv.ax
    w = CW * 8
    h = HEADH + DOWH + WH * n_weeks + FOOTH + 4
    code = b.get("code", "")
    chip = f"MTD  B {fmt(b['mtd'])}" if b.get("mtd") is not None else None
    _card(cv, x0, y0, w, h, code, f"{code} · {NAMES.get(code, code)}", chip)
    y = y0 + HEADH
    for j, dname in enumerate(DOW):
        col = ACCENT if j == 7 else ("#94A3B8" if j >= 5 else MUTED)
        cv.text(x0 + CW * j + CW / 2, y + DOWH / 2 + 1, dname, 7, True, col)
    y += DOWH
    scale = Scale([c["base"] for wk in b["weeks"] for c in wk if c["in_month"] and c["base"] is not None])
    target = None
    for k in range(n_weeks):
        wk = b["weeks"][k] if k < len(b["weeks"]) else None
        for j in range(7):
            x = x0 + CW * j
            c = wk[j] if wk else None
            if c is None:
                continue
            known = c["base"] is not None
            if c["in_month"] and known:
                _rbox(ax, x + 2.5, y + 2.5, CW - 5, WH - 5, _soft(scale.color(c["base"]) or "#FFFFFF"), r=7)
            elif c["in_month"]:
                _rbox(ax, x + 2.5, y + 2.5, CW - 5, WH - 5, "#FFFFFF", LINE, 0.7, r=7)
            else:
                _rbox(ax, x + 2.5, y + 2.5, CW - 5, WH - 5, "#F8FAFC", r=7)
            day_col = "#334155" if c["in_month"] else "#CBD5E1"
            cv.text(x + 8, y + 11, str(c["day"]), 6.6, c["in_month"], day_col, "left")
            if known:
                ink = INK if c["in_month"] else "#A3AEC0"
                cv.text(x + CW / 2, y + 25, "B " + fmt(c["base"]), 8.3, c["in_month"], ink)
                if c["peak"] is not None:
                    cv.text(x + CW / 2, y + 38, "P " + fmt(c["peak"]), 7, False, ink)
                if c["in_month"] and c["mtd"] is not None:
                    cv.text(x + CW / 2, y + 50, "MTD " + fmt(c["mtd"]), 6, False, "#475569")
            if c.get("target"):
                target = (x, y)
        if wk is not None:
            xw = x0 + CW * 7
            _rbox(ax, xw + 2.5, y + 2.5, CW - 5, WH - 5, "#F1F5F9", r=7)
            if b["week_base"][k] is not None:
                cv.text(xw + CW / 2, y + 15, "B " + fmt(b["week_base"][k]), 8.3, True, INK)
                if b["week_peak"][k] is not None:
                    cv.text(xw + CW / 2, y + 28, "P " + fmt(b["week_peak"][k]), 7, False, INK)
                ws = (b.get("week_spread") or [None] * len(b["weeks"]))[k]
                if ws is not None:
                    _rbox(ax, xw + 5, y + 34, CW - 10, 13, SPREAD_BG, r=6, z=3)
                    cv.text(xw + CW / 2, y + 40.5, f"{b['spread_name']} {signed(ws)}", 5.9, True, SPREAD_INK)
                if b["week_partial"][k]:
                    cv.text(xw + CW / 2, y + 53, "so far", 5.6, False, "#94A3B8", italic=True)
        y += WH
    if target:
        _rbox(ax, target[0] + 1.5, target[1] + 1.5, CW - 3, WH - 3, "none", ACCENT, 2.4, r=8, z=5)
    # footer
    fy = y0 + h - FOOTH - 2
    ax.add_patch(Rectangle((x0 + 1, fy), w - 2, 0.8, facecolor=LINE, edgecolor="none", zorder=2))
    parts = f"{b.get('month', '')} MTD   B {fmt(b.get('mtd')) or '-'}   P {fmt(b.get('mtd_p')) or '-'}"
    if b.get("dod") is not None:
        parts += f"   ·   DoD {signed(b['dod'])}"
    cv.text(x0 + 14, fy + FOOTH / 2 + 1, parts, 7.6, True, INK, "left")
    if b.get("spread_mtd") is not None:
        txt = f"{b['spread_name']} MTD {signed(b['spread_mtd'])}"
        cw = 10 + 5.2 * len(txt)
        _rbox(ax, x0 + w - 12 - cw, fy + 6, cw, FOOTH - 12, SPREAD_BG, r=8, z=3)
        cv.text(x0 + w - 12 - cw / 2, fy + FOOTH / 2 + 0.5, txt, 7, True, SPREAD_INK)


def _recap_block(cv: Canvas, x0, y0, b: dict, height):
    """Week recap card: one row per spread with both flags, Mon-Sun week columns and the month."""
    ax = cv.ax
    w = CW * 8
    _card(cv, x0, y0, w, height, b.get("code", ""), b["title"], "EUR/MWh")
    y, hh = y0 + HEADH + 4, 30
    lw_ = 112
    cw = (w - lw_ - 14) / max(1, len(b["columns"]))
    cv.text(x0 + 14, y + hh / 2, "Spread", 7.6, True, MUTED, "left")
    for j, (l1, l2) in enumerate(b["columns"]):
        cx = x0 + lw_ + cw * j + cw / 2
        last = j == len(b["columns"]) - 1
        cv.text(cx, y + 10, l1, 8, True, ACCENT if last else INK)
        cv.text(cx, y + 22, l2, 6.4, False, MUTED)
    y += hh
    rows = b["rows"]
    rh = min(30, (height - HEADH - hh - FOOTH - 10) / max(1, len(rows)))
    scale = Scale([v for r in rows for v in r[1]], mid=0)
    for i, r in enumerate(rows):
        name, vals = r[0], r[1]
        codes = r[2] if len(r) > 2 else None
        if i % 2 == 0:
            ax.add_patch(Rectangle((x0 + 6, y), w - 12, rh, facecolor="#F8FAFC", edgecolor="none", zorder=1.6))
        if codes:
            fh = min(11, rh * 0.42)
            draw_flag(ax, codes[0], x0 + 14, y + rh / 2 - fh / 2, fh * 1.5, fh)
            draw_flag(ax, codes[1], x0 + 16 + fh * 1.5, y + rh / 2 - fh / 2, fh * 1.5, fh)
        cv.text(x0 + 22 + fh * 3 if codes else x0 + 14, y + rh / 2 + 0.5, name, 8.4, True, INK, "left")
        for j, v in enumerate(vals):
            cx = x0 + lw_ + cw * j
            if v is not None:
                _rbox(ax, cx + 4, y + 3, cw - 8, rh - 6, _soft(scale.color(v) or "#FFFFFF", 0.12), r=7, z=2)
            cv.text(cx + cw / 2, y + rh / 2 + 0.5, signed(v), 8.2, j == len(vals) - 1, INK)
        y += rh
    fy = y0 + height - FOOTH - 2
    ax.add_patch(Rectangle((x0 + 1, fy), w - 2, 0.8, facecolor=LINE, edgecolor="none", zorder=2))
    cv.text(x0 + 14, fy + FOOTH / 2 + 1, b["footer"], 6.8, False, MUTED, "left")


def calendars_png(title: str, blocks: list[dict], ncols: int = 3, legend: str = "") -> bytes:
    """One picture of calendar cards (and the week recap), ncols per row."""
    cals = [b for b in blocks if b["kind"] != "recap"]
    n_weeks = max((len(b["weeks"]) for b in cals), default=5)
    bw, gap, top = CW * 8, 20, 64 if legend else 44
    bh = HEADH + DOWH + WH * n_weeks + FOOTH + 4
    rows = math.ceil(len(blocks) / ncols)
    W = ncols * bw + (ncols + 1) * gap
    H = top + rows * (bh + gap) + gap / 2
    cv = Canvas(W, H)
    cv.fig.patch.set_facecolor(PAGE)
    cv.text(gap, 24, title, 14, True, INK, "left")
    if legend:
        cv.text(gap, 46, legend, 8.4, False, MUTED, "left")
    for i, b in enumerate(blocks):
        r, c = divmod(i, ncols)
        x, y = gap + c * (bw + gap), top + r * (bh + gap)
        if b["kind"] == "recap":
            _recap_block(cv, x, y, b, bh)
        else:
            _calendar_block(cv, x, y, b, n_weeks)
    return cv.png()


# --------------------------------------------------------------------------- Excel (monthly file, 2 sheets)
def _x():
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    thin = Side(style="thin", color="BFBFBF")
    return {"rule": ColorScaleRule, "Al": Alignment, "Font": Font, "gcl": get_column_letter,
            "fill": lambda h: PatternFill("solid", fgColor=h.lstrip("#")),
            "border": Border(left=thin, right=thin, top=thin, bottom=thin),
            "thick": Border(**{k: Side(style="medium", color="1F4E79") for k in ("left", "right", "top", "bottom")}),
            "center": Alignment(horizontal="center", vertical="center")}


def xl_hourly(ws, top: int, left: int, t: dict) -> int:
    """One hourly table (banner, header, state, hours, Avg, Peak) at (top, left); returns rows used."""
    X = _x()
    n = len(t["headers"])
    ws.cell(top, left, t["date"]).font = X["Font"](bold=True)
    ws.cell(top, left).fill = X["fill"](BANNER)
    ws.cell(top, left + 1, t["weekday"]).font = X["Font"](bold=True)
    ws.merge_cells(start_row=top, start_column=left + 2, end_row=top, end_column=left + n)
    bc = ws.cell(top, left + 2, t["banner"])
    bc.font, bc.fill, bc.alignment = X["Font"](bold=True, color="C65911"), X["fill"](BANNER), X["center"]
    for j, h in enumerate(["Hour"] + t["headers"]):
        c = ws.cell(top + 1, left + j, h)
        c.font, c.fill, c.alignment, c.border = X["Font"](bold=True), X["fill"](HEAD), X["center"], X["border"]
    ws.cell(top + 2, left, "state").font = X["Font"](bold=True)
    for j, s in enumerate(t["states"], start=1):
        c = ws.cell(top + 2, left + j, s)
        c.fill = X["fill"]({"ACTUAL": OKFILL, "waiting": WAITFILL}.get(s, NAFILL))
        c.font, c.alignment, c.border = X["Font"](bold=s == "ACTUAL", size=9), X["center"], X["border"]
    r0 = top + 3
    for i, row in enumerate(t["rows"]):
        ws.cell(r0 + i, left, t["hours"][i]).alignment = X["center"]
        for j, v in enumerate(row, start=1):
            c = ws.cell(r0 + i, left + j, v)
            c.number_format, c.alignment, c.border = "0.00", X["center"], X["border"]
    r1 = r0 + len(t["rows"]) - 1
    ws.conditional_formatting.add(
        f"{X['gcl'](left + 1)}{r0}:{X['gcl'](left + n)}{r1}",
        X["rule"](start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50,
                  mid_color="FFEB84", end_type="max", end_color="F8696B"))
    for k, (label, vals, fill) in enumerate((("Avg", t["avg"], BANNER), ("Peak", t["peak"], None))):
        r = r1 + 1 + k
        ws.cell(r, left, label).font = X["Font"](bold=True)
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, left + j, v)
            c.number_format, c.font, c.alignment, c.border = "0.00", X["Font"](bold=True), X["center"], X["border"]
            if fill:
                c.fill = X["fill"](fill)
    ws.cell(r1 + 3, left, t.get("footer", "")).font = X["Font"](italic=True, color="808080", size=9)
    return r1 + 4 - top


def xl_calendars(ws, top: int, blocks: list[dict], ncols: int = 3) -> int:
    """Calendars (B, P, MTD per day; week B, P and spread) and the week recap, ncols per row; returns rows used."""
    X = _x()
    cals = [b for b in blocks if b["kind"] != "recap"]
    n_weeks = max((len(b["weeks"]) for b in cals), default=5)
    bh = 2 + 4 * n_weeks + 2
    for i, b in enumerate(blocks):
        r, c = divmod(i, ncols)
        t0, l0 = top + r * bh, 1 + c * 9
        ws.merge_cells(start_row=t0, start_column=l0, end_row=t0, end_column=l0 + 7)
        tc = ws.cell(t0, l0, b["title"])
        tc.font, tc.fill, tc.alignment = X["Font"](bold=True), X["fill"](HEAD), X["center"]
        if b["kind"] == "recap":
            for j, h in enumerate(["Spread"] + [f"{a} {z}" for a, z in b["columns"]]):
                hc = ws.cell(t0 + 1, l0 + j, h)
                hc.font, hc.fill, hc.alignment, hc.border = X["Font"](bold=True, size=9), X["fill"]("#F2F2F2"), X["center"], X["border"]
            sc = Scale([v for row in b["rows"] for v in row[1]], mid=0)
            for k, row in enumerate(b["rows"]):
                name, vals = row[0], row[1]
                ws.cell(t0 + 2 + k, l0, name).font = X["Font"](bold=True)
                for j, v in enumerate(vals, start=1):
                    vc = ws.cell(t0 + 2 + k, l0 + j, v)
                    vc.number_format, vc.alignment, vc.border = "+0.00;-0.00;0.00", X["center"], X["border"]
                    vc.font = X["Font"](bold=j == len(vals))
                    if sc.color(v):
                        vc.fill = X["fill"](sc.color(v))
            ws.cell(t0 + 2 + len(b["rows"]), l0, b["footer"]).font = X["Font"](italic=True, size=8, color="808080")
            continue
        scale = Scale([x["base"] for wk in b["weeks"] for x in wk if x["in_month"] and x["base"] is not None])
        for j, d in enumerate(DAYS + ["Week"]):
            hc = ws.cell(t0 + 1, l0 + j, d)
            hc.font, hc.fill, hc.alignment, hc.border = X["Font"](bold=True), X["fill"]("#F2F2F2"), X["center"], X["border"]
        for k, wk in enumerate(b["weeks"]):
            rr = t0 + 2 + 4 * k
            for j, x in enumerate(wk):
                cells = [ws.cell(rr, l0 + j, x["day"]), ws.cell(rr + 1, l0 + j, x["base"]),
                         ws.cell(rr + 2, l0 + j, x["peak"]), ws.cell(rr + 3, l0 + j, x["mtd"] if x["in_month"] else None)]
                grey = "000000" if x["in_month"] else "9A9A9A"
                cells[0].font = X["Font"](size=8, color="808080" if x["in_month"] else "B0B0B0")
                cells[1].number_format, cells[1].font = '"B "0.00', X["Font"](bold=x["in_month"], color=grey)
                cells[2].number_format, cells[2].font = '"P "0.00', X["Font"](size=9, color=grey)
                cells[3].number_format, cells[3].font = '"MTD "0.00', X["Font"](size=8, italic=True, color="404040")
                col = scale.color(x["base"]) if (x["in_month"] and x["base"] is not None) else None
                for cc in cells:
                    cc.alignment = X["center"] if cc is not cells[0] else X["Al"](horizontal="left")
                    cc.fill = X["fill"](col or ("#F7F7F7" if not x["in_month"] else "#FFFFFF"))
                if x.get("target"):
                    cells[1].border = X["thick"]
            wcol = l0 + 7
            for q in range(4):
                ws.cell(rr + q, wcol).fill = X["fill"]("#F2F2F2")
            if b["week_base"][k] is not None:
                if b["week_partial"][k]:
                    ws.cell(rr, wcol, "so far").font = X["Font"](size=8, italic=True, color="808080")
                wb_ = ws.cell(rr + 1, wcol, b["week_base"][k])
                wb_.number_format, wb_.font, wb_.alignment = '"B "0.00', X["Font"](bold=True), X["center"]
                wp = ws.cell(rr + 2, wcol, b["week_peak"][k])
                wp.number_format, wp.font, wp.alignment = '"P "0.00', X["Font"](size=9), X["center"]
                ws_ = (b.get("week_spread") or [None] * len(b["weeks"]))[k]
                if ws_ is not None:
                    sc = ws.cell(rr + 3, wcol, f"{b['spread_name']} {signed(ws_)}")
                    sc.font, sc.alignment = X["Font"](size=8, color="1F4E79"), X["center"]
        ws.cell(t0 + 2 + 4 * len(b["weeks"]), l0, b["footer"]).font = X["Font"](bold=True)
    return math.ceil(len(blocks) / ncols) * bh


def xl_matrix(ws, top: int, left: int, title: str, corner: str, cols: list[str], rows: list[tuple]) -> int:
    """Simple coloured table (e.g. daily average capacity price per border); returns rows used."""
    X = _x()
    ws.cell(top, left, title).font = X["Font"](bold=True, size=11)
    for j, h in enumerate([corner] + cols):
        c = ws.cell(top + 1, left + j, h)
        c.font, c.fill, c.alignment, c.border = X["Font"](bold=True), X["fill"](HEAD), X["center"], X["border"]
    scale = Scale([v for _, vals, _b in rows if not _b for v in vals])
    for i, (name, vals, bold) in enumerate(rows):
        r = top + 2 + i
        ws.cell(r, left, name).font = X["Font"](bold=True)
        for j, v in enumerate(vals, start=1):
            c = ws.cell(r, left + j, v)
            c.number_format, c.alignment, c.border = "0.00", X["center"], X["border"]
            c.font = X["Font"](bold=bold)
            col = BANNER if bold else scale.color(v)
            if col and v is not None:
                c.fill = X["fill"](col)
    return len(rows) + 3


def workbook_bytes(build) -> bytes:
    from openpyxl import Workbook
    wb = Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
