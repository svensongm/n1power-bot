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


# --------------------------------------------------------------------------- calendars
CW, TH, HH, WH, FH = 64, 24, 18, 56, 20


def _calendar_block(cv: Canvas, x0, y0, b: dict, n_weeks: int):
    """Day cell: B (base), P (peak 08-20), MTD (month average of base to that day).
    Week column: Mon-Sun B and P, the spread to HU (blue) when given."""
    w = CW * 8
    scale = Scale([c["base"] for wk in b["weeks"] for c in wk if c["in_month"] and c["base"] is not None])
    cv.cell(x0, y0, w, TH, b["title"], HEAD, 9.5, True)
    y = y0 + TH
    for j, d in enumerate(DAYS + ["Week"]):
        cv.cell(x0 + CW * j, y, CW, HH, d, "#F2F2F2", 8, True)
    y += HH
    target = None
    for k in range(n_weeks):
        wk = b["weeks"][k] if k < len(b["weeks"]) else None
        for j in range(7):
            x = x0 + CW * j
            if wk is None:
                cv.box(x, y, CW, WH, "white")
                continue
            c = wk[j]
            known = c["base"] is not None
            fill = (scale.color(c["base"]) if known else "white") if c["in_month"] else "#F7F7F7"
            cv.box(x, y, CW, WH, fill or "white")
            cv.text(x + 4, y + 7, str(c["day"]), 6.3, False, GREYTXT if c["in_month"] else "#B0B0B0", "left")
            if known:
                col = "black" if c["in_month"] else "#9A9A9A"
                cv.text(x + CW / 2, y + 20, "B " + fmt(c["base"]), 8.4, c["in_month"], col)
                if c["peak"] is not None:
                    cv.text(x + CW / 2, y + 33, "P " + fmt(c["peak"]), 7.4, False, col)
                if c["in_month"] and c["mtd"] is not None:
                    cv.text(x + CW / 2, y + 46, "MTD " + fmt(c["mtd"]), 6.4, False, "#404040")
            if c.get("target"):
                target = (x, y)
        xw = x0 + CW * 7
        cv.box(xw, y, CW, WH, "#F2F2F2" if wk is not None else "white")
        if wk is not None and b["week_base"][k] is not None:
            cv.text(xw + CW / 2, y + 12, "B " + fmt(b["week_base"][k]), 8.4, True)
            if b["week_peak"][k] is not None:
                cv.text(xw + CW / 2, y + 24, "P " + fmt(b["week_peak"][k]), 7.4)
            ws = (b.get("week_spread") or [None] * len(b["weeks"]))[k]
            if ws is not None:
                cv.text(xw + CW / 2, y + 36, f"{b['spread_name']} {signed(ws)}", 6.2, False, DARK)
            if b["week_partial"][k]:
                cv.text(xw + CW / 2, y + 48, "so far", 5.8, False, GREYTXT, italic=True)
        y += WH
    if target:
        cv.box(target[0], target[1], CW, WH, "none", DARK, 2.0)
    cv.cell(x0, y, w, FH, b["footer"], "white", 7.6, True, "black", "left")


def _recap_block(cv: Canvas, x0, y0, b: dict, height):
    """Week recap: one row per spread, one column per Mon-Sun week, plus the month."""
    w = CW * 8
    cv.cell(x0, y0, w, TH, b["title"], HEAD, 9.5, True)
    y, hh = y0 + TH, 30
    cw = (w - CW) / max(1, len(b["columns"]))
    cv.cell(x0, y, CW, hh, "Spread", "#F2F2F2", 8, True)
    for j, (l1, l2) in enumerate(b["columns"]):
        x = x0 + CW + cw * j
        cv.box(x, y, cw, hh, "#F2F2F2")
        cv.text(x + cw / 2, y + 10, l1, 8, True)
        cv.text(x + cw / 2, y + 22, l2, 6.5, False, GREYTXT)
    y += hh
    rh = min(26, (height - TH - hh - FH) / max(1, len(b["rows"])))
    scale = Scale([v for _, vals in b["rows"] for v in vals], mid=0)
    for name, vals in b["rows"]:
        cv.cell(x0, y, CW, rh, name, "white", 8.5, True)
        for j, v in enumerate(vals):
            cv.cell(x0 + CW + cw * j, y, cw, rh, signed(v), scale.color(v) or "white", 8.5, j == len(vals) - 1)
        y += rh
    cv.box(x0, y0 + height - FH, w, FH, "white")
    cv.text(x0 + 4, y0 + height - FH / 2, b["footer"], 7, False, GREYTXT, "left")


def calendars_png(title: str, blocks: list[dict], ncols: int = 4) -> bytes:
    """One picture with every country calendar and the week recap."""
    cals = [b for b in blocks if b["kind"] != "recap"]
    n_weeks = max((len(b["weeks"]) for b in cals), default=5)
    bw, bh, gap, top = CW * 8, TH + HH + WH * n_weeks + FH, 16, 34
    rows = math.ceil(len(blocks) / ncols)
    W = ncols * bw + (ncols + 1) * gap
    H = top + rows * (bh + gap) + gap
    cv = Canvas(W, H)
    cv.text(gap, top / 2 + 3, title, 12, True, "black", "left")
    for i, b in enumerate(blocks):
        r, c = divmod(i, ncols)
        x, y = gap + c * (bw + gap), top + r * (bh + gap)
        (_recap_block if b["kind"] == "recap" else _calendar_block)(cv, x, y, b, bh if b["kind"] == "recap" else n_weeks)
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
            sc = Scale([v for _, vals in b["rows"] for v in vals], mid=0)
            for k, (name, vals) in enumerate(b["rows"]):
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
