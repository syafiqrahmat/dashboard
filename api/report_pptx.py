"""Renders the Development progress report as a .pptx.

The deck mimics api/LKTN_Enhancement_Progress.pptx: 16:9 slides, drawn by
the PIL helpers below at SCALE=2 -- but unlike the first version, each
slide is emitted as NATIVE PowerPoint objects (shapes, text boxes, lines),
so every element stays editable/selectable in PowerPoint. Canvas records
each primitive as it paints; _emit_slide replays the records onto the
slide (see "native emitter"). render_slide still returns the PIL raster:
it is the visual reference the native output is diffed against. The only
picture on a slide is the Software Wizards logo (and the faint grid bg).

Copy is Malay, like the original, and the palette is the trio the original
deck leans on: slate (structure/neutral), amber (in progress), green (done).
"""
import io
import os
import unicodedata
from datetime import date, datetime, timedelta
from math import atan2, degrees, hypot

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn, nsdecls

# Design canvas (logical px) and supersampling factor. All coordinates in
# the drawing helpers below are logical; _px() maps them to device pixels.
W, H = 1376, 768
SCALE = 2

M = 60                      # side margin
BANNER_Y, BANNER_H = 648, 60

SLATE = (100, 116, 139)
SLATE_DARK = (71, 85, 105)
AMBER = (245, 185, 66)
GREEN = (47, 163, 107)
INK = (30, 41, 59)
MUTED = (148, 163, 184)
BG = (244, 246, 247)
CARD = (255, 255, 255)
TRACK = (226, 232, 240)
GRID = (233, 237, 242)
VIOLET = (167, 139, 250)
BORDER = (185, 196, 208)
WHITE = (255, 255, 255)

GREEN_BG = (226, 245, 236)
AMBER_BG = (254, 243, 220)
SLATE_BG = (237, 240, 244)

MONTHS_MS = ["Jan", "Feb", "Mac", "Apr", "Mei", "Jun",
             "Jul", "Ogo", "Sep", "Okt", "Nov", "Dis"]

STATUS_MS = {
    "completed": "Selesai",
    "inprogress": "Sedang Berjalan",
    "notstarted": "Belum Bermula",
}

_FONT_CACHE = {}
# id(font) -> (logical size, bold, family) -- the native-emitter side of the
# deck needs size/bold/family back from the PIL font object handed to c.text.
_FONT_INFO = {}


def _px(v):
    return int(round(v * SCALE))


def _font_files(bold):
    """Candidate font files, Windows first, then matplotlib's bundled
    DejaVu so the report still renders on a machine without Office fonts."""
    names = ["segoeuib.ttf", "tahomabd.ttf", "arialbd.ttf"] if bold else \
            ["segoeui.ttf", "tahoma.ttf", "arial.ttf"]
    out = [os.path.join("C:\\Windows\\Fonts", n) for n in names]
    try:
        import matplotlib
        ttf = os.path.join(os.path.dirname(matplotlib.__file__),
                           "mpl-data", "fonts", "ttf")
        out.append(os.path.join(ttf, "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"))
    except Exception:
        pass
    return out


def _font(size, bold=False):
    key = (round(size, 1), bold)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    font = None
    family = "Segoe UI"
    for path in _font_files(bold):
        if os.path.exists(path):
            font = ImageFont.truetype(path, _px(size))
            base = os.path.basename(path).lower()
            family = ("Segoe UI" if base.startswith("segoe") else
                      "Tahoma" if base.startswith("tahoma") else
                      "Arial" if base.startswith("arial") else "DejaVu Sans")
            break
    if font is None:
        font = ImageFont.load_default()
        family = "Arial"
    _FONT_CACHE[key] = font
    _FONT_INFO[id(font)] = (round(size, 1), bool(bold), family)
    return font


def _status_key(pct, status_text=None):
    if status_text:
        norm = str(status_text).lower().replace(" ", "")
        if norm in STATUS_MS:
            return norm
    if pct is None:
        return "notstarted"
    if pct >= 100:
        return "completed"
    if pct <= 0:
        return "notstarted"
    return "inprogress"


def _status_color(key):
    return {"completed": GREEN, "inprogress": AMBER}.get(key, SLATE)


def _status_bg(key):
    return {"completed": GREEN_BG, "inprogress": AMBER_BG}.get(key, SLATE_BG)


def _fmt_pct(pct):
    if pct is None:
        return "-"
    if abs(pct - round(pct)) < 0.05:
        return f"{int(round(pct))}%"
    return f"{pct:.1f}%"


def _dt(text):
    """dd/mm/yyyy -> date, or None."""
    if not text:
        return None
    try:
        return datetime.strptime(str(text).strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


# Software Wizards mark -- the same file the client-side Download Report
# (PPTX) loads (templates/dashboard.html, loadImageDataUrl). The deck's
# slides are grey and the PNG's background is white, so the near-white
# pixels are keyed out once and cached.
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "..", "static", "logo_report.png")
_LOGO_CACHE = {}


def _logo():
    """RGBA logo with its white background keyed out, or None when the
    file can't be read -- a missing logo must never fail the whole deck."""
    if "rgba" in _LOGO_CACHE:
        return _LOGO_CACHE["rgba"]
    try:
        src = Image.open(LOGO_PATH).convert("RGBA")
        px = src.load()
        for yy in range(src.size[1]):
            for xx in range(src.size[0]):
                r, g, b, _ = px[xx, yy]
                if r >= 245 and g >= 245 and b >= 245:
                    px[xx, yy] = (r, g, b, 0)
        _LOGO_CACHE["rgba"] = src
    except Exception:
        _LOGO_CACHE["rgba"] = None
    return _LOGO_CACHE["rgba"]


def _draw_logo(c, x, y, w):
    """Paste the logo's top-left corner at (x, y), w logical px wide."""
    logo = _logo()
    if not logo:
        return
    h = w * logo.size[1] / logo.size[0]
    c.logo(x, y, w, h, logo)


def _grid_paint(d):
    """Faint 64px background grid, shared by the PIL canvas and the
    native emitter's cached grid-background picture."""
    step = 64
    for x in range(step, W, step):
        d.line([_px(x), 0, _px(x), _px(H)], fill=GRID + (110,), width=_px(1))
    for y in range(step, H, step):
        d.line([0, _px(y), _px(W), _px(y)], fill=GRID + (110,), width=_px(1))


class Canvas:
    def __init__(self, bg=BG, grid=False, rec=None):
        self.bg = bg
        self.grid = grid
        self.rec = rec          # native-op sink; None = raster only
        self.img = Image.new("RGB", (_px(W), _px(H)), bg)
        self.d = ImageDraw.Draw(self.img, "RGBA")
        if grid:
            _grid_paint(self.d)

    def _op(self, kind, **kw):
        if self.rec is not None:
            self.rec.append({"op": kind, **kw})

    def rrect(self, xy, r, fill, outline=None, width=1, shadow=False):
        x0, y0, x1, y1 = xy
        if x1 <= x0 or y1 <= y0:
            return
        r = max(1, min(r, (x1 - x0) / 2, (y1 - y0) / 2))
        self._op("rrect", xy=(x0, y0, x1, y1), r=r, fill=fill,
                 outline=outline, width=width, shadow=shadow)
        if shadow:
            off = 5
            self.d.rounded_rectangle(
                [_px(x0) + _px(1), _px(y0) + _px(off), _px(x1) + _px(1), _px(y1) + _px(off)],
                radius=_px(r), fill=(15, 23, 42, 22))
        self.d.rounded_rectangle(
            [_px(x0), _px(y0), _px(x1), _px(y1)], radius=_px(r),
            fill=fill, outline=outline, width=_px(width))

    def rect(self, xy, fill):
        x0, y0, x1, y1 = xy
        if x1 <= x0 or y1 <= y0:
            return
        self._op("rect", xy=(x0, y0, x1, y1), fill=fill)
        self.d.rectangle([_px(x0), _px(y0), _px(x1), _px(y1)], fill=fill)

    def text(self, xy, s, font, fill, anchor="la", spc=None):
        self.d.text((_px(xy[0]), _px(xy[1])), s, font=font, fill=fill, anchor=anchor)
        if self.rec is not None:
            self._record_text(xy, [(s, font, fill)], anchor, spc=spc)

    def text_runs(self, xy, runs, anchor="la", space=None):
        """Draw word runs sequentially on one line (one editable text box)."""
        if not runs:
            return
        x, y = xy
        space = self.textw(" ", runs[0][1]) if space is None else space
        pos = x
        for s, font, fill in runs:
            self.d.text((_px(pos), _px(y)), s, font=font, fill=fill, anchor=anchor)
            pos += self.textw(s, font) + space
        if self.rec is not None:
            self._record_text(xy, runs, anchor, space)

    def _record_text(self, xy, runs, anchor, space=None, spc=None):
        # textlength() rejects "\n"; widest line wins for the frame width.
        w = sum(max(self.textw(part, font) for part in str(s).split("\n"))
                for s, font, _ in runs)
        if space is not None and len(runs) > 1:
            w += space * (len(runs) - 1)
        f0 = runs[0][1]
        meta = _FONT_INFO.get(id(f0))
        size, family = (meta[0], meta[2]) if meta else (14, "Segoe UI")
        try:
            asc, desc = f0.getmetrics()
            lh = (asc + desc) / float(SCALE)
        except Exception:
            lh = size * 1.25
        rec_runs = []
        for s, font, fill in runs:
            m = _FONT_INFO.get(id(font))
            rec_runs.append((str(s), m if m else (size, False, family), fill))
        lines = max(str(r[0]).count("\n") + 1 for r in rec_runs)
        self.rec.append({
            "op": "text", "xy": (xy[0], xy[1]), "anchor": anchor,
            "w": w, "lh": lh, "lines": lines, "runs": rec_runs,
            "spc": spc,
        })

    def textw(self, s, font):
        return self.d.textlength(s, font=font) / SCALE

    def line(self, xy, fill, width=1):
        xy = tuple(xy)
        self._op("line", xy=xy, fill=fill, width=width)
        self.d.line([_px(v) for v in xy], fill=fill, width=_px(width))

    def circle(self, cx, cy, r, fill=None, outline=None, width=1):
        self._op("circle", cx=cx, cy=cy, r=r, fill=fill, outline=outline,
                 width=width)
        self.d.ellipse([_px(cx - r), _px(cy - r), _px(cx + r), _px(cy + r)],
                       fill=fill, outline=outline, width=_px(width))

    def arc(self, cx, cy, r, a0, a1, fill, width=1):
        self._op("arc", cx=cx, cy=cy, r=r, a0=a0, a1=a1, fill=fill, width=width)
        box = [_px(cx - r), _px(cy - r), _px(cx + r), _px(cy + r)]
        self.d.arc(box, a0, a1, fill=fill, width=_px(width))

    def logo(self, x, y, w, h, image):
        scaled = image.resize((_px(w), _px(h)), Image.LANCZOS)
        self.img.paste(scaled, (_px(x), _px(y)), scaled)
        self._op("logo", xy=(x, y, w, h))


def _truncate(c, s, font, max_w):
    if c.textw(s, font) <= max_w:
        return s
    ell = "…"
    while s and c.textw(s + ell, font) > max_w:
        s = s[:-1]
    return s.rstrip() + ell


def _wrap(c, s, font, max_w):
    lines, cur = [], ""
    for word in str(s).split():
        trial = word if not cur else cur + " " + word
        if c.textw(trial, font) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines or [""]


def _wrap_segments(c, segments, reg, bold, max_w):
    """Wrap [(text, is_bold)] into lines of [(word, is_bold)].

    Pure-punctuation words (".", ",") glue onto the previous word instead of
    standing alone -- otherwise the space-before-token shows up as
    "44.6% ." in the banner."""
    tokens = []
    for text, is_bold in segments:
        for word in str(text).split():
            if tokens and word and all(ch in ".,;:!?" for ch in word):
                tokens[-1] = (tokens[-1][0] + word, tokens[-1][1])
            else:
                tokens.append((word, is_bold))
    space = c.textw(" ", reg)
    lines, cur, cur_w = [], [], 0.0
    for word, is_bold in tokens:
        f = bold if is_bold else reg
        w = c.textw(word, f)
        add = w if not cur else w + space
        if not cur or cur_w + add <= max_w:
            cur.append((word, is_bold))
            cur_w += add
        else:
            lines.append(cur)
            cur, cur_w = [(word, is_bold)], w
    if cur:
        lines.append(cur)
    return lines


def _draw_segments_line(c, line, cx, y, reg, bold, fill, fill_bold):
    if not line:
        return
    space = c.textw(" ", reg)
    widths = [c.textw(w, bold if b else reg) for w, b in line]
    total = sum(widths) + space * (len(line) - 1)
    x = cx - total / 2
    runs = [(word, bold if is_bold else reg,
             fill_bold if is_bold else fill) for word, is_bold in line]
    c.text_runs((x, y), runs, anchor="lm", space=space)


def _donut(c, cx, cy, r, pct, color, width=22, track=TRACK):
    c.arc(cx, cy, r, 0, 360, track, width)
    pct = max(0.0, min(float(pct or 0), 100.0))
    if pct > 0:
        c.arc(cx, cy, r, -90, -90 + 360 * pct / 100.0, color, width)


def _pill(c, x, y, text, key, size=14, h=28, align="left"):
    font = _font(size, True)
    padx = 14
    w = c.textw(text, font) + padx * 2
    if align == "right":
        x = x - w
    c.rrect([x, y, x + w, y + h], h / 2, _status_bg(key))
    c.text((x + w / 2, y + h / 2), text, font, _status_color(key), anchor="mm")
    return w


def _draw_title(c, title, sub=None, stamp=None):
    _draw_logo(c, 40, 14, 88)
    c.text((W / 2, 44), title, _font(32, True), INK, anchor="ma")
    if sub:
        c.text((W / 2, 92), sub, _font(15), MUTED, anchor="ma")
    if stamp:
        c.text((W - M, 46), stamp, _font(12), MUTED, anchor="ra")


def _banner(c, segments):
    """Bottom narrative banner (slate, bold highlights), like the original."""
    c.rrect([M, BANNER_Y, W - M, BANNER_Y + BANNER_H], 16, SLATE, shadow=True)
    reg, bold = _font(15), _font(15, True)
    lines = _wrap_segments(c, segments, reg, bold, W - 2 * M - 80)
    lh = 22
    top = BANNER_Y + BANNER_H / 2 - lh * len(lines) / 2
    for i, line in enumerate(lines):
        _draw_segments_line(c, line, W / 2, top + lh * i + lh / 2, reg, bold,
                            (219, 229, 238), WHITE)


# ---------------------------------------------------------------- slides ---

def _draw_cover(c, report):
    # Centred above the headline, mirroring the source deck's cover.
    _draw_logo(c, (W - 320) / 2, 44, 320)
    # Faint tech-illustration stand-in: concentric rings, bottom right.
    for r, col in ((230, SLATE), (176, AMBER), (126, GREEN), (80, SLATE)):
        c.circle(W - 150, 430, r, outline=col + (46,), width=10)
    c.circle(W - 150, 430, 44, outline=GREEN + (70,), width=6)

    c.text((M, 200), "Laporan Kemajuan Sistem:", _font(46, True), INK, anchor="la")
    c.text((M, 264), "Pembangunan (Development)", _font(46, True), SLATE, anchor="la")

    only = (report.get("clients") or [None])[0] if len(report.get("clients") or []) == 1 else None
    card_b = 516 if only else 486
    c.rrect([M, 372, M + 660, card_b], 16, CARD, shadow=True)
    c.rrect([M, 372, M + 10, card_b], 5, GREEN)
    c.text((M + 34, 402), f"Dijana pada {report.get('generated_at') or '-'}",
           _font(20, True), INK, anchor="la")
    s = report.get("summary") or {}
    c.text((M + 34, 440),
           f"{s.get('total', 0)} klien · {report.get('module_count', 0)} modul · "
           f"{report.get('milestone_count', 0)} milestone",
           _font(16), SLATE, anchor="la")
    if only:
        # Per-client deck: name the client on the cover.
        c.text((M + 34, 474),
               _truncate(c, f"Klien: {only.get('client') or '-'} · "
                            f"{only.get('projek_name') or '-'} · ID {only.get('projek_id') or '-'}",
                         _font(15), 660 - 68),
               _font(15), SLATE, anchor="la")

    c.text((M, H - 60), "Dijana automatik daripada dashboard — jangan sunting tangan.",
           _font(13), MUTED, anchor="la")


def _draw_gauges(c, report):
    clients = report.get("clients") or []
    _draw_title(c, "Kemajuan Klien Development",
                "Purata Percentage daripada baris Project Details setiap klien",
                stamp=report.get("generated_at"))

    n = max(len(clients), 1)
    rows = 1 if n <= 3 else 2
    per_row = (n + rows - 1) // rows
    top, bottom = 128, 632
    gap_x, gap_y = 30, 24
    row_h = (bottom - top - gap_y * (rows - 1)) / rows

    for i, cl in enumerate(clients):
        r, col = divmod(i, per_row)
        count = min(per_row, n - r * per_row)
        # Cap the card width and centre the row: with one client a
        # full-bleed card looks empty, a ~560px card reads as a stat panel.
        cw = min((W - 2 * M - gap_x * (count - 1)) / count, 560)
        row_w = count * cw + gap_x * (count - 1)
        x0 = M + ((W - 2 * M) - row_w) / 2 + col * (cw + gap_x)
        y0 = top + r * (row_h + gap_y)
        x1, y1 = x0 + cw, y0 + row_h

        key = _status_key(cl.get("overall_pct"), cl.get("progress_status"))
        c.rrect([x0, y0, x1, y1], 18, CARD, shadow=True)

        c.text(((x0 + x1) / 2, y0 + 34), _truncate(c, cl.get("client") or "-",
               _font(23, True), cw - 40), _font(23, True), INK, anchor="ma")
        c.text(((x0 + x1) / 2, y0 + 64),
               _truncate(c, cl.get("projek_name") or "-", _font(14), cw - 40),
               _font(14), MUTED, anchor="ma")

        pct = cl.get("overall_pct")
        cx, cy = (x0 + x1) / 2, y0 + row_h / 2 + 8
        _donut(c, cx, cy, 74, pct, _status_color(key), width=24)
        c.text((cx, cy - 8), _fmt_pct(pct), _font(34, True), INK, anchor="mm")
        c.text((cx, cy + 24), "kemajuan", _font(13), MUTED, anchor="mm")

        _pill(c, (x0 + x1) / 2, y1 - 74, STATUS_MS[key], key, align="center")
        if cl.get("technology"):
            c.text((cx, y1 - 36), cl.get("technology"), _font(12), MUTED, anchor="ma")

    s = report.get("summary") or {}
    _banner(c, [
        ("Daripada ", False), (f"{s.get('total', 0)} klien", True),
        (" Development: ", False), (f"{s.get('completed', 0)} Selesai", True),
        (", ", False), (f"{s.get('in_progress', 0)} Sedang Berjalan", True),
        (", ", False), (f"{s.get('not_started', 0)} Belum Bermula", True),
        (". Purata kemajuan keseluruhan ", False),
        (f"{report.get('avg_pct') or '-'}%", True), (".", False),
    ])


def _contract_window(clients):
    """(earliest start, latest end) across clients' contract dates, as
    dd/mm/yyyy strings -- either side falls back to '-' when no client
    carries a usable date."""
    starts = [d for d in (_dt(cl.get("start_date")) for cl in clients) if d]
    ends = [d for d in (_dt(cl.get("end_date")) for cl in clients) if d]
    lo = min(starts) if starts else None
    hi = max(ends) if ends else None
    return (lo.strftime("%d/%m/%Y") if lo else "-",
            hi.strftime("%d/%m/%Y") if hi else "-")


def _ms_done(ms):
    """Milestones counted as complete (Progress >= 100), shared by the
    Executive Summary line and the per-client milestone slide."""
    return sum(1 for m in ms if (m.get("progress") or 0) >= 100)


def _draw_exec_summary(c, report):
    """Executive Summary -- the deck's port of the Download Report (PPTX)
    slide 2 (templates/dashboard.html buildProjectReportPptx): one big
    overall-completion ring, the contract window, and up to four module
    ring cards. Module cards flatten every client's modules and sort the
    lowest percentage first, so the row reads as a risk list; the deck's
    per-client Gauges slide already carries one ring per client."""
    clients = report.get("clients") or []
    _draw_title(c, "Executive Summary",
                "Kemajuan keseluruhan klien Development",
                stamp=report.get("generated_at"))

    avg = report.get("avg_pct")
    # Left card -- the overall ring, sized like the gauges slide's donut.
    c.rrect([M, 130, M + 400, 350], 16, CARD, shadow=True)
    cx, cy = M + 200, 218
    _donut(c, cx, cy, 76, avg, _status_color(_status_key(avg)), width=26)
    c.text((cx, cy - 6), _fmt_pct(avg), _font(34, True), INK, anchor="mm")
    c.text((cx, cy + 24), "kemajuan", _font(13), MUTED, anchor="mm")
    c.text((cx, 304), "Kemajuan Keseluruhan", _font(15, True), INK, anchor="ma", spc=2)
    c.text((cx, 326),
           f"{report.get('module_count', 0)} modul · "
           f"{report.get('milestone_count', 0)} milestone",
           _font(12), MUTED, anchor="ma")

    # Contract window -- the analogue of the source slide's
    # "Contract window: x to y." line, set beside the ring.
    lo, hi = _contract_window(clients)
    c.text((M + 440, 168), f"Tempoh kontrak: {lo} hingga {hi}",
           _font(17, True), INK, anchor="la")
    c.text((M + 440, 198),
           f"{len(clients)} klien Development dalam laporan ini.",
           _font(13), MUTED, anchor="la")
    total_ms = report.get("milestone_count", 0)
    done_ms = _ms_done([m for cl in clients for m in (cl.get("milestones") or [])])
    if total_ms:
        c.text((M + 440, 226),
               f"{done_ms} daripada {total_ms} milestone siap.",
               _font(15, True),
               GREEN if done_ms >= total_ms else INK, anchor="la", spc=2)
    else:
        c.text((M + 440, 226), "Tiada milestone direkod.",
               _font(13), MUTED, anchor="la")

    # Module cards -- capped at four, like the source slide's slice(0, 4).
    mods = []
    for cl in clients:
        for m in (cl.get("modules") or cl.get("task_detail_modules") or []):
            if not m.get("name"):
                continue
            mods.append({"client": cl.get("client") or "-",
                         "name": m.get("name"), "pct": m.get("pct"),
                         "rows": m.get("rows") or 0})
    mods.sort(key=lambda m: (m.get("pct") is None, m.get("pct") or 0.0))

    cardsY, cardH, cardW, gap = 392, 210, 296, 18
    n = min(len(mods), 4)
    if not n:
        c.text((W / 2, 490), "Tiada modul direkod.",
               _font(15), MUTED, anchor="ma")
    else:
        total_w = n * cardW + (n - 1) * gap
        start_x = M + ((W - 2 * M) - total_w) / 2
        name_f, meta_f, pct_f = _font(13, True), _font(11), _font(20, True)
        for i, m in enumerate(mods[:n]):
            x0 = start_x + i * (cardW + gap)
            x1 = x0 + cardW
            c.rrect([x0, cardsY, x1, cardsY + cardH], 16, CARD,
                    outline=BORDER, width=2)
            mid = (x0 + x1) / 2
            pct = m.get("pct")
            _donut(c, mid, cardsY + 76, 52, pct,
                   _status_color(_status_key(pct)), width=18)
            c.text((mid, cardsY + 76), _fmt_pct(pct), pct_f,
                   _status_color(_status_key(pct)), anchor="mm")
            all_lines = _wrap(c, m["name"], name_f, cardW - 30)
            if len(all_lines) > 2:
                all_lines = all_lines[:2]
                all_lines[1] = _truncate(c, all_lines[1] + "…",
                                         name_f, cardW - 30)
            for j, line in enumerate(all_lines):
                c.text((mid, cardsY + 136 + j * 17), line, name_f, INK,
                       anchor="ma")
            c.text((mid, cardsY + 170),
                   _truncate(c, m["client"], meta_f, cardW - 24),
                   meta_f, MUTED, anchor="ma")
            c.text((mid, cardsY + 186), f"{m['rows']} baris",
                   meta_f, MUTED, anchor="ma")

    if len(mods) > 4:
        c.text((W / 2, 616),
               f"+ {len(mods) - 4} modul lagi — lihat slide Status Modul.",
               _font(12), MUTED, anchor="ma")

    s = report.get("summary") or {}
    _banner(c, [
        ("Purata kemajuan keseluruhan ", False),
        (f"{report.get('avg_pct') or '-'}%", True),
        (" merentasi ", False), (f"{s.get('total', 0)} klien", True),
        (", ", False), (f"{report.get('module_count', 0)} modul", True),
        (", ", False), (f"{report.get('milestone_count', 0)} milestone", True),
        (". Tempoh kontrak ", False), (f"{lo} hingga {hi}", True), (".", False),
    ])


# Project Details rows drawn per "Garis Masa Pembangunan" slide: one slide
# per Title, holding that Title's Description rows. 15 fits the longest
# Title in the data (LKTN Procurement) inside the 390px grid alongside the
# client header line; a Title that ever outgrows it paginates within
# itself rather than squeezing the labels.
TIMELINE_ROWS_PER_PAGE = 15


def _dates_text(row):
    """The row's Target period as "dd/mm/yyyy – dd/mm/yyyy".

    A single-day period (start == end, common in the Target columns)
    prints one date instead of "x – x".
    """
    sd, ed = row.get("start_date") or "", row.get("end_date") or ""
    if not sd or not ed:
        return sd or ed or "-"
    return sd if sd == ed else f"{sd} – {ed}"


def _timeline_slides(report):
    """Flatten every client's `timeline` rows and split them by Title.

    Returns (slides, axis, bounds): one slide per Title -- titled by it --
    each carrying that Title's Description rows; axis is the padded
    (min, max) date range every slide draws the SAME month grid from, so
    bars on any two slides line up; bounds is the unpadded range the
    narrative banner quotes. Rows without a usable period are dropped --
    the rule the old per-client bars followed -- and each row carries its
    client context for the header line.
    """
    rows = []
    for cl in report.get("clients") or []:
        for t in cl.get("timeline") or []:
            if not (_dt(t.get("start_date")) and _dt(t.get("end_date"))):
                continue
            row = dict(t)
            row.update({
                "client": cl.get("client"),
                "projek_name": cl.get("projek_name"),
                "overall_pct": cl.get("overall_pct"),
                "progress_status": cl.get("progress_status"),
            })
            rows.append(row)
    if not rows:
        return [], None, None
    bounds = (min(_dt(r["start_date"]) for r in rows),
              max(_dt(r["end_date"]) for r in rows))
    axis = (bounds[0] - timedelta(days=20), bounds[1] + timedelta(days=20))

    # Group by (client, title) in arrival order: rows arrive client by
    # client in Project Details order, so Title blocks stay contiguous and
    # one client's slides stay together.
    groups = []
    for r in rows:
        key = (r.get("client"), r.get("title"))
        if not groups or groups[-1][0] != key:
            groups.append((key, []))
        groups[-1][1].append(r)

    slides = []
    for _, grp in groups:
        chunks = [grp[i:i + TIMELINE_ROWS_PER_PAGE]
                  for i in range(0, len(grp), TIMELINE_ROWS_PER_PAGE)]
        for i, chunk in enumerate(chunks, start=1):
            slides.append({
                "title": grp[0].get("title"),
                "client": grp[0].get("client"),
                "projek_name": grp[0].get("projek_name"),
                "rows": chunk,
                "page": i,
                "pages": len(chunks),
            })
    return slides, axis, bounds


def _draw_timeline(c, data):
    report = data.get("report") or {}
    rows = data.get("rows") or []
    axis = data.get("axis")
    bounds = data.get("bounds")
    page, pages = data.get("page") or 1, data.get("pages") or 1
    title = data.get("title")

    sub = (f"Tempoh Target · {title}" if title
           else "Tempoh Target setiap baris Project Details")
    if pages > 1:
        sub += f" · Halaman {page}/{pages}"
    _draw_title(c, "Garis Masa Pembangunan", sub, stamp=report.get("generated_at"))

    if not rows or not axis:
        c.text((W / 2, 340), "Tiada tarikh mula/tamat direkod.", _font(18), MUTED, anchor="ma")
        _banner(c, [("Tiada data garis masa untuk dipaparkan.", True)])
        return

    min_d, max_d = axis
    span = max((max_d - min_d).days, 1)
    x0, x1 = 330, W - M - 70
    plot_w = x1 - x0

    def x_of(d):
        return x0 + (d - min_d).days / span * plot_w

    # Month grid + labels
    axis_y, grid_top, grid_bottom = 138, 166, 556
    months = (max_d.year - min_d.year) * 12 + (max_d.month - min_d.month) + 1
    step = 1
    while months / step * 66 > plot_w and step < 24:
        step += 1
    y_, m_ = min_d.year, min_d.month
    while (y_, m_) <= (max_d.year, max_d.month):
        tick = date(y_, m_, 1)
        if min_d <= tick <= max_d:
            x = x_of(tick)
            c.line([x, grid_top, x, grid_bottom], GRID, 1)
            if ((y_ - min_d.year) * 12 + m_ - min_d.month) % step == 0:
                c.text((x, axis_y), f"{MONTHS_MS[m_ - 1]} {y_}",
                       _font(12, True), SLATE, anchor="ma")
        m_ += 1
        if m_ > 12:
            m_, y_ = 1, y_ + 1

    # Today marker. The pill sits above the grid, so it can be drawn early;
    # the dashed line is drawn AFTER the bars (see below) -- bars crossing
    # it used to bury it. tx is where every bar splits into past/future.
    today = date.today()
    if today <= min_d:
        tx = x0
    elif today >= max_d:
        tx = x1 + 1
    else:
        tx = x_of(today)
    if min_d <= today <= max_d:
        c.rrect([tx - 44, 112, tx + 44, 134], 11, SLATE_DARK)
        c.text((tx, 123), "Hari Ini", _font(12, True), WHITE, anchor="mm")

    # One slide = one Title. The header names the client/project the Title
    # belongs to and the Title itself sits as a section line above the rows;
    # the slide carries no percentages any more, so each row spends its
    # second line on its Target period instead.
    header_h = 52 if title else 26
    grid_h = grid_bottom - grid_top
    row_h = min(max(16, (grid_h - header_h) / len(rows)), 64)
    t_size = 14 if row_h >= 26 else 12
    d_size = 10 if row_h >= 26 else 9
    desc_f, date_f = _font(t_size, True), _font(d_size)
    block = t_size + 2 + d_size

    y = grid_top + max(0, (grid_h - (row_h * len(rows) + header_h)) / 2)
    c.text((M, y + 4),
           _truncate(c, f"{data.get('client') or '-'} · {data.get('projek_name') or '-'}",
                     _font(15, True), 300),
           _font(15, True), INK, anchor="la", spc=2)
    c.line([M, y + 23, x1, y + 23], TRACK, 1)
    if title:
        c.text((M, y + 28), _truncate(c, title, _font(17, True), 300),
               _font(17, True), INK, anchor="la")
    y += header_h

    for r in rows:
        cy = y + row_h / 2
        y += row_h

        s, e = _dt(r.get("start_date")), _dt(r.get("end_date"))
        if not (s and e):
            continue

        top = cy - block / 2
        c.text((M, top), _truncate(c, r.get("description") or "-", desc_f, 250),
               desc_f, INK, anchor="la")
        c.text((M, top + t_size + 2), _truncate(c, _dates_text(r), date_f, 250),
               date_f, MUTED, anchor="la")

        bx0, bx1 = x_of(s), x_of(e)
        # Target dates are often a single day (start == end), and a
        # zero-width rect draws nothing at all -- keep every row's bar at
        # least a few px wide so those rows still read as a bar.
        bx1 = max(bx1, bx0 + 8)
        bh = max(4, min(13, row_h * 0.3))
        c.rrect([bx0, cy - bh, bx1, cy + bh], bh, TRACK)
        bar_w = bx1 - bx0
        pct = r.get("pct") or 0
        fill_w = bar_w * max(0.0, min(float(pct), 100.0)) / 100.0
        if fill_w > 2:
            c.rrect([bx0, cy - bh, bx0 + min(max(fill_w, 8), bar_w), cy + bh],
                    bh, _status_color(_status_key(pct)))
        # Everything right of the Hari Ini line is future time -- paint it
        # solid violet over the fill so those stretches (which used to show
        # only the near-invisible TRACK colour) read at a glance.
        if bx1 > tx:
            vx0 = max(bx0, tx)
            c.rrect([vx0, cy - bh, bx1, cy + bh], bh, VIOLET)
            if vx0 > bx0:
                # A mid-bar start would round the segment's left edge too,
                # letting the fill colour poke through around the Hari Ini
                # line -- square that cap off.
                c.rrect([vx0, cy - bh, min(vx0 + bh, bx1), cy + bh], 1, VIOLET)

    # Redraw the dashed today line over the bars: bars crossing it used to
    # bury it, so the violet segments have no visible boundary.
    if min_d <= today <= max_d:
        y = grid_top
        while y < grid_bottom:
            c.line([tx, y, tx, min(y + 9, grid_bottom)], SLATE_DARK, 2)
            y += 15

    # Narrative banner -- one global period for the whole report, so every
    # slide quotes the same range.
    lo, hi = bounds or (min_d, max_d)
    _banner(c, [
        ("Tempoh projek merentasi ", False),
        (f"{MONTHS_MS[lo.month - 1]} {lo.year} hingga "
         f"{MONTHS_MS[hi.month - 1]} {hi.year}", True),
        (". Penanda hari ini ", False),
        (today.strftime("%d/%m/%Y"), True),
        (" menunjukkan klien masih dalam fasa pembangunan",
         False),
        (" · bahagian selepas hari ini ditunjukkan dalam ungu", False),
        (".", False),
    ])


def _bar_row(c, x, y, w, label, pct, row_h, label_w):
    key = _status_key(pct)
    font = _font(13)
    c.text((x, y + row_h / 2), _truncate(c, label, font, label_w), font, INK, anchor="lm")
    tx0 = x + label_w + 12
    tx1 = x + w - 52
    c.rrect([tx0, y + row_h / 2 - 6, tx1, y + row_h / 2 + 6], 6, TRACK)
    p = max(0.0, min(float(pct or 0), 100.0))
    if p > 0:
        fw = (tx1 - tx0) * p / 100.0
        c.rrect([tx0, y + row_h / 2 - 6, tx0 + max(fw, 12), y + row_h / 2 + 6], 6,
                _status_color(key))
    c.text((x + w, y + row_h / 2), _fmt_pct(pct), _font(13, True),
           _status_color(key), anchor="rm")


def _draw_client(c, report, cl):
    key = _status_key(cl.get("overall_pct"), cl.get("progress_status"))
    _draw_title(c, f"Status Modul: {cl.get('projek_name') or '-'}",
                f"{cl.get('client')} · ID {cl.get('projek_id') or '-'} · "
                f"{cl.get('technology') or '-'}",
                stamp=report.get("generated_at"))

    # Profile card
    px0, py0, px1, py1 = M, 124, M + 380, 440
    c.rrect([px0, py0, px1, py1], 16, CARD, shadow=True)
    c.text((px0 + 26, py0 + 30), "Profil Projek", _font(17, True), INK, anchor="la")
    fields = [
        ("ID Projek", cl.get("projek_id")),
        ("Teknologi", cl.get("technology")),
        ("Mula", cl.get("start_date")),
        ("Tamat", cl.get("end_date")),
    ]
    fy = py0 + 64
    for label, val in fields:
        c.text((px0 + 26, fy), label, _font(13), MUTED, anchor="la")
        c.text((px1 - 26, fy), _truncate(c, val or "-", _font(14, True), 190),
               _font(14, True), INK, anchor="ra")
        c.line([px0 + 26, fy + 24, px1 - 26, fy + 24], TRACK, 1)
        fy += 44
    c.text((px0 + 26, fy + 4), "Status", _font(13), MUTED, anchor="la")
    _pill(c, px0 + 92, fy - 8, STATUS_MS[key], key)
    c.text((px0 + 26, fy + 34), _fmt_pct(cl.get("overall_pct")),
           _font(34, True), _status_color(key), anchor="la")
    c.text((px0 + 140, fy + 42), "kemajuan\nkeseluruhan", _font(12), MUTED, anchor="la")

    # Module bars card (project modules + task-detail modules side by side)
    bx0, by0, bx1, by1 = M + 410, 124, W - M, 440
    c.rrect([bx0, by0, bx1, by1], 16, CARD, shadow=True)
    modules = cl.get("modules") or []
    td = cl.get("task_detail_modules") or []
    avail = (by1 - by0) - 58 - 14

    if modules and td:
        col_w = (bx1 - bx0 - 78) / 2
        groups = [((bx0 + 24), col_w, "Project Details", modules),
                  ((bx0 + 54 + col_w), col_w, "Task Detail", td)]
    else:
        groups = [((bx0 + 24), bx1 - bx0 - 48, "Kemajuan Modul", modules or td)]

    if not modules and not td:
        c.text(((bx0 + bx1) / 2, (by0 + by1) / 2),
               "Tiada baris modul untuk projek ini.", _font(16), MUTED, anchor="mm")

    for gx, gw, header, items in groups:
        c.text((gx, by0 + 30), header, _font(16, True), INK, anchor="la")
        c.line([gx, by0 + 48, gx + gw, by0 + 48], TRACK, 1)
        if not items:
            c.text((gx, by0 + 80), "Tiada data.", _font(13), MUTED, anchor="la")
            continue
        cap = int(avail // 26)
        shown, overflow = (items[:cap], len(items) - cap) if len(items) > cap else (items, 0)
        row_h = min(30, avail / max(len(shown) + (0.45 if overflow else 0), 1))
        ry = by0 + 60
        label_w = max(90, gw * 0.34)
        for it in shown:
            _bar_row(c, gx, ry, gw, it.get("name") or "-", it.get("pct"), row_h, label_w)
            ry += row_h
        if overflow:
            c.text((gx, ry + 8), f"+ {overflow} lagi…", _font(12, True), MUTED, anchor="la")

    # Milestones as chips
    ms = cl.get("milestones") or []
    mx0, my0, mx1, my1 = M, 458, W - M, 636
    c.rrect([mx0, my0, mx1, my1], 16, CARD, shadow=True)
    c.text((mx0 + 24, my0 + 26), f"Pencapaian (Milestone) — {len(ms)}",
           _font(16, True), INK, anchor="la")
    if not ms:
        c.text((mx0 + 24, my0 + 64), "Tiada milestone direkod untuk projek ini.",
               _font(13), MUTED, anchor="la")
    else:
        chip_w, chip_h, gap = 234, 32, 8
        per_row = int((mx1 - mx0 - 48 + gap) // (chip_w + gap))
        cap = per_row * 3
        shown, overflow = (ms[:cap], len(ms) - cap) if len(ms) > cap else (ms, 0)
        cy = my0 + 54
        for i, m in enumerate(shown):
            r, col = divmod(i, per_row)
            if r > 2:
                break
            x = mx0 + 24 + col * (chip_w + gap)
            y = cy + r * (chip_h + gap)
            mk = _status_key(m.get("progress"))
            c.rrect([x, y, x + chip_w, y + chip_h], 9, _status_bg(mk))
            c.circle(x + 17, y + chip_h / 2, 7, fill=_status_color(mk))
            if mk == "completed":
                c.line([x + 13, y + chip_h / 2 + 1, x + 16, y + chip_h / 2 + 4], WHITE, 2)
                c.line([x + 16, y + chip_h / 2 + 4, x + 21, y + chip_h / 2 - 3], WHITE, 2)
            else:
                c.line([x + 13, y + chip_h / 2, x + 21, y + chip_h / 2], WHITE, 2)
            name = _truncate(c, m.get("task_name") or "-", _font(13, True), chip_w - 92)
            c.text((x + 32, y + chip_h / 2), name, _font(13, True), INK, anchor="lm")
            c.text((x + chip_w - 12, y + chip_h / 2), _fmt_pct(m.get("progress")),
                   _font(13, True), _status_color(mk), anchor="rm")
        if overflow:
            c.text((mx1 - 24, my0 + 26), f"+ {overflow} lagi…",
                   _font(12, True), MUTED, anchor="ra")

    # Narrative banner
    mods = [m for m in (modules or td) if m.get("pct") is not None]
    segments = [(f"{cl.get('client')}: ", True)]
    if mods:
        hi = max(mods, key=lambda m: m["pct"])
        lo = min(mods, key=lambda m: m["pct"])
        segments += [
            (f"{len(mods)} modul, kemajuan ", False),
            (_fmt_pct(cl.get("overall_pct")), True),
            (". Tertinggi ", False), (f"{hi['name']} ({_fmt_pct(hi['pct'])})", True),
            (", terendah ", False), (f"{lo['name']} ({_fmt_pct(lo['pct'])})", True),
            (".", False),
        ]
    else:
        segments += [("kemajuan keseluruhan ", False),
                     (_fmt_pct(cl.get("overall_pct")), True),
                     (" pada status ", False), (STATUS_MS[key], True), (".", False)]
    _banner(c, segments)


def _draw_milestones(c, data):
    """One slide per client holding that client's full milestone list --
    the full detail the Status Modul slide only hints at with its chip row
    (which drops the dates and caps at three rows): name, Target period
    and a progress bar per milestone, ordered by start date."""
    report = data.get("report") or {}
    cl = data.get("client") or {}
    ms = sorted(cl.get("milestones") or [],
                key=lambda m: (_dt(m.get("start_date")) or date.max,
                               _dt(m.get("end_date")) or date.max))
    _draw_title(c, f"Pencapaian (Milestone): {cl.get('projek_name') or '-'}",
                f"{cl.get('client')} · ID {cl.get('projek_id') or '-'} · "
                f"{len(ms)} milestone",
                stamp=report.get("generated_at"))

    cx0, cy0, cx1, cy1 = M, 124, W - M, 636
    c.rrect([cx0, cy0, cx1, cy1], 16, CARD, shadow=True)
    if not ms:
        c.text(((cx0 + cx1) / 2, (cy0 + cy1) / 2),
               "Tiada milestone direkod untuk projek ini.",
               _font(16), MUTED, anchor="mm")
        _banner(c, [(f"{cl.get('client') or '-'}: ", True),
                    ("tiada milestone direkod.", False)])
        return

    top, bottom = 194, 616
    avail = bottom - top
    max_rows = int(avail // 26)
    if len(ms) <= max_rows:
        shown, overflow = ms, 0
        row_h = min(44, avail / len(ms))
    else:
        shown, overflow = ms[:max_rows], len(ms) - max_rows
        row_h = avail / max_rows

    name_x, start_x, end_x = cx0 + 26, cx0 + 700, cx0 + 820
    bar_x0, bar_x1, pct_x = cx0 + 960, cx0 + 1180, cx1 - 26
    head_f = _font(13, True)
    name_f, date_f, pct_f = _font(14, True), _font(13), _font(14, True)
    for label, lx in (("Milestone", name_x), ("Mula", start_x),
                      ("Tamat", end_x), ("Kemajuan", bar_x0)):
        c.text((lx, 152), label, head_f, INK, anchor="la")
    if overflow:
        c.text((cx1 - 26, 152), f"+ {overflow} milestone lagi…",
               head_f, MUTED, anchor="ra")
    c.line([cx0 + 26, 178, cx1 - 26, 178], TRACK, 1)

    # Rows top-align under the header, table-style; row height only grows
    # when a short list would otherwise leave the card looking empty.
    y = top
    for i, m in enumerate(shown):
        cy = y + row_h / 2
        if i % 2:
            c.rrect([cx0 + 16, y + 2, cx1 - 16, y + row_h - 2], 6,
                    TRACK + (95,))
        y += row_h
        c.text((name_x, cy),
               _truncate(c, m.get("task_name") or "-", name_f, 654),
               name_f, INK, anchor="lm")
        c.text((start_x, cy), m.get("start_date") or "-", date_f, SLATE,
               anchor="lm")
        c.text((end_x, cy), m.get("end_date") or "-", date_f, SLATE,
               anchor="lm")
        key = _status_key(m.get("progress"))
        c.rrect([bar_x0, cy - 6, bar_x1, cy + 6], 6, TRACK)
        fw = (bar_x1 - bar_x0) * max(0.0, min(float(m.get("progress") or 0),
                                              100.0)) / 100.0
        if fw > 2:
            c.rrect([bar_x0, cy - 6, bar_x0 + max(fw, 6), cy + 6], 6,
                    _status_color(key))
        c.text((pct_x, cy), _fmt_pct(m.get("progress")), pct_f,
               _status_color(key), anchor="rm")

    have = [m.get("progress") for m in ms if m.get("progress") is not None]
    avg = round(sum(have) / len(have), 1) if have else None
    _banner(c, [
        (f"{cl.get('client') or '-'}: ", True), (f"{len(ms)} milestone", True),
        (", ", False), (f"{_ms_done(ms)} siap", True),
        (" · purata kemajuan ", False), (_fmt_pct(avg), True), (".", False),
    ])


def _draw_summary(c, report):
    clients = report.get("clients") or []
    s = report.get("summary") or {}
    _draw_title(c, "Rumusan & Langkah Seterusnya",
                "Ringkasan status klien Development", stamp=report.get("generated_at"))

    buckets = [("completed", "Selesai", GREEN),
               ("in_progress", "Sedang Berjalan", AMBER),
               ("not_started", "Belum Bermula", SLATE)]
    # summary counts use in_progress/not_started; _status_key() returns
    # inprogress/notstarted -- map before matching members to a bucket.
    bkey_to_status = {"completed": "completed", "in_progress": "inprogress",
                      "not_started": "notstarted"}
    gap = 32
    cw = (W - 2 * M - gap * 2) / 3
    for i, (bkey, label, col) in enumerate(buckets):
        x0 = M + i * (cw + gap)
        y0, y1 = 134, 372
        c.rrect([x0, y0, x0 + cw, y1], 16, CARD, shadow=True)
        c.rrect([x0, y0, x0 + cw, y0 + 66], 16, col)
        c.rect([x0, y0 + 40, x0 + cw, y0 + 66], col)
        c.text((x0 + cw / 2, y0 + 33), label, _font(19, True), WHITE, anchor="mm")
        c.text((x0 + cw / 2, y0 + 118), str(s.get(bkey, 0)),
               _font(54, True), col, anchor="mm")
        members = [cl.get("client") or "-" for cl in clients
                   if _status_key(cl.get("overall_pct"), cl.get("progress_status"))
                   == bkey_to_status[bkey]]
        text = " · ".join(members) if members else "—"
        lines = _wrap(c, text, _font(14), cw - 48)
        ly = y0 + 176
        for line in lines[:4]:
            c.text((x0 + cw / 2, ly), line, _font(14), SLATE, anchor="ma")
            ly += 22

    # Next steps
    nx0, ny0, nx1, ny1 = M, 396, W - M, 632
    c.rrect([nx0, ny0, nx1, ny1], 16, CARD, shadow=True)
    c.text((nx0 + 26, ny0 + 32), "Langkah Seterusnya", _font(18, True), INK, anchor="la")
    steps = []
    for cl in clients:
        mods = [m for m in (cl.get("modules") or []) if m.get("pct") is not None]
        if not mods:
            continue
        low = min(mods, key=lambda m: m["pct"])
        if (low.get("pct") or 0) < 100:
            steps.append(
                (cl,
                 f"Segerakan modul {low['name']} ({_fmt_pct(low['pct'])}) untuk "
                 f"{cl.get('client')} — tamat {cl.get('end_date') or '-'}"))
    if not steps:
        steps = [(None, "Semua modul telah mencapai 100% — sedia untuk serahan UAT/FAT.")]
    sy = ny0 + 74
    for cl, text in steps[:4]:
        key = _status_key(cl.get("overall_pct") if cl else None,
                          cl.get("progress_status") if cl else "completed")
        c.circle(nx0 + 34, sy + 10, 6, fill=_status_color(key))
        for j, line in enumerate(_wrap(c, text, _font(15), nx1 - nx0 - 90)[:2]):
            c.text((nx0 + 54, sy + 10 + j * 22), line, _font(15), INK, anchor="la")
            if j == 0:
                sy += 0
        sy += 46
    if len(steps) > 4:
        c.text((nx1 - 26, ny0 + 32), f"+ {len(steps) - 4} lagi…",
               _font(12, True), MUTED, anchor="ra")

    _banner(c, [
        ("Purata kemajuan keseluruhan ", False),
        (f"{report.get('avg_pct') or '-'}%", True), (" merentasi ", False),
        (f"{s.get('total', 0)} klien", True), (" Development: ", False),
        (f"{s.get('completed', 0)} Selesai", True), (", ", False),
        (f"{s.get('in_progress', 0)} Sedang Berjalan", True), (", ", False),
        (f"{s.get('not_started', 0)} Belum Bermula", True), (".", False),
    ])


def _slide_canvas(kind, data, rec=None):
    """Draw one slide onto a Canvas (rec=[]) and return it."""
    if kind == "cover":
        c = Canvas(grid=False, rec=rec)
        _draw_cover(c, data)
    elif kind == "gauges":
        c = Canvas(grid=True, rec=rec)
        _draw_gauges(c, data)
    elif kind == "exec":
        c = Canvas(grid=True, rec=rec)
        _draw_exec_summary(c, data)
    elif kind == "timeline":
        c = Canvas(grid=True, rec=rec)
        _draw_timeline(c, data)
    elif kind == "client":
        c = Canvas(grid=True, rec=rec)
        _draw_client(c, data["report"], data["client"])
    elif kind == "milestones":
        c = Canvas(grid=True, rec=rec)
        _draw_milestones(c, data)
    elif kind == "summary":
        c = Canvas(grid=True, rec=rec)
        _draw_summary(c, data)
    else:
        raise ValueError(f"unknown slide kind: {kind}")
    return c


def render_slide(kind, data):
    """Render one slide and return a PIL image (1376x768 * SCALE)."""
    return _slide_canvas(kind, data).img


def _png_bytes(img):
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    buf.seek(0)
    return buf.read()


# ------------------------------------------------------- native emitter ---
# Canvas recorded every primitive as it painted; here the records are
# replayed onto the slide as native PowerPoint objects. Geometry: the
# 1376x768 logical canvas maps uniformly onto the 16256000x9144000 EMU
# slide (circles stay circles), vertically centred like the old inset.

EMU_PX = 16256000.0 / W                 # logical px -> EMU (both axes)
EMU_TOP = round((9144000.0 - H * EMU_PX) / 2)

_ALIGN = {"l": PP_ALIGN.LEFT, "m": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}

# PowerPoint seats glyphs slightly higher inside its line box than PIL does
# (different ascent/line-gap handling). These per-anchor nudges -- fitted on
# a probe slide exported through PowerPoint against the PIL reference --
# shift the text box back onto the raster's ink position. Units: logical
# canvas px, as k * font_size + b.
_TEXT_NUDGE = {
    "la": (0.09943, 0.773),
    "ra": (0.09943, 0.773),
    "ma": (0.10889, 0.558),
    "lm": (0.03011, 0.662),
    "rm": (0.03011, 0.662),
    "mm": (0.03030, 0.623),
}

# PIL's basic layout accumulates per-glyph advance rounding that PowerPoint's
# unhinted fractional layout does not, so some sizes end up a couple of px
# short over a long line (worst at 12pt bold, ~0.08 px/char). These per-gap
# letter-spacing corrections (1/100 pt, positive = widen) were measured
# against the PIL reference on a probe deck; sizes not listed drift <0.02
# px/char (within measurement noise) and are left alone.
_TEXT_SPC = {
    (11, False): 3, (11, True): 7,
    (12, False): 3, (12, True): 10,
    (13, True): 1,
    (17, True): 2,
    (20, False): 4, (20, True): 8,
    (46, False): -2, (46, True): 10,
}

# The baked PIL shadow: rounded rect offset (+1,+5) logical px, colour
# (15,23,42) at alpha 22/255 -- an outer shadow with the same geometry.
_SHADOW_XML = (
    '<a:effectLst %s><a:outerShdw blurRad="0" dist="%d" dir="%d" '
    'algn="ctr" rotWithShape="0"><a:srgbClr val="0F171A">'
    '<a:alpha val="8627"/></a:srgbClr></a:outerShdw></a:effectLst>'
    % (nsdecls("a"), round(hypot(1, 5) * EMU_PX),
       round(degrees(atan2(5, 1)) * 60000)))


def _ex(v):
    return int(round(v * EMU_PX))


def _ey(v):
    return int(round(EMU_TOP + v * EMU_PX))


def _ed(v):
    return int(round(v * EMU_PX))


def _ept(px_size):
    """Logical canvas px -> PowerPoint points (1376 px = 16256000 EMU)."""
    return round(px_size * EMU_PX / 12700.0, 2)


def _alpha_of(color):
    if color is not None and len(color) > 3:
        return round(color[3] / 255.0 * 100000)
    return None


def _fill_shape(shape, color):
    if color is None:
        shape.fill.background()
        return
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(color[0], color[1], color[2])
    alpha = _alpha_of(color)
    if alpha is not None:
        srgb = shape._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        srgb.append(parse_xml('<a:alpha %s val="%d"/>' % (nsdecls("a"), alpha)))


def _line_shape(shape, color, width_px):
    if color is None:
        shape.line.fill.background()
        return
    shape.line.color.rgb = RGBColor(color[0], color[1], color[2])
    shape.line.width = Emu(_ed(width_px))
    alpha = _alpha_of(color)
    if alpha is not None:
        ln = shape._element.spPr.find(qn("a:ln"))
        srgb = ln.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        srgb.append(parse_xml('<a:alpha %s val="%d"/>' % (nsdecls("a"), alpha)))


def _shadow_shape(shape):
    shape._element.spPr.append(parse_xml(_SHADOW_XML))


def _strip_style(shape):
    """python-pptx stamps <p:style> (theme fill/line/effect references) on
    every autoshape it creates. Its effectRef idx=2 paints a soft theme
    shadow around every bar/chip/oval -- visible through translucent fills
    as a grey wash -- so drop the block and let only our explicit fill/line/
    effect XML render."""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def _emit_rrect(slide, op):
    x0, y0, x1, y1 = op["xy"]
    w, h = x1 - x0, y1 - y0
    # PIL strokes a rectangle's outline INSIDE its bbox; PowerPoint centers
    # the line on the path, so inset the shape by half the line width to
    # put the outer edge back on x1/y1.
    ow = op["width"] if op["outline"] is not None else 0
    half = ow / 2.0
    if half > 0 and w > ow and h > ow:
        x0 += half
        y0 += half
        w -= ow
        h -= ow
    # PIL's inclusive right/bottom edge pixel covers half a logical px of
    # the 2x reference canvas; a PowerPoint shape covers [x, x+w) only, so
    # extend by half a px to land the far edge on the reference pixel.
    w += 0.5
    h += 0.5
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                _ex(x0), _ey(y0), _ed(w), _ed(h))
    _strip_style(sh)
    try:
        sh.adjustments[0] = min(op["r"] / min(w, h), 0.5)
    except ZeroDivisionError:
        pass
    _fill_shape(sh, op["fill"])
    _line_shape(sh, op["outline"], op["width"])
    if op["shadow"]:
        _shadow_shape(sh)


def _emit_rect(slide, op):
    x0, y0, x1, y1 = op["xy"]
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                _ex(x0), _ey(y0),
                                _ed(x1 - x0 + 0.5), _ed(y1 - y0 + 0.5))
    _strip_style(sh)
    _fill_shape(sh, op["fill"])
    _line_shape(sh, None, 0)


def _emit_circle(slide, op):
    r = op["r"]
    # Same inside-vs-straddling stroke rule as _emit_rrect: PIL's ellipse
    # outline sits inside the bbox, PowerPoint's line straddles the path.
    ow = op["width"] if op["outline"] is not None else 0
    half = ow / 2.0
    if half > 0 and 2 * r > ow:
        r -= half
    # Same half-px inclusive-edge extension as the rects (reference canvas
    # is 2x): keeps the left/top edge and pulls the right/bottom pixel in.
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                _ex(op["cx"] - r), _ey(op["cy"] - r),
                                _ed(2 * r + 0.5), _ed(2 * r + 0.5))
    _strip_style(sh)
    _fill_shape(sh, op["fill"])
    _line_shape(sh, op["outline"], op["width"])


def _emit_line(slide, op):
    x0, y0, x1, y1 = op["xy"]
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      _ex(x0), _ey(y0), _ex(x1), _ey(y1))
    _strip_style(conn)
    _line_shape(conn, op["fill"], op["width"])


def _emit_arc(slide, op):
    """PIL arc -> DONUT (full ring) / BLOCK_ARC + rotation (segment).
    PIL strokes run from the path inward, matching these presets' outer
    edge; angles are 0=3 o'clock, clockwise (same as OOXML), and negative
    OOXML angles get clamped, so the start is always normalised to [0,360)
    and carried by shape.rotation."""
    cx, cy, r, width = op["cx"], op["cy"], op["r"], op["width"]
    a0, a1 = op["a0"], op["a1"]
    span = a1 - a0
    x, y, d = _ex(cx - r), _ey(cy - r), _ed(2 * r)
    if span >= 360 or span <= 0:
        sh = slide.shapes.add_shape(MSO_SHAPE.DONUT, x, y, d, d)
        _strip_style(sh)
        sh.adjustments[0] = min(width / (2.0 * r), 0.5)
    else:
        sh = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, x, y, d, d)
        _strip_style(sh)
        sh.adjustments[0] = 0.0
        sh.adjustments[1] = round(span * 60000 / 100000.0, 5)
        sh.adjustments[2] = min(width / (2.0 * r), 0.5)
        sh.rotation = float(a0 % 360)
    _fill_shape(sh, op["fill"])
    _line_shape(sh, None, 0)


def _emit_text(slide, op):
    (x, y) = op["xy"]
    ha, va = op["anchor"][0], op["anchor"][1]
    w, lh = op["w"], op["lh"]
    n = max(1, op["lines"])
    # PIL anchor semantics: horizontal uses the advance width, vertical "a"
    # puts the line-box top on y and "m" puts the line-box centre on y.
    left = {"l": x, "m": x - w / 2.0, "r": x - w}[ha]
    k, b = _TEXT_NUDGE.get(op["anchor"], (0.0, 0.0))
    size0 = op["runs"][0][1][0] if op["runs"] else 14
    top = (y if va == "a" else y - lh / 2.0) + k * size0 + b
    tb = slide.shapes.add_textbox(_ex(left), _ey(top), _ed(w), _ed(lh * n))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP if va == "a" else MSO_ANCHOR.MIDDLE

    # Newlines become paragraphs; word runs (banner) become one paragraph
    # with a space inserted between words, exactly as PIL laid them out.
    paras = [[]]
    for s, meta, fill in op["runs"]:
        for i, part in enumerate(str(s).split("\n")):
            if i:
                paras.append([])
            paras[-1].append((part, meta, fill))
    for pi, runs in enumerate(paras):
        p = tf.paragraphs[0] if pi == 0 else tf.add_paragraph()
        p.alignment = _ALIGN[ha]
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 1.0
        first = True
        for s, meta, fill in runs:
            if not s:
                continue
            text = s if first else " " + s
            first = False
            size, bold, family = meta
            run = p.add_run()
            run.text = text
            run.font.size = Pt(_ept(size))
            run.font.bold = bool(bold)
            run.font.name = family
            run.font.color.rgb = RGBColor(fill[0], fill[1], fill[2])
            # PIL's basic layout neither kerns nor subpixel-positions, so
            # push PowerPoint's kern threshold out of the way (kern is the
            # minimum size at which kern pairs apply) to keep line widths
            # as close to the reference raster as possible.
            rPr = run._r.get_or_add_rPr()
            rPr.set("kern", "99900")
            spc = op.get("spc")
            if spc is None:
                spc = _TEXT_SPC.get((int(round(size)), bool(bold)))
            if spc:
                rPr.set("spc", str(spc))
            t = run._r.find(qn("a:t"))
            if t is not None:
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _logo_png_bytes():
    if "png" in _LOGO_CACHE:
        return _LOGO_CACHE["png"]
    logo = _logo()
    _LOGO_CACHE["png"] = _png_bytes(logo) if logo is not None else None
    return _LOGO_CACHE["png"]


def _emit_logo(slide, op):
    data = _logo_png_bytes()
    if data is None:
        return
    x, y, w, h = op["xy"]
    slide.shapes.add_picture(io.BytesIO(data), _ex(x), _ey(y), _ed(w), _ed(h))


def _grid_bg_bytes(bg):
    """One cached bg+grid PNG per background colour (decorative layer)."""
    if bg not in _GRID_BG_CACHE:
        img = Image.new("RGB", (_px(W), _px(H)), bg)
        _grid_paint(ImageDraw.Draw(img, "RGBA"))
        _GRID_BG_CACHE[bg] = _png_bytes(img)
    return _GRID_BG_CACHE[bg]


_GRID_BG_CACHE = {}

_EMITTERS = {
    "rrect": _emit_rrect,
    "rect": _emit_rect,
    "circle": _emit_circle,
    "line": _emit_line,
    "arc": _emit_arc,
    "text": _emit_text,
    "logo": _emit_logo,
}


def _emit_slide(slide, c):
    """Replay a Canvas' recorded ops onto a blank slide, in draw order."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(c.bg[0], c.bg[1], c.bg[2])
    if c.grid:
        slide.shapes.add_picture(io.BytesIO(_grid_bg_bytes(c.bg)),
                                 _ex(0), _ey(0), _ed(W), _ed(H))
    for op in c.rec or []:
        emit = _EMITTERS.get(op["op"])
        if emit is not None:
            emit(slide, op)


def build_development_deck(report):
    """report (db.fetch_development_report_data) -> BytesIO holding the .pptx."""
    report = dict(report or {})
    report["module_count"] = sum(len(c.get("modules") or []) for c in report.get("clients") or [])
    report["milestone_count"] = sum(len(c.get("milestones") or []) for c in report.get("clients") or [])
    pcts = [c.get("overall_pct") for c in report.get("clients") or []
            if c.get("overall_pct") is not None]
    report["avg_pct"] = round(sum(pcts) / len(pcts), 1) if pcts else None

    # The Gantt is split by Title: one slide per Title (chunked only when a
    # Title outgrows TIMELINE_ROWS_PER_PAGE), all on the same axis (see
    # _timeline_slides); an empty report still keeps its slide, showing the
    # "no data" state.
    tl_slides, tl_axis, tl_bounds = _timeline_slides(report)
    slides = [("cover", report), ("exec", report)]
    if tl_slides:
        slides += [("timeline", {**s, "report": report, "axis": tl_axis,
                                 "bounds": tl_bounds})
                   for s in tl_slides]
    else:
        slides.append(("timeline", {"report": report, "rows": [], "title": None,
                                    "client": None, "projek_name": None,
                                    "axis": None, "bounds": None,
                                    "page": 1, "pages": 1}))
    # Each client gets its Status Modul slide immediately followed by that
    # client's milestone slide, so the two stay together when paginating.
    for cl in report.get("clients") or []:
        slides.append(("client", {"report": report, "client": cl}))
        slides.append(("milestones", {"report": report, "client": cl}))
    slides.append(("summary", report))

    prs = Presentation()
    prs.slide_width = Emu(16256000)
    prs.slide_height = Emu(9144000)
    blank = prs.slide_layouts[6]
    for kind, data in slides:
        c = _slide_canvas(kind, data, rec=[])
        slide = prs.slides.add_slide(blank)
        _emit_slide(slide, c)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
