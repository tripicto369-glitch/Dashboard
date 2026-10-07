"""Raster icons for the LHOS dashboard (transparent PNG, outline style).

The dashboard is an Excel workbook, so icons are embedded as pictures. They are
drawn here from simple vector geometry and rasterised with Pillow only (no SVG
rasteriser is assumed to be installed):

1. every shape is described in the icon's own pixel coordinates (a 64 px nav
   icon is designed on a 64 x 64 grid, a 48 px UI icon on 48 x 48, ...);
2. shapes are rasterised on a canvas ``SUPERSAMPLE`` times larger; strokes get
   round caps and round joins (quads between consecutive points + a disc on every
   vertex, i.e. the Minkowski sum of the centre line and a disc);
3. each colour layer is reduced back to the target size with a box filter,
   which yields the exact area coverage of every pixel (proper anti-aliasing
   without the ringing a Lanczos kernel adds around hard edges);
4. layers are composited, and fully transparent pixels get the icon colour as
   their RGB so that a viewer that scales without premultiplied alpha (Excel
   shows the 64 px nav icons at 50 %) does not produce dark fringes.

Style follows the reference screenshot: thin rounded light-grey strokes,
~3.5 px at 64 px (about 1.75 px once Excel shows the icon at 32 px). The
visual weight of the 48 px UI icons (3 px stroke, shown at 24 px) matches.

Files (``assets/icons/<name>.png``):

* ``nav_<key>_normal`` / ``nav_<key>_active`` (64 x 64) for the left menu,
  keys in ``NAV_KEYS`` order: dashboard, sikn, trends, alarms, reports, settings;
* ``ui_import``, ``ui_logout``, ``ui_check_circle``, ``ui_doc``, ``ui_shield``,
  ``ui_calendar`` (48 x 48);
* ``logo_mark`` (96 x 96): a generic droplet with a flame inside (deliberately
  not modelled on any real company logo);
* ``status_ok``, ``status_warn``, ``status_alarm`` (96 x 96): large shields
  for the "overall state" panel (green check / amber and red exclamation);
  an addition to the phase-1 list, matching the screenshot's "НОРМА" badge.

All of them are meant to be inserted at 50 % (x_scale = y_scale = 0.5) so they
stay sharp on high-DPI screens; ``render_icon(name, size=...)`` re-renders the
vector geometry at any other size. ``tests/test_icons.py`` checks geometry,
colours, centring and that the PNGs in ``assets/icons`` are up to date.

Run ``python3 src/lhos/icons.py`` to regenerate everything; ``--sheet PATH``
additionally writes a contact sheet (icons on the dashboard background, scaled
up, plus a preview at the size Excel displays them) for visual review.
"""

from __future__ import annotations

import argparse
import math
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass
from functools import partial
from itertools import pairwise
from pathlib import Path

from PIL import Image, ImageColor, ImageDraw, ImageFont

# --------------------------------------------------------------------- palette
NAV_NORMAL = "#C9D3D9"   # menu icon, idle
NAV_ACTIVE = "#FFFFFF"   # menu icon of the current page
UI_LIGHT = "#C9D3D9"
UI_WHITE = "#FFFFFF"
UI_MUTED = "#8FA0AA"
OK_GREEN = "#4CC35A"
WARN_YELLOW = "#F5B70F"
ALARM_RED = "#F04B46"
LOGO_GREEN = "#2FA84F"
LOGO_RED = "#E53935"
PAGE_BG = "#04121A"      # dashboard page colour (used for the contact sheet)

NAV_SIZE = 64
UI_SIZE = 48
BIG_SIZE = 96

NAV_STROKE = 3.5         # at 64 px
UI_STROKE = 3.0          # at 48 px
BIG_STROKE = 5.0         # at 96 px

SUPERSAMPLE = 16         # rasterisation scale before the box-filter reduction

# Left-menu entries, top to bottom; files are nav_<key>_normal/_active.png.
NAV_KEYS = ("dashboard", "sikn", "trends", "alarms", "reports", "settings")

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_OUT_DIR = REPO_ROOT / "assets" / "icons"

Point = tuple[float, float]


# -------------------------------------------------------------------- geometry
def polar(cx: float, cy: float, r: float, deg: float) -> Point:
    """Point on a circle; ``deg`` is counter-clockwise from +x (y axis up)."""
    a = math.radians(deg)
    return (cx + r * math.cos(a), cy - r * math.sin(a))


class Shape:
    """A list of polylines built with move/line/arc/cubic commands.

    Curves are flattened immediately into short segments (about half a pixel
    of the final image each), which is plenty for supersampled rasterisation.
    """

    def __init__(self) -> None:
        self.subpaths: list[tuple[list[Point], bool]] = []

    # -- building ----------------------------------------------------------
    @property
    def _pts(self) -> list[Point]:
        if not self.subpaths:
            raise ValueError("path has no current point; call move() first")
        return self.subpaths[-1][0]

    def move(self, x: float, y: float) -> Shape:
        self.subpaths.append(([(x, y)], False))
        return self

    def line(self, x: float, y: float) -> Shape:
        self._pts.append((x, y))
        return self

    def lines(self, *pts: Point) -> Shape:
        for p in pts:
            self.line(*p)
        return self

    def arc(self, cx: float, cy: float, r: float, a0: float, a1: float,
            *, connect: bool = True) -> Shape:
        """Circular arc from angle ``a0`` to ``a1`` (degrees, CCW, y up).

        With ``connect`` the arc continues the current polyline (a straight
        segment joins the current point to the arc start); otherwise it
        starts a new subpath.
        """
        steps = max(4, math.ceil(abs(math.radians(a1 - a0)) * r * 2))
        pts = [polar(cx, cy, r, a0 + (a1 - a0) * i / steps) for i in range(steps + 1)]
        if connect and self.subpaths:
            self._pts.extend(pts)
        else:
            self.subpaths.append((pts, False))
        return self

    def cubic(self, c1: Point, c2: Point, p: Point) -> Shape:
        """Cubic Bezier from the current point through controls c1, c2 to p."""
        p0 = self._pts[-1]
        approx = (math.dist(p0, c1) + math.dist(c1, c2) + math.dist(c2, p))
        steps = max(4, math.ceil(approx * 2))
        for i in range(1, steps + 1):
            t = i / steps
            u = 1 - t
            self._pts.append((
                u ** 3 * p0[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t ** 3 * p[0],
                u ** 3 * p0[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t ** 3 * p[1],
            ))
        return self

    def close(self) -> Shape:
        pts, _ = self.subpaths[-1]
        self.subpaths[-1] = (pts, True)
        return self

    def extend(self, other: Shape) -> Shape:
        self.subpaths.extend(other.subpaths)
        return self


def polyline(*pts: Point, closed: bool = False) -> Shape:
    p = Shape().move(*pts[0]).lines(*pts[1:])
    return p.close() if closed else p


def circle(cx: float, cy: float, r: float) -> Shape:
    return Shape().arc(cx, cy, r, 0, 360, connect=False).close()


def rounded_rect(x0: float, y0: float, x1: float, y1: float, r: float) -> Shape:
    """Closed rectangle outline with corner radius ``r`` (centre-line coords)."""
    p = Shape().move(x0 + r, y0)
    p.arc(x1 - r, y0 + r, r, 90, 0)
    p.arc(x1 - r, y1 - r, r, 0, -90)
    p.arc(x0 + r, y1 - r, r, 270, 180)
    p.arc(x0 + r, y0 + r, r, 180, 90)
    return p.close()


# -------------------------------------------------------------------- canvas
def _rgb(color: str) -> tuple[int, int, int]:
    """``#RRGGBB`` (or any other colour string Pillow understands, e.g.
    ``#RGB`` or ``white``) -> RGB tuple; raises ValueError for anything else."""
    return ImageColor.getrgb(color)[:3]


@dataclass
class _Layer:
    color: str
    opacity: float
    mask: Image.Image   # "L", supersampled


class IconCanvas:
    """Supersampled drawing surface addressed in design-grid coordinates.

    ``design`` is the size of the grid the icon is drawn on (64 for nav
    icons); ``size`` is the output size in pixels (defaults to ``design``),
    so the same drawing code renders crisply at any resolution.

    Consecutive operations with the same colour share one coverage mask, so
    overlapping strokes of one colour never double up their alpha.
    """

    def __init__(self, design: int, size: int | None = None, ss: int = SUPERSAMPLE) -> None:
        self.size = size or design
        self.ss = ss
        self.k = self.size * ss / design      # design unit -> supersampled pixels
        self.layers: list[_Layer] = []

    def _mask(self, color: str, opacity: float) -> ImageDraw.ImageDraw:
        if not self.layers or (self.layers[-1].color, self.layers[-1].opacity) != (color, opacity):
            big = self.size * self.ss
            self.layers.append(_Layer(color, opacity, Image.new("L", (big, big), 0)))
        return ImageDraw.Draw(self.layers[-1].mask)

    def _scaled(self, pts: Iterable[Point]) -> list[Point]:
        # Pillow fills a polygon/ellipse including the pixels its edges pass
        # through on both sides, i.e. one supersampled pixel too wide. The
        # half-pixel shift centres that bias and strokes shrink their
        # half-width by 0.5 (see _half_width), so stroke edges land exactly;
        # filled polygons stay within 1/SUPERSAMPLE px of the ideal shape.
        k = self.k
        return [(x * k - 0.5, y * k - 0.5) for x, y in pts]

    def _half_width(self, width: float) -> float:
        return max(0.5, width * self.k / 2 - 0.5)

    @staticmethod
    def _stroke_into(draw: ImageDraw.ImageDraw, pts: Sequence[Point], closed: bool,
                     hw: float, value: int) -> None:
        """Round-capped, round-joined thick polyline (already scaled)."""
        segs = list(pairwise(pts))
        if closed and len(pts) > 2:
            segs.append((pts[-1], pts[0]))
        for (x0, y0), (x1, y1) in segs:
            length = math.hypot(x1 - x0, y1 - y0)
            if length < 1e-9:
                continue
            nx, ny = -(y1 - y0) / length * hw, (x1 - x0) / length * hw
            draw.polygon([(x0 + nx, y0 + ny), (x1 + nx, y1 + ny),
                          (x1 - nx, y1 - ny), (x0 - nx, y0 - ny)], fill=value)
        for x, y in pts:
            draw.ellipse((x - hw, y - hw, x + hw, y + hw), fill=value)

    # -- drawing ops ---------------------------------------------------------
    def stroke(self, shape: Shape, color: str, width: float, opacity: float = 1.0) -> None:
        """Outline ``shape`` with round caps/joins; ``width`` in design units."""
        draw = self._mask(color, opacity)
        hw = self._half_width(width)
        for pts, closed in shape.subpaths:
            self._stroke_into(draw, self._scaled(pts), closed, hw, 255)

    def fill(self, shape: Shape, color: str, opacity: float = 1.0) -> None:
        """Fill every subpath of ``shape`` (each one as a simple polygon)."""
        draw = self._mask(color, opacity)
        for pts, _ in shape.subpaths:
            draw.polygon(self._scaled(pts), fill=255)

    def dot(self, x: float, y: float, r: float, color: str, opacity: float = 1.0) -> None:
        self.fill(circle(x, y, r), color, opacity)

    def erase(self, shape: Shape, *, grow: float = 0.0) -> None:
        """Knock ``shape`` (dilated by ``grow`` design units) out of all layers."""
        for layer in self.layers:
            draw = ImageDraw.Draw(layer.mask)
            for pts, _ in shape.subpaths:
                scaled = self._scaled(pts)
                draw.polygon(scaled, fill=0)
                if grow > 0:
                    self._stroke_into(draw, scaled, True, grow * self.k + 0.5, 0)

    # -- output --------------------------------------------------------------
    def render(self) -> Image.Image:
        size = (self.size, self.size)
        out = Image.new("RGBA", size, (0, 0, 0, 0))
        for layer in self.layers:
            alpha = layer.mask.reduce(self.ss)          # box filter = exact coverage
            if layer.opacity < 1.0:
                alpha = alpha.point(lambda v, k=layer.opacity: round(v * k))
            solid = Image.new("RGBA", size, _rgb(layer.color) + (0,))
            solid.putalpha(alpha)
            out = Image.alpha_composite(out, solid)
        # Colour-bleed fully transparent pixels (see module docstring, step 4).
        alpha = out.getchannel("A")
        base = _rgb(self.layers[0].color) if self.layers else (0, 0, 0)
        covered = alpha.point(lambda v: 255 if v else 0)
        out = Image.composite(out, Image.new("RGBA", size, base + (255,)), covered)
        out.putalpha(alpha)
        return out


# ------------------------------------------------------------ nav icons (64)
# All coordinates below are pixels of the final 64 x 64 image. Axis-aligned
# strokes sit on odd coordinates where possible, which puts them on pixel
# centres once Excel shows the icon at 32 px.

def _nav_dashboard(c: IconCanvas, col: str) -> None:
    """Speedometer: 270-degree bezel, five ticks, needle and hub."""
    w, cx, cy, r = NAV_STROKE, 32.0, 35.0, 24.0
    p = Shape().move(*polar(cx, cy, r - 5.5, 225)).arc(cx, cy, r, 225, -45)
    p.line(*polar(cx, cy, r - 5.5, -45))           # bezel ends turn inwards
    c.stroke(p, col, w)
    for a in (180, 135, 90, 45, 0):
        c.stroke(polyline(polar(cx, cy, 14.0, a), polar(cx, cy, 17.0, a)), col, w)
    # Hub ring: its hole (~4.9 px here) stays open at the 32 px display size.
    c.stroke(circle(cx, cy, 4.2), col, w)
    # Needle between the 45 and 90 degree ticks with >= 1.5 px of clear space
    # to both, so it stays a separate shape (an earlier needle at 58 degrees /
    # r = 13 ran into the 45 degree tick and read as one bent squiggle).
    c.stroke(polyline(polar(cx, cy, 6.2, 65), polar(cx, cy, 11.5, 65)), col, w)


def _nav_sikn(c: IconCanvas, col: str) -> None:
    """Metering line: pipe with end flanges and a gate valve with handwheel."""
    w = NAV_STROKE
    c.stroke(polyline((21, 9), (43, 9)), col, w)                 # handwheel
    c.stroke(polyline((32, 9), (32, 21)), col, w)                # stem
    c.stroke(rounded_rect(24, 21, 40, 53, 3), col, w)            # valve body
    for x0, x1 in ((15, 24), (40, 49)):                          # pipe walls
        c.stroke(polyline((x0, 31), (x1, 31)), col, w)
        c.stroke(polyline((x0, 43), (x1, 43)), col, w)
    c.stroke(rounded_rect(8, 25, 15, 49, 2), col, w)             # flanges
    c.stroke(rounded_rect(49, 25, 56, 49, 2), col, w)


def _nav_trends(c: IconCanvas, col: str) -> None:
    """Line chart with axes and a rising arrow."""
    w = NAV_STROKE
    c.stroke(polyline((9, 8), (9, 55), (56, 55)), col, w)
    c.stroke(polyline((17, 45), (28, 33), (37, 41), (53, 21)), col, w)
    c.stroke(polyline((44, 21), (53, 21), (53, 30)), col, w)


def _nav_alarms(c: IconCanvas, col: str) -> None:
    """Bell with a flared rim and a clapper."""
    w = NAV_STROKE
    p = Shape().move(10, 49)
    p.cubic((14, 47), (17, 44), (17, 38))
    p.line(17, 29)
    p.arc(32, 29, 15, 180, 0)
    p.line(47, 38)
    p.cubic((47, 44), (50, 47), (54, 49))
    p.close()
    c.stroke(p, col, w)
    c.stroke(polyline((32, 9), (32, 13)), col, w)               # hanger
    c.stroke(Shape().arc(32, 51, 5, 180, 360, connect=False), col, w)   # clapper


def _nav_reports(c: IconCanvas, col: str) -> None:
    """Document with a folded corner and text lines."""
    w = NAV_STROKE
    c.stroke(polyline((15, 7), (38, 7), (49, 18), (49, 57), (15, 57), closed=True), col, w)
    fold = polyline((38, 7), (38, 18), (49, 18), closed=True)
    c.stroke(fold, col, w)
    c.fill(fold, col)
    c.stroke(polyline((23, 29), (41, 29)), col, w)
    c.stroke(polyline((23, 38), (41, 38)), col, w)
    c.stroke(polyline((23, 47), (34, 47)), col, w)


def _gear(cx: float, cy: float, teeth: int, r_tip: float, r_root: float,
          half_tip: float, half_root: float, phase: float = 90.0) -> Shape:
    """Closed gear outline: flat tooth tips, tapered flanks, arcs between."""
    step = 360.0 / teeth
    p = Shape()
    for i in range(teeth):
        a = phase + i * step
        if i == 0:
            p.move(*polar(cx, cy, r_root, a - step + half_root))
        p.arc(cx, cy, r_root, a - step + half_root, a - half_root)
        p.line(*polar(cx, cy, r_tip, a - half_tip))
        p.arc(cx, cy, r_tip, a - half_tip, a + half_tip)
        p.line(*polar(cx, cy, r_root, a + half_root))
    return p.close()


def _nav_settings(c: IconCanvas, col: str) -> None:
    """Eight-tooth gear with a round hub."""
    w = NAV_STROKE
    c.stroke(_gear(32, 32, 8, 25.5, 19.5, 8.5, 12.5), col, w)
    c.stroke(circle(32, 32, 7.5), col, w)


_NAV_DRAW: dict[str, Callable[[IconCanvas, str], None]] = {
    "dashboard": _nav_dashboard,
    "sikn": _nav_sikn,
    "trends": _nav_trends,
    "alarms": _nav_alarms,
    "reports": _nav_reports,
    "settings": _nav_settings,
}


# ------------------------------------------------------------- UI icons (48)
def _ui_import(c: IconCanvas, col: str) -> None:
    """Download arrow dropping into a tray."""
    w = UI_STROKE
    c.stroke(polyline((24, 6), (24, 30)), col, w)
    c.stroke(polyline((15, 21), (24, 30), (33, 21)), col, w)
    c.stroke(polyline((7, 30), (7, 41), (41, 41), (41, 30)), col, w)


def _ui_logout(c: IconCanvas, col: str) -> None:
    """Exit: door frame open to the right, arrow leaving it."""
    w = UI_STROKE
    c.stroke(polyline((20, 7), (8, 7), (8, 41), (20, 41)), col, w)
    c.stroke(polyline((19, 24), (40, 24)), col, w)
    c.stroke(polyline((32, 16), (40, 24), (32, 32)), col, w)


def _ui_check_circle(c: IconCanvas, col: str) -> None:
    """"All systems OK" mark: check inside a circle."""
    w = UI_STROKE
    c.stroke(circle(24, 24, 19), col, w)
    c.stroke(polyline((15, 24.5), (21, 30.5), (33, 18.5)), col, w)


def _ui_doc(c: IconCanvas, col: str) -> None:
    """Small document with a folded corner and three lines (like nav_reports)."""
    w = UI_STROKE
    c.stroke(polyline((11, 5), (29, 5), (37, 13), (37, 43), (11, 43), closed=True), col, w)
    fold = polyline((29, 5), (29, 13), (37, 13), closed=True)
    c.stroke(fold, col, w)
    c.fill(fold, col)
    c.stroke(polyline((17, 22), (31, 22)), col, w)
    c.stroke(polyline((17, 29), (31, 29)), col, w)
    c.stroke(polyline((17, 36), (25, 36)), col, w)


def _shield(cx: float, top: float, half_w: float, height: float) -> Shape:
    """Classic shield: flat shoulders, straight sides, curved point."""
    x0, x1 = cx - half_w, cx + half_w
    shoulder = top + height * 0.13
    side_end = top + height * 0.45
    bottom = top + height
    p = Shape().move(cx, top)
    p.line(x1, shoulder).line(x1, side_end)
    p.cubic((x1, top + height * 0.72), (cx + half_w * 0.45, bottom - height * 0.09), (cx, bottom))
    p.cubic((cx - half_w * 0.45, bottom - height * 0.09), (x0, top + height * 0.72), (x0, side_end))
    p.line(x0, shoulder)
    return p.close()


def _ui_shield(c: IconCanvas, col: str) -> None:
    """Access level: shield outline with a check."""
    w = UI_STROKE
    c.stroke(_shield(24, 5, 15, 38), col, w)
    c.stroke(polyline((17, 23.5), (22, 28.5), (31, 19.5)), col, w)


def _ui_calendar(c: IconCanvas, col: str) -> None:
    """Calendar page with binder rings and a few day marks."""
    w = UI_STROKE
    c.stroke(rounded_rect(7, 10, 41, 41, 3.5), col, w)
    c.stroke(polyline((7, 19), (41, 19)), col, w)
    c.stroke(polyline((16, 6), (16, 13)), col, w)
    c.stroke(polyline((32, 6), (32, 13)), col, w)
    for x in (16, 24, 32):
        for y in (27, 34):
            c.dot(x, y, 2.2, col)


# --------------------------------------------------------- large marks (96)
def _flame(tip: Point, cx: float, half_w: float, mid: float, base: float,
           right: Point, left: Point) -> Shape:
    """Flame silhouette: round bottom centred on ``cx``, pointed ``tip``.

    ``right``/``left`` are the control-point offsets (from the tip) that shape
    the flanks next to the tip; asymmetric values make the flame lean.
    """
    tx, ty = tip
    k = 0.55                                    # bottom roundness
    f = Shape().move(tx, ty)
    f.cubic((tx + right[0], ty + right[1]), (cx + half_w, mid - (mid - ty) * 0.4),
            (cx + half_w, mid))
    f.cubic((cx + half_w, base - (base - mid) * 0.5), (cx + half_w * k, base), (cx, base))
    f.cubic((cx - half_w * k, base), (cx - half_w, base - (base - mid) * 0.5), (cx - half_w, mid))
    f.cubic((cx - half_w, mid - (mid - ty) * 0.37), (tx + left[0], ty + left[1]), (tx, ty))
    return f.close()


def _logo_mark(c: IconCanvas, _col: str) -> None:
    """Generic oil-and-gas mark: green droplet holding a leaning red flame.

    The flame is separated from the droplet by a transparent gap and has a
    transparent core, so the mark reads on any dark background.
    """
    drop = Shape().move(48, 5)
    drop.cubic((54, 18), (79, 38), (79, 60))
    drop.arc(48, 60, 31, 0, -180)
    drop.cubic((17, 38), (42, 18), (48, 5))
    c.fill(drop.close(), LOGO_GREEN)

    flame = _flame((56, 30), 48, 17, 63, 81, right=(-1, 10), left=(-10, 12))
    c.erase(flame, grow=3.5)
    c.fill(flame, LOGO_RED)
    c.erase(_flame((51, 53), 48, 7.5, 67, 76, right=(2, 5), left=(-4, 6)))


def _status_shield(c: IconCanvas, col: str, mark: str) -> None:
    """Overall-state badge: tinted shield with a check or an exclamation mark."""
    w = BIG_STROKE
    body = _shield(48, 10, 31, 76)
    c.fill(body, col, opacity=0.22)
    c.stroke(body, col, w)
    if mark == "check":
        c.stroke(polyline((34, 47), (44, 57), (63, 38)), col, w + 0.5)
    else:  # exclamation mark
        c.stroke(polyline((48, 30), (48, 52)), col, w + 0.5)
        c.dot(48, 63, 3.8, col)


# ------------------------------------------------------------------ registry
@dataclass(frozen=True)
class IconSpec:
    """A registered icon: output file stem, pixel size, main colour, painter.

    ``draw(canvas, color)`` paints on a canvas whose design grid equals
    ``size``; multi-colour marks (logo) ignore ``color`` for their accents.
    """

    name: str
    size: int
    color: str
    draw: Callable[[IconCanvas, str], None]


def _registry() -> dict[str, IconSpec]:
    specs: list[IconSpec] = []
    for key in NAV_KEYS:
        specs.append(IconSpec(f"nav_{key}_normal", NAV_SIZE, NAV_NORMAL, _NAV_DRAW[key]))
        specs.append(IconSpec(f"nav_{key}_active", NAV_SIZE, NAV_ACTIVE, _NAV_DRAW[key]))
    specs += [
        IconSpec("ui_import", UI_SIZE, UI_WHITE, _ui_import),
        IconSpec("ui_logout", UI_SIZE, UI_LIGHT, _ui_logout),
        IconSpec("ui_check_circle", UI_SIZE, OK_GREEN, _ui_check_circle),
        IconSpec("ui_doc", UI_SIZE, UI_MUTED, _ui_doc),
        IconSpec("ui_shield", UI_SIZE, UI_LIGHT, _ui_shield),
        IconSpec("ui_calendar", UI_SIZE, UI_MUTED, _ui_calendar),
        IconSpec("logo_mark", BIG_SIZE, LOGO_GREEN, _logo_mark),
        IconSpec("status_ok", BIG_SIZE, OK_GREEN, partial(_status_shield, mark="check")),
        IconSpec("status_warn", BIG_SIZE, WARN_YELLOW, partial(_status_shield, mark="exclaim")),
        IconSpec("status_alarm", BIG_SIZE, ALARM_RED, partial(_status_shield, mark="exclaim")),
    ]
    return {s.name: s for s in specs}


ICONS: dict[str, IconSpec] = _registry()


def render_icon(name: str, *, color: str | None = None, size: int | None = None) -> Image.Image:
    """Render one registered icon as an RGBA image.

    ``color`` overrides the main colour (all strokes of single-colour icons);
    ``size`` renders at another resolution by scaling the vector geometry,
    stroke widths included, so the result stays sharp.
    """
    spec = ICONS[name]
    canvas = IconCanvas(spec.size, size)
    spec.draw(canvas, color or spec.color)
    return canvas.render()


def icon_path(name: str, out_dir: Path | str = DEFAULT_OUT_DIR) -> Path:
    """Where ``generate_all`` stores icon ``name``."""
    if name not in ICONS:
        raise KeyError(f"unknown icon {name!r}")
    return Path(out_dir) / f"{name}.png"


def generate_all(out_dir: Path | str = DEFAULT_OUT_DIR) -> list[Path]:
    """Write every registered icon as ``<out_dir>/<name>.png``; return paths."""
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    written = []
    for name in ICONS:
        path = out / f"{name}.png"
        render_icon(name).save(path, format="PNG", optimize=True)
        written.append(path)
    return written


# ------------------------------------------------------------- contact sheet
def contact_sheet(path: Path | str, *, scale: int = 3, bg: str = PAGE_BG) -> Path:
    """Compose a review sheet of all icons on the dashboard background.

    Every cell shows, left to right: the icon rendered ``scale`` times larger
    (from the vector geometry); the PNG reduced to the size it is displayed
    at in the workbook (50 %), magnified with nearest-neighbour so single
    pixels are visible; and that display size 1:1 for a realistic impression.
    """
    names = list(ICONS)
    cols = 3
    big_side, disp_max = BIG_SIZE * scale, BIG_SIZE // 2
    cell_w = big_side + 12 + disp_max * scale + 16 + disp_max + 24
    cell_h = big_side + 40
    rows = math.ceil(len(names) / cols)
    sheet = Image.new("RGB", (cols * cell_w, rows * cell_h), _rgb(bg))
    draw = ImageDraw.Draw(sheet)
    font = ImageFont.load_default(size=14)
    for i, name in enumerate(names):
        img = render_icon(name)
        x0, y0 = (i % cols) * cell_w + 12, (i // cols) * cell_h + 8
        draw.text((x0, y0), f"{name}  {img.width}px", fill=(150, 170, 180), font=font)
        y0 += 24
        big = render_icon(name, size=img.width * scale)
        sheet.paste(big, (x0, y0), big)
        disp = img.width // 2
        small = img.resize((disp, disp), Image.Resampling.LANCZOS)
        pixels = small.resize((disp * scale, disp * scale), Image.Resampling.NEAREST)
        x1 = x0 + big_side + 12
        sheet.paste(pixels, (x1, y0), pixels)
        sheet.paste(small, (x1 + pixels.width + 16, y0), small)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(path)
    return path


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Render dashboard icons to PNG.")
    ap.add_argument("--out", default=str(DEFAULT_OUT_DIR), help="output directory")
    ap.add_argument("--sheet", help="also write a contact sheet PNG to this path")
    ap.add_argument("--scale", type=int, default=3, help="contact sheet magnification")
    args = ap.parse_args(argv)
    paths = generate_all(args.out)
    print(f"wrote {len(paths)} icons to {args.out}")
    if args.sheet:
        print(f"contact sheet: {contact_sheet(args.sheet, scale=args.scale)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
