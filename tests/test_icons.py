"""Tests for src/lhos/icons.py (dashboard PNG icons).

Runs under pytest or standalone: ``python3 -I tests/test_icons.py``
(exit code 1 on failure).
"""

from __future__ import annotations

import io
import sys
import tempfile
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from PIL import Image  # noqa: E402

from lhos import icons  # noqa: E402

# Files the dashboard builder relies on (phase-1 spec), with size and colour.
REQUIRED = {
    **{f"nav_{k}_normal": (64, "#C9D3D9") for k in icons.NAV_KEYS},
    **{f"nav_{k}_active": (64, "#FFFFFF") for k in icons.NAV_KEYS},
    "ui_import": (48, "#FFFFFF"),
    "ui_logout": (48, "#C9D3D9"),
    "ui_check_circle": (48, "#4CC35A"),
    "ui_doc": (48, "#8FA0AA"),
    "ui_shield": (48, "#C9D3D9"),
    "ui_calendar": (48, "#8FA0AA"),
    "logo_mark": (96, None),
}
MULTI_COLOUR = {"logo_mark"}


def _rgb(hex_color: str) -> tuple[int, int, int]:
    return icons._rgb(hex_color)


def _pixels(img: Image.Image) -> list[tuple[int, ...]]:
    """All RGBA pixels (``getdata`` is deprecated from Pillow 12.1 on)."""
    flat = getattr(img, "get_flattened_data", None)
    return list(flat() if flat else img.getdata())


def _alpha_bbox(img: Image.Image, threshold: int = 64) -> tuple[int, int, int, int]:
    return img.getchannel("A").point(lambda v: 255 if v > threshold else 0).getbbox()


def _shapes(img: Image.Image, threshold: int) -> int:
    """Number of 8-connected blobs of pixels with alpha > ``threshold``."""
    a = img.getchannel("A")
    w, h = a.size
    px = a.load()
    seen: set[tuple[int, int]] = set()
    count = 0
    for y in range(h):
        for x in range(w):
            if px[x, y] <= threshold or (x, y) in seen:
                continue
            count += 1
            stack = [(x, y)]
            seen.add((x, y))
            while stack:
                cx, cy = stack.pop()
                for dx in (-1, 0, 1):
                    for dy in (-1, 0, 1):
                        q = (cx + dx, cy + dy)
                        if (0 <= q[0] < w and 0 <= q[1] < h and q not in seen
                                and px[q] > threshold):
                            seen.add(q)
                            stack.append(q)
    return count


def test_registry_covers_spec() -> None:
    assert icons.NAV_KEYS == ("dashboard", "sikn", "trends", "alarms", "reports", "settings")
    for name, (size, color) in REQUIRED.items():
        assert name in icons.ICONS, name
        spec = icons.ICONS[name]
        assert spec.size == size, (name, spec.size)
        if color:
            assert spec.color.upper() == color, (name, spec.color)


def test_render_basic_properties() -> None:
    for name, spec in icons.ICONS.items():
        img = icons.render_icon(name)
        assert img.mode == "RGBA" and img.size == (spec.size, spec.size), name
        a = img.getchannel("A")
        w = spec.size - 1
        for corner in ((0, 0), (w, 0), (0, w), (w, w)):
            assert a.getpixel(corner) == 0, (name, corner)
        ink = sum(a.histogram()[v] * v for v in range(256)) / 255
        # Neither empty nor a blob: 5..60 % of the area is covered.
        assert 0.05 < ink / spec.size ** 2 < 0.60, (name, ink)
        # Centred: the visible bounding box is centred within 1 px (at 64 px).
        x0, y0, x1, y1 = _alpha_bbox(img)
        tol = spec.size / 64
        assert abs((x0 + x1) / 2 - spec.size / 2) <= tol, (name, (x0, y0, x1, y1))
        assert abs((y0 + y1) / 2 - spec.size / 2) <= tol, (name, (x0, y0, x1, y1))
        # Some breathing room at the edges (nothing touches the border).
        assert min(x0, y0) >= 2 and max(x1, y1) <= spec.size - 2, (name, (x0, y0, x1, y1))


def test_single_colour_icons_use_exact_colour() -> None:
    for name, spec in icons.ICONS.items():
        if name in MULTI_COLOUR or name.startswith("status_"):
            continue
        img = icons.render_icon(name)
        colours = {px[:3] for px in _pixels(img)}
        # Every pixel, transparent ones included (colour bleed), has the icon colour.
        assert colours == {_rgb(spec.color)}, (name, colours)


def test_logo_colours() -> None:
    img = icons.render_icon("logo_mark")
    opaque = {px[:3] for px in _pixels(img) if px[3] == 255}
    assert _rgb(icons.LOGO_GREEN) in opaque and _rgb(icons.LOGO_RED) in opaque
    # The flame has a transparent gap around it and a transparent core.
    a = img.getchannel("A")
    assert a.getpixel((48, 72)) == 0, "flame core should be transparent"


def test_stroke_width_is_exact() -> None:
    # nav_trends: horizontal x-axis centred on y = 55, NAV_STROKE wide.
    a = icons.render_icon("nav_trends_normal").getchannel("A")
    column = sum(a.getpixel((30, y)) for y in range(45, 64)) / 255
    assert abs(column - icons.NAV_STROKE) < 0.1, column
    # Same geometry rendered at 2x keeps proportions (vector scaling).
    a2 = icons.render_icon("nav_trends_normal", size=128).getchannel("A")
    column2 = sum(a2.getpixel((60, y)) for y in range(90, 128)) / 255
    assert abs(column2 - 2 * icons.NAV_STROKE) < 0.1, column2


def test_gauge_parts_do_not_touch() -> None:
    # nav_dashboard = bezel + 5 ticks + hub-with-needle: 7 separate shapes,
    # both as drawn (64 px) and at the 32 px Excel shows. A needle running
    # into a tick (a previous defect) makes this 6.
    img = icons.render_icon("nav_dashboard_normal")
    assert _shapes(img, 16) == 7
    assert _shapes(img.resize((32, 32), Image.Resampling.BOX), 64) == 7


def test_colour_override_and_determinism() -> None:
    red = icons.render_icon("nav_sikn_normal", color="#FF0000")
    assert red.getpixel((0, 0))[:3] == (255, 0, 0)
    # Short hex works too; garbage is a clear ValueError.
    assert icons.render_icon("ui_doc", color="#0f0").getpixel((0, 0))[:3] == (0, 255, 0)
    try:
        icons.render_icon("ui_doc", color="#12")
    except ValueError:
        pass
    else:
        raise AssertionError("invalid colour accepted")

    def png_bytes() -> bytes:
        buf = io.BytesIO()
        icons.render_icon("nav_settings_normal").save(buf, format="PNG", optimize=True)
        return buf.getvalue()

    assert png_bytes() == png_bytes()


def test_generate_all_and_contact_sheet() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        paths = icons.generate_all(tmp)
        assert {p.stem for p in paths} == set(icons.ICONS)
        for p in paths:
            with Image.open(p) as img:
                assert img.mode == "RGBA" and img.size[0] == icons.ICONS[p.stem].size
        sheet = icons.contact_sheet(Path(tmp) / "sheet" / "contact.png")
        with Image.open(sheet) as img:
            assert img.width > 500 and img.height > 500


def test_committed_assets_are_up_to_date() -> None:
    """assets/icons/*.png must match the current drawing code (regenerate with
    ``python3 src/lhos/icons.py`` after changing it)."""
    out = icons.DEFAULT_OUT_DIR
    missing = [n for n in icons.ICONS if not icons.icon_path(n, out).exists()]
    assert not missing, f"missing icons in {out}: {missing}"
    for name in icons.ICONS:
        with Image.open(icons.icon_path(name, out)) as disk:
            fresh = icons.render_icon(name)
            assert disk.convert("RGBA").tobytes() == fresh.tobytes(), f"{name}.png is stale"


def main() -> int:
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    failed = 0
    for name, fn in tests:
        try:
            fn()
        except Exception:  # noqa: BLE001 - report every failure, keep going
            failed += 1
            print(f"FAIL {name}")
            traceback.print_exc()
        else:
            print(f"ok   {name}")
    print(f"{len(tests) - failed}/{len(tests)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
