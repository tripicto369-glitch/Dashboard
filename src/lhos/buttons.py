"""Картинки кнопок дашборда (PNG с двойным разрешением).

Кнопки в Excel — это картинки с назначенным макросом: так их вид полностью
совпадает с макетом (скругление, значок, шрифт) в любой версии Excel.
Запуск: ``python3 src/lhos/buttons.py`` — пересоздает assets/icons/btn_*.png.
"""

from __future__ import annotations

import os

from PIL import Image, ImageDraw, ImageFont

try:
    from .theme import C
except ImportError:  # запуск как скрипта
    import sys
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from lhos.theme import C

SCALE = 2  # картинка рисуется вдвое крупнее и вставляется с масштабом 50 %

BUTTONS = {
    # имя: (ширина, высота, текст, значок, заливка, рамка, цвет текста, значок справа)
    "import": (184, 34, "ИМПОРТ ДАННЫХ", "ui_import", C["btn"], C["btn_line"], "#FFFFFF", False),
    "clear": (184, 34, "ОЧИСТИТЬ БАЗУ", None, C["btn2"], C["btn2_line"], C["text"], False),
    "exit": (108, 28, "ВЫЙТИ", "ui_logout", C["btn2"], C["btn2_line"], C["text"], True),
}

_FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/inter/Inter-SemiBold.ttf",
    "/usr/share/fonts/opentype/inter/Inter-SemiBold.otf",
    "C:/Windows/Fonts/seguisb.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]


def _font(size: int) -> ImageFont.FreeTypeFont:
    for p in _FONT_CANDIDATES:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    # поиск Inter SemiBold среди установленных шрифтов
    for root in ("/usr/share/fonts", "/usr/local/share/fonts"):
        for dp, _dn, files in os.walk(root):
            for f in files:
                if f.lower().startswith("inter") and "semibold" in f.lower():
                    return ImageFont.truetype(os.path.join(dp, f), size)
    return ImageFont.load_default(size)


def _rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def render_button(out_path: str, icons_dir: str, w: int, h: int, text: str, icon: str | None,
                  fill: str, line: str, color: str, icon_right: bool) -> None:
    s = SCALE
    ss = 4  # суперсэмплинг для гладких скруглений
    W, H = w * s * ss, h * s * ss
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = 5 * s * ss
    d.rounded_rectangle((0, 0, W - 1, H - 1), radius=r, fill=_rgb(fill) + (255,),
                        outline=_rgb(line) + (255,), width=max(1, int(1.2 * s * ss)))
    img = img.resize((w * s, h * s), Image.LANCZOS)
    d = ImageDraw.Draw(img)

    font = _font(int(9.5 * 1.333 * s) if h >= 32 else int(8.5 * 1.333 * s))
    tb = d.textbbox((0, 0), text, font=font)
    tw, th = tb[2] - tb[0], tb[3] - tb[1]
    icon_px = 20 * s if h >= 32 else 16 * s
    gap = 8 * s
    total = tw + (icon_px + gap if icon else 0)
    x = (w * s - total) // 2
    y_text = (h * s - th) // 2 - tb[1]
    if icon and not icon_right:
        ic = Image.open(os.path.join(icons_dir, icon + ".png")).convert("RGBA")
        ic = ic.resize((icon_px, icon_px), Image.LANCZOS)
        img.alpha_composite(ic, (x, (h * s - icon_px) // 2))
        x += icon_px + gap
    d.text((x, y_text), text, font=font, fill=_rgb(color) + (255,))
    if icon and icon_right:
        ic = Image.open(os.path.join(icons_dir, icon + ".png")).convert("RGBA")
        ic = ic.resize((icon_px, icon_px), Image.LANCZOS)
        img.alpha_composite(ic, (x + tw + gap, (h * s - icon_px) // 2))
    img.save(out_path)


def render_all(icons_dir: str, out_dir: str | None = None) -> list[str]:
    out_dir = out_dir or icons_dir
    os.makedirs(out_dir, exist_ok=True)
    paths = []
    for name, (w, h, text, icon, fill, line, color, right) in BUTTONS.items():
        p = os.path.join(out_dir, f"btn_{name}.png")
        render_button(p, icons_dir, w, h, text, icon, fill, line, color, right)
        paths.append(p)
    return paths


if __name__ == "__main__":
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    icons = os.path.join(root, "assets", "icons")
    for p in render_all(icons) + render_states(icons) + render_status_dots(icons):
        print(p)


# ----------------------------------------------------------------------------
# Значки общего состояния (щит), переключаются макросом UpdateStateIcon
# ----------------------------------------------------------------------------
STATES = {
    # состояние: (цвет, символ)
    0: ("#6B7F8A", "?"),
    1: (C["green"], "check"),
    2: (C["yellow"], "!"),
    3: (C["red"], "!"),
    4: (C["red"], "x"),
}


def render_state_icon(out_path: str, color: str, glyph: str, size: int = 64) -> None:
    ss = 8
    W = size * SCALE * ss
    img = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    rgb = _rgb(color)

    def P(x, y):  # координаты в квадрате 100x100
        return (x * W / 100, y * W / 100)

    # контур щита: прямые плечи, скругленный низ
    pts = [P(50, 5), P(87, 18)]
    for i in range(0, 21):  # правый бок -> низ
        t = i / 20
        x = 87 - (87 - 50) * t ** 1.6
        y = 18 + (95 - 18) * (t ** 0.75)
        pts.append(P(x, y))
    left = [(W - x, y) for (x, y) in reversed(pts[1:])]
    poly = pts + left[1:]
    d.polygon(poly, fill=rgb + (46,))
    lw = int(W * 0.055)
    d.line(poly + [poly[0]], fill=rgb + (255,), width=lw, joint="curve")
    gw = int(W * 0.08)
    if glyph == "check":
        d.line([P(33, 50), P(45, 62), P(68, 37)], fill=(255, 255, 255, 255), width=gw,
               joint="curve")
    elif glyph == "x":
        d.line([P(36, 36), P(64, 64)], fill=(255, 255, 255, 255), width=gw)
        d.line([P(64, 36), P(36, 64)], fill=(255, 255, 255, 255), width=gw)
    elif glyph == "!":
        d.line([P(50, 28), P(50, 56)], fill=(255, 255, 255, 255), width=gw)
        r = gw * 0.62
        cx, cy = P(50, 69)
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(255, 255, 255, 255))
    else:
        f = _font(int(W * 0.42))
        tb = d.textbbox((0, 0), glyph, font=f)
        d.text(((W - (tb[2] - tb[0])) / 2 - tb[0], W * 0.47 - (tb[3] - tb[1]) / 2 - tb[1]),
               glyph, font=f, fill=(255, 255, 255, 230))
    img = img.resize((size * SCALE, size * SCALE), Image.LANCZOS)
    img.save(out_path)


def render_states(out_dir: str) -> list[str]:
    paths = []
    for st, (color, glyph) in STATES.items():
        p = os.path.join(out_dir, f"state_{st}.png")
        render_state_icon(p, color, glyph)
        paths.append(p)
    return paths


def render_status_dot(out_path: str, color: str, glyph: str, size: int = 24) -> None:
    """Маленький значок состояния для подвала: кольцо + символ."""
    ss = 8
    W = size * SCALE * ss
    img = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    rgb = _rgb(color)
    lw = int(W * 0.09)
    m = lw // 2 + int(W * 0.04)
    d.ellipse((m, m, W - m, W - m), outline=rgb + (255,), width=lw)

    def P(x, y):
        return (x * W / 100, y * W / 100)

    gw = int(W * 0.1)
    if glyph == "check":
        d.line([P(30, 52), P(44, 65), P(70, 38)], fill=rgb + (255,), width=gw, joint="curve")
    elif glyph == "!":
        d.line([P(50, 26), P(50, 58)], fill=rgb + (255,), width=gw)
        r = gw * 0.6
        cx, cy = P(50, 73)
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=rgb + (255,))
    elif glyph == "x":
        d.line([P(35, 35), P(65, 65)], fill=rgb + (255,), width=gw)
        d.line([P(65, 35), P(35, 65)], fill=rgb + (255,), width=gw)
    else:
        f = _font(int(W * 0.55))
        tb = d.textbbox((0, 0), glyph, font=f)
        d.text(((W - (tb[2] - tb[0])) / 2 - tb[0], (W - (tb[3] - tb[1])) / 2 - tb[1]),
               glyph, font=f, fill=rgb + (255,))
    img = img.resize((size * SCALE, size * SCALE), Image.LANCZOS)
    img.save(out_path)


def render_status_dots(out_dir: str) -> list[str]:
    paths = []
    for st, (color, glyph) in STATES.items():
        p = os.path.join(out_dir, f"foot_{st}.png")
        render_status_dot(p, color, glyph)
        paths.append(p)
    return paths
