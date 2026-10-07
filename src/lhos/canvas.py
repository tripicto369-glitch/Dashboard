"""Пиксельная разметка листа Excel.

Макет описывается прямоугольниками в пикселях (x0, y0, x1, y1). По всем
границам прямоугольников строится сетка столбцов и строк точной ширины/высоты,
после чего каждый элемент превращается в ячейку или объединённый диапазон.
Так дизайн задаётся «как в макете», а не подгоняется под ширины столбцов.

Порядок отрисовки:
  1. fill()  — заливка областей (фон страницы, панели);
  2. frame()/hline()/vline() — линии по краям ячеек;
  3. put()   — содержимое (текст, число, формула) с объединением ячеек;
  4. image()/chart()/textbox() — плавающие объекты, привязанные к пикселям.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from .theme import C, FONT

Rect = tuple[int, int, int, int]

# Кэш результатов формул {(лист, строка, столбец): значение}. Excel показывает
# сохраненные результаты до пересчета (например, в режиме защищенного просмотра),
# поэтому сборка может заполнить их заранее (см. build.py).
FORMULA_CACHE: dict = {}


class Rich:
    """Статичный текст из нескольких фрагментов разного стиля: Rich(("TAT", {...}), ...)."""

    def __init__(self, *parts: tuple[str, dict]):
        self.parts = parts


def cached(ws, r: int, c: int):
    """Сохраненный результат формулы (0 — как пишет xlsxwriter по умолчанию)."""
    v = FORMULA_CACHE.get((ws.name, r, c))
    return 0 if v is None else v

_BORDER_SIDES = ("left", "right", "top", "bottom")


@dataclass
class _Content:
    rect: Rect
    value: Any
    style: dict
    kind: str  # "value" | "formula" | "array" | "url" | "blank"
    extra: dict = field(default_factory=dict)


class Canvas:
    """Рисует пиксельный макет на листе xlsxwriter."""

    def __init__(self, wb, ws, width: int, height: int, *, bg: str = C["page"],
                 base_font: str = FONT):
        self.wb = wb
        self.ws = ws
        self.width = width
        self.height = height
        self.bg = bg
        self.base_font = base_font
        self.xs: set[int] = {0, width}
        self.ys: set[int] = {0, height}
        self._fills: list[tuple[Rect, dict]] = []
        self._borders: list[tuple[Rect, str, str, str]] = []
        self._contents: list[_Content] = []
        self._objects: list[tuple[str, int, int, tuple, dict]] = []
        self._fmt_cache: dict[tuple, Any] = {}
        self.col_x: list[int] = []
        self.row_y: list[int] = []
        self._built = False

    # ------------------------------------------------------------------ API
    def fill(self, rect: Rect, color: str) -> None:
        self._add_rect(rect)
        self._fills.append((rect, {"bg": color}))

    def frame(self, rect: Rect, color: str = C["border"], style: str = "thin",
              sides: str = "ltrb") -> None:
        """Рамка по краям прямоугольника (стороны: l, t, r, b)."""
        self._add_rect(rect)
        names = {"l": "left", "t": "top", "r": "right", "b": "bottom"}
        for s in sides:
            self._borders.append((rect, names[s], color, style))

    def hline(self, x0: int, x1: int, y: int, color: str = C["line"]) -> None:
        """Горизонтальная линия на границе y (нижняя рамка ячеек над y)."""
        self.xs.update((x0, x1))
        self.ys.add(y)
        self._borders.append(((x0, y, x1, y), "hline", color, "thin"))

    def vline(self, x: int, y0: int, y1: int, color: str = C["line"]) -> None:
        """Вертикальная линия на границе x (левая рамка ячеек справа от x)."""
        self.xs.add(x)
        self.ys.update((y0, y1))
        self._borders.append(((x, y0, x, y1), "vline", color, "thin"))

    def put(self, rect: Rect, value: Any = None, **style) -> None:
        """Значение/формула в прямоугольник (объединяется при необходимости).

        Формула — строка, начинающаяся с «=»; формула массива — с «{=».
        Особые ключи style: url, url_tip, validation (dict для data_validation).
        """
        self._add_rect(rect)
        extra = {k: style.pop(k) for k in ("url", "url_tip", "validation", "name")
                 if k in style}
        if value is None:
            kind = "blank"
        elif isinstance(value, Rich):
            kind = "rich"
        elif isinstance(value, str) and value.startswith("{="):
            kind = "array"
        elif isinstance(value, str) and value.startswith("="):
            kind = "formula"
        else:
            kind = "value"
        if "url" in extra:
            kind = "url"
        self._contents.append(_Content(rect, value, style, kind, extra))

    def image(self, x: int, y: int, path: str, *, scale: float = 1.0, url: str | None = None,
              tip: str | None = None, descr: str | None = None) -> None:
        opts: dict[str, Any] = {"x_scale": scale, "y_scale": scale, "object_position": 3}
        if url:
            opts["url"] = url
            if tip:
                opts["tip"] = tip
        if descr:
            opts["description"] = descr
        self._objects.append(("image", x, y, (path,), opts))

    def chart(self, rect: Rect, chart) -> None:
        x0, y0, x1, y1 = rect
        chart.set_size({"width": x1 - x0, "height": y1 - y0})
        self._objects.append(("chart", x0, y0, (chart,), {"object_position": 3}))

    def textbox(self, rect: Rect, text: str, opts: dict) -> None:
        x0, y0, x1, y1 = rect
        o = dict(opts)
        o.update({"width": x1 - x0, "height": y1 - y0, "object_position": 3})
        self._objects.append(("textbox", x0, y0, (text,), o))

    def sparkline(self, rect: Rect, options: dict) -> None:
        self.put(rect, None)
        self._objects.append(("sparkline", rect[0], rect[1], (), options))

    def cell(self, x: int, y: int) -> tuple[int, int]:
        """(row, col) ячейки с левым верхним углом в точке (x, y)."""
        self._ensure_grid()
        return self.row_y.index(y), self.col_x.index(x)

    def cell_range(self, rect: Rect) -> tuple[int, int, int, int]:
        self._ensure_grid()
        x0, y0, x1, y1 = rect
        return (self.row_y.index(y0), self.col_x.index(x0),
                self.row_y.index(y1) - 1, self.col_x.index(x1) - 1)

    def ref(self, rect: Rect, absolute: bool = True) -> str:
        """A1-ссылка на левую верхнюю ячейку прямоугольника."""
        from xlsxwriter.utility import xl_rowcol_to_cell
        r, c, _, _ = self.cell_range(rect)
        return xl_rowcol_to_cell(r, c, row_abs=absolute, col_abs=absolute)

    def range_ref(self, rect: Rect) -> str:
        from xlsxwriter.utility import xl_range_abs
        r0, c0, r1, c1 = self.cell_range(rect)
        return xl_range_abs(r0, c0, r1, c1)

    @property
    def n_cols(self) -> int:
        self._ensure_grid()
        return len(self.col_x) - 1

    @property
    def n_rows(self) -> int:
        self._ensure_grid()
        return len(self.row_y) - 1

    # -------------------------------------------------------------- internals
    def _add_rect(self, rect: Rect) -> None:
        if self._built:
            raise RuntimeError("canvas already built")
        x0, y0, x1, y1 = rect
        if not (0 <= x0 < x1 <= self.width and 0 <= y0 < y1 <= self.height):
            raise ValueError(f"rect {rect} out of canvas {self.width}x{self.height}")
        self.xs.update((x0, x1))
        self.ys.update((y0, y1))

    def _ensure_grid(self) -> None:
        if not self.col_x:
            self.col_x = sorted(self.xs)
            self.row_y = sorted(self.ys)

    def fmt(self, style: dict):
        """Кэшированный формат xlsxwriter из словаря стиля."""
        key = tuple(sorted(style.items()))
        f = self._fmt_cache.get(key)
        if f is None:
            props: dict[str, Any] = {
                "font_name": style.get("font", self.base_font),
                "font_size": style.get("size", 10),
                "font_color": style.get("color", C["text"]),
                "valign": style.get("valign", "vcenter"),
            }
            if style.get("bg"):
                props["bg_color"] = style["bg"]
                props["pattern"] = 1
            for k_src, k_dst in (("bold", "bold"), ("italic", "italic"), ("align", "align"),
                                 ("num", "num_format"), ("indent", "indent"),
                                 ("wrap", "text_wrap"), ("shrink", "shrink"),
                                 ("underline", "underline"), ("locked", "locked")):
                if k_src in style:
                    props[k_dst] = style[k_src]
            for side in _BORDER_SIDES:
                b = style.get("b_" + side)
                if b:
                    color, bstyle = b
                    props[side] = {"thin": 1, "medium": 2, "dashed": 3, "dotted": 4,
                                   "thick": 5, "hair": 7}[bstyle]
                    props[side + "_color"] = color
            f = self.wb.add_format(props)
            self._fmt_cache[key] = f
        return f

    def build(self) -> None:
        """Записывает сетку, заливки, линии и содержимое на лист."""
        self._ensure_grid()
        self._built = True
        ws = self.ws
        ncols, nrows = len(self.col_x) - 1, len(self.row_y) - 1
        page_fmt = self.fmt({"bg": self.bg})

        for c in range(ncols):
            w = self.col_x[c + 1] - self.col_x[c]
            ws.set_column_pixels(c, c, w, page_fmt)
        for r in range(nrows):
            h = self.row_y[r + 1] - self.row_y[r]
            ws.set_row_pixels(r, h)

        # Базовый стиль каждой ячейки: заливки, затем рамки.
        base: dict[tuple[int, int], dict] = {}
        for rect, st in self._fills:
            r0, c0, r1, c1 = self.cell_range(rect)
            for r in range(r0, r1 + 1):
                for c in range(c0, c1 + 1):
                    base.setdefault((r, c), {}).update(st)
        for rect, side, color, bstyle in self._borders:
            if side == "hline":
                x0, y, x1, _ = rect
                r = self.row_y.index(y) - 1
                c0, c1 = self.col_x.index(x0), self.col_x.index(x1) - 1
                cells, side = [(r, c) for c in range(c0, c1 + 1)], "bottom"
            elif side == "vline":
                x, y0, _, y1 = rect
                c = self.col_x.index(x)
                r0, r1 = self.row_y.index(y0), self.row_y.index(y1) - 1
                cells, side = [(r, c) for r in range(r0, r1 + 1)], "left"
            else:
                r0, c0, r1, c1 = self.cell_range(rect)
                cells = {
                    "left": [(r, c0) for r in range(r0, r1 + 1)],
                    "right": [(r, c1) for r in range(r0, r1 + 1)],
                    "top": [(r0, c) for c in range(c0, c1 + 1)],
                    "bottom": [(r1, c) for c in range(c0, c1 + 1)],
                }[side]
            for rc in cells:
                base.setdefault(rc, {})["b_" + side] = (color, bstyle)

        covered: set[tuple[int, int]] = set()
        for ct in self._contents:
            r0, c0, r1, c1 = self.cell_range(ct.rect)
            # стиль = заливка левой верхней ячейки + рамки, лежащие на краях диапазона
            st: dict = {k: v for k, v in base.get((r0, c0), {}).items()
                        if not k.startswith("b_")}
            edges = {
                "b_left": [(r, c0) for r in range(r0, r1 + 1)],
                "b_right": [(r, c1) for r in range(r0, r1 + 1)],
                "b_top": [(r0, c) for c in range(c0, c1 + 1)],
                "b_bottom": [(r1, c) for c in range(c0, c1 + 1)],
            }
            for key, cells in edges.items():
                for rc in cells:
                    if key in base.get(rc, {}):
                        st[key] = base[rc][key]
                        break
            st.update(ct.style)
            if "bg" not in st:
                st["bg"] = self.bg
            f = self.fmt(st)
            for r in range(r0, r1 + 1):
                for c in range(c0, c1 + 1):
                    if (r, c) in covered:
                        raise ValueError(f"overlapping content at {(r, c)} for {ct.rect}")
                    covered.add((r, c))
            merged = (r0, c0) != (r1, c1)
            if merged:
                ws.merge_range(r0, c0, r1, c1, "", f)
            self._write(r0, c0, ct, f)
            ct.extra["_cell"] = (r0, c0, r1, c1)

        # Ячейки без содержимого, но с заливкой/рамкой.
        for (r, c), st in base.items():
            if (r, c) in covered:
                continue
            s = dict(st)
            s.setdefault("bg", self.bg)
            ws.write_blank(r, c, None, self.fmt(s))

        for kind, x, y, args, opts in self._objects:
            self._place(kind, x, y, args, opts)

    def _write(self, r: int, c: int, ct: _Content, f) -> None:
        ws = self.ws
        if ct.kind == "blank":
            ws.write_blank(r, c, None, f)
        elif ct.kind == "array":
            ws.write_array_formula(r, c, r, c, ct.value, f, cached(ws, r, c))
        elif ct.kind == "formula":
            ws.write_formula(r, c, ct.value, f, cached(ws, r, c))
        elif ct.kind == "rich":
            args = []
            for text, st in ct.value.parts:
                base = {k: v for k, v in ct.style.items() if not k.startswith("b_")}
                base.update(st)
                base["bg"] = None
                args += [self.fmt(base), text]
            ws.write_rich_string(r, c, *args, f)
        elif ct.kind == "url":
            ws.write_url(r, c, ct.extra["url"], f, string=ct.value,
                         tip=ct.extra.get("url_tip"))
        else:
            ws.write(r, c, ct.value, f)
        if "validation" in ct.extra:
            r0, c0, r1, c1 = self.cell_range(ct.rect)
            ws.data_validation(r0, c0, r1, c1, ct.extra["validation"])
        if "name" in ct.extra:
            from xlsxwriter.utility import xl_rowcol_to_cell
            ref = xl_rowcol_to_cell(r, c, row_abs=True, col_abs=True)
            self.wb.define_name(ct.extra["name"], f"='{self.ws.name}'!{ref}")

    def _locate(self, x: int, y: int) -> tuple[int, int, int, int]:
        """Ячейка, содержащая точку, и смещение внутри неё."""
        self._ensure_grid()
        c = max(i for i, cx in enumerate(self.col_x[:-1]) if cx <= x)
        r = max(i for i, ry in enumerate(self.row_y[:-1]) if ry <= y)
        return r, c, x - self.col_x[c], y - self.row_y[r]

    def _place(self, kind: str, x: int, y: int, args: tuple, opts: dict) -> None:
        ws = self.ws
        if kind == "sparkline":
            r, c = self.cell(x, y)
            ws.add_sparkline(r, c, opts)
            return
        r, c, dx, dy = self._locate(x, y)
        o = dict(opts)
        o["x_offset"] = dx
        o["y_offset"] = dy
        if kind == "image":
            ws.insert_image(r, c, args[0], o)
        elif kind == "chart":
            ws.insert_chart(r, c, args[0], o)
        elif kind == "textbox":
            ws.insert_textbox(r, c, args[0], o)
