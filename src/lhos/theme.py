"""Палитра, шрифты и общие константы оформления дашборда.

Цвета подобраны по эталонному скриншоту (тёмная тема, зелёный/жёлтый/красный
акценты). Шрифт — Segoe UI (есть на любой Windows); при его отсутствии Excel
подставит ближайший.
"""

FONT = "Segoe UI"
FONT_SEMI = "Segoe UI Semibold"

C = {
    # фоны
    "page": "#04121A",
    "header": "#020B11",
    "panel": "#0A1D26",
    "panel_hdr": "#0E2530",
    "input": "#0F2733",
    "nav": "#071923",
    "nav_active": "#15323F",
    "footer": "#061822",
    # линии
    "border": "#1B3340",
    "line": "#122833",
    "divider": "#1E3A47",
    # текст
    "text": "#E4EBEF",
    "text2": "#B9C6CD",
    "muted": "#8C9DA7",
    "dim": "#5D707B",
    # статусы
    "green": "#4CC35A",
    "green_soft": "#4FBF7A",
    "yellow": "#F5B70F",
    "red": "#F04B46",
    "green_block": "#49B649",
    "yellow_block": "#F9AE0A",
    "red_block": "#DF1F1E",
    "nodata": "#2A3D47",
    # элементы управления
    "btn": "#17663A",
    "btn_line": "#2FA84F",
    "btn2": "#13303D",
    "btn2_line": "#2A4A58",
    "logo": "#2FA84F",
    "logo_red": "#E2231A",
}

# Коды зон (совпадают с формулами листа «Расчет»)
ZONE_EMPTY = -1   # пустая строка таблицы
ZONE_NODATA = 0   # объект есть, значения нет
ZONE_GREEN = 1
ZONE_YELLOW = 2
ZONE_RED = 3
ZONE_OVER = 4     # выше норматива

ZONE_COLOR = {
    ZONE_NODATA: C["dim"],
    ZONE_GREEN: C["green"],
    ZONE_YELLOW: C["yellow"],
    ZONE_RED: C["red"],
    ZONE_OVER: C["red"],
}

MONTHS_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]
