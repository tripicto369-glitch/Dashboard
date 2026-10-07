#!/usr/bin/env python3
"""Сборка дашборда ЛХОС.

    python3 build.py            # дашборд + мокап-данные в dist/
    python3 build.py --no-mock  # только дашборд

Результат:
    dist/Дашборд_ЛХОС.xlsm                — книга-дашборд с макросами импорта;
    dist/Мокап_ЛХОС_2026.xlsx             — пример исходных данных (выдуманные СИКН);
    dist/Мокап_ЛХОС_2026_дополнение.xlsx  — тот же файл + новые дни/объект (проверка
                                            режима «добавить только новые»).
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "src"))

from lhos.buttons import render_all, render_states, render_status_dots  # noqa: E402
from lhos.dashboard import NAV, build_workbook  # noqa: E402
from lhos.vbaproject import VbaModule, build_vba_project  # noqa: E402

VBA_DIR = os.path.join(ROOT, "src", "vba")
ICONS_DIR = os.path.join(ROOT, "assets", "icons")
DIST = os.path.join(ROOT, "dist")
DASHBOARD = os.path.join(DIST, "Дашборд_ЛХОС.xlsm")

# кодовые имена листов (совпадают с dashboard.build_workbook)
SHEET_CODES = [n[3] for n in NAV] + ["shDB", "shCalc"]


def _read(name: str) -> str:
    path = os.path.join(VBA_DIR, name)
    if not os.path.exists(path):
        return "Option Explicit\n"
    with open(path, encoding="utf-8") as f:
        return f.read()


def vba_modules() -> list[VbaModule]:
    mods = [VbaModule("ThisWorkbook", "workbook", _read("ThisWorkbook.cls"))]
    for code in SHEET_CODES:
        mods.append(VbaModule(code, "worksheet", _read(code + ".cls")))
    for fn in sorted(os.listdir(VBA_DIR)):
        if fn.endswith(".bas"):
            mods.append(VbaModule(fn[:-4], "standard", _read(fn)))
    return mods


def formula_results(path: str) -> dict:
    """Результаты формул пустой книги, вычисленные LibreOffice (если он установлен).

    xlsxwriter не вычисляет формулы и сохраняет у них результат 0, а Excel до
    пересчета (например, в режиме защищенного просмотра) показывает именно
    сохраненные результаты. Поэтому книга собирается дважды: второй раз — с
    результатами из первого прохода.
    """
    try:
        sys.path.insert(0, os.path.join(ROOT, "tests"))
        from lo_harness import LibreOffice  # noqa: E402
    except Exception as exc:  # нет LibreOffice/uno — собираем без кэша
        print("  (результаты формул не вычислены: нет LibreOffice/uno:", exc, ")")
        return {}
    cache: dict = {}
    with LibreOffice() as lo:
        lo.set_config("/org.openoffice.Office.Calc/Formula/Load", OOXMLRecalcMode=0)
        lo.set_config("/org.openoffice.Setup/L10N", ooSetupSystemLocale="ru-RU")
        doc = lo.open(path, macros=False)
        doc.recalc()
        sheets = doc.doc.Sheets
        for i in range(sheets.Count):
            sh = sheets.getByIndex(i)
            cells = sh.queryContentCells(16).getCells().createEnumeration()  # 16 = FORMULA
            while cells.hasMoreElements():
                cell = cells.nextElement()
                a = cell.getCellAddress()
                if cell.getError():
                    v = cell.getString() or "#N/A"
                elif cell.FormulaResultType2 == 2:  # STRING
                    v = cell.getString()
                else:
                    v = cell.getValue()
                cache[(sh.Name, a.Row, a.Column)] = v
        doc.close()
    return cache


def build_dashboard(path: str = DASHBOARD, cache_results: bool = True) -> str:
    from lhos import canvas

    os.makedirs(os.path.dirname(path), exist_ok=True)
    render_all(ICONS_DIR)
    render_states(ICONS_DIR)
    render_status_dots(ICONS_DIR)
    with tempfile.TemporaryDirectory() as tmp:
        bin_path = os.path.join(tmp, "vbaProject.bin")
        with open(bin_path, "wb") as f:
            f.write(build_vba_project(vba_modules()))
        canvas.FORMULA_CACHE.clear()
        build_workbook(path, ICONS_DIR, bin_path)
        if cache_results:
            canvas.FORMULA_CACHE.update(formula_results(path))
            if canvas.FORMULA_CACHE:
                build_workbook(path, ICONS_DIR, bin_path)
            canvas.FORMULA_CACHE.clear()
    return path


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--no-mock", action="store_true", help="не пересоздавать мокап-данные")
    ap.add_argument("-o", "--output", default=DASHBOARD, help="путь к .xlsm")
    ap.add_argument("--no-cache", action="store_true",
                    help="не вычислять результаты формул через LibreOffice")
    args = ap.parse_args()
    print("Дашборд:", build_dashboard(args.output, cache_results=not args.no_cache))
    if not args.no_mock:
        from lhos.mockdata import make_mock
        main_p = os.path.join(DIST, "Мокап_ЛХОС_2026.xlsx")
        upd_p = os.path.join(DIST, "Мокап_ЛХОС_2026_дополнение.xlsx")
        make_mock(main_p, upd_p)
        print("Мокап:", main_p)
        print("Мокап (дополнение):", upd_p)


if __name__ == "__main__":
    main()
