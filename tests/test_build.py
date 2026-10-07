"""Сборка книги и проверка, что VBA-проект компилируется.

Запуск: ``python3 -I tests/test_build.py`` (нужен LibreOffice с python3-uno).

LibreOffice компилирует все модули библиотеки вместе: одна синтаксическая
ошибка (например, переменная с именем зарезервированного оператора ``Imp``)
делает недоступными все макросы. Excel в такой ситуации тоже выдает ошибку
компиляции, поэтому проверка ловит реальные дефекты.
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import build  # noqa: E402

# Ключевые слова и операторы VBA, которые нельзя использовать как имена переменных
RESERVED = {
    "and", "as", "boolean", "byref", "byte", "byval", "call", "case", "const", "currency",
    "date", "decimal", "declare", "dim", "do", "double", "each", "else", "elseif", "empty",
    "end", "enum", "eqv", "erase", "error", "event", "exit", "false", "for", "friend",
    "function", "get", "global", "gosub", "goto", "if", "imp", "implements", "in", "integer",
    "is", "let", "like", "long", "loop", "lset", "me", "mod", "new", "next", "not",
    "nothing", "null", "object", "on", "optional", "or", "paramarray", "preserve", "private",
    "property", "public", "raiseevent", "redim", "rem", "resume", "return", "rset", "select",
    "set", "single", "static", "step", "stop", "string", "sub", "then", "to", "true", "type",
    "typeof", "until", "variant", "wend", "while", "with", "withevents", "xor",
}


def _declared_names(code: str) -> list[str]:
    names = []
    for m in re.finditer(r"(?im)^\s*(?:dim|private|public|static)\s+(?!sub\b|function\b|const\b)(.+)$",
                         code):
        for part in re.split(r",(?![^(]*\))", m.group(1)):
            tok = part.strip().split()
            if tok:
                names.append(tok[0].split("(")[0])
    for m in re.finditer(r"(?im)^\s*(?:public\s+|private\s+)?(?:sub|function)\s+\w+\s*\((.*?)\)",
                         code):
        for part in m.group(1).split(","):
            tok = [t for t in part.split() if t.lower() not in ("byval", "byref", "optional")]
            if tok:
                names.append(tok[0].split("(")[0])
    return names


def test_no_reserved_identifiers() -> None:
    bad = []
    for mod in build.vba_modules():
        for nm in _declared_names(mod.code):
            if nm.lower() in RESERVED:
                bad.append(f"{mod.name}: {nm}")
    assert not bad, "зарезервированные слова как имена: " + ", ".join(bad)


def test_vba_project_compiles_in_libreoffice() -> None:
    from lo_harness import LibreOffice

    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "dash.xlsm")
        build.build_dashboard(path, cache_results=False)
        with LibreOffice() as lo:
            doc = lo.open(path)
            lib = doc.doc.BasicLibraries.getByName("VBAProject")
            lib.insertByName("modProbe", 'Option VBASupport 1\nFunction P0() As String\n'
                                         ' P0 = "alive"\nEnd Function\n')
            assert doc.run_macro("modProbe", "P0", library="VBAProject") == "alive", \
                "библиотека VBAProject не компилируется"
            assert doc.run_macro("modImport", "MonthLabel", (46300.0,),
                                 library="VBAProject") == "Октябрь 2026"
            modules = set(doc.basic_modules()["VBAProject"])
            expected = {m.name for m in build.vba_modules()} | {"modProbe"}
            assert modules == expected, (modules, expected)
            doc.close()


def test_chart_values_are_numeric_refs() -> None:
    """Ряды значений диаграмм — числовые (numRef): строковый кэш (например, из
    сохраненного «#N/A») заставил бы рисовать точки нулями."""
    import zipfile

    path = ROOT / "dist" / "Дашборд_ЛХОС.xlsm"
    with zipfile.ZipFile(path) as z:
        charts = [n for n in z.namelist() if n.startswith("xl/charts/chart")]
        assert charts
        for n in charts:
            xml = z.read(n).decode("utf-8")
            for val in re.findall(r"<c:val>(.*?)</c:val>", xml, flags=re.S):
                assert "<c:numRef>" in val and "<c:strRef>" not in val, n


if __name__ == "__main__":
    failed = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print("PASS", name)
            except Exception as exc:  # noqa: BLE001
                failed += 1
                print("FAIL", name, "-", exc)
    sys.exit(1 if failed else 0)
