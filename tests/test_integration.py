"""End-to-end test of the LHOS dashboard workbook in headless LibreOffice.

The expected results come from ``tests/oracle.py`` — an independent Python
model built from the user's requirements (it never reads the workbook's
formulas).  The scenario:

1. open a copy of ``dist/Дашборд_ЛХОС.xlsm`` -> empty state;
2. first import (mode 0) of the main mock file -> report, БД, sys_* cells, log,
   Расчет KPIs / sorted table / trends / max-per-day / month list and the
   visible texts of «Сводка» for TODAY (real date) and for several months
   chosen through ``sel_Month`` (incl. February with "н/д", March with the
   3.2 exceedance on 17.03 and the "0,45" text value, September with a blank
   last day); the rendered value column (PDF text) shows "*" for stale values;
   the «Обновление: …» label for several ages of the last import;
   other "today" dates (1st of a month, past month, exceedance day, "н/д" day)
   by replacing c_Today with a constant in the working copy;
3. threshold changes on «Настройки» re-zone everything;
4. append import (mode 2) of the update file -> counts, new object, the
   conflicting value is NOT overwritten; mode 0 on a non-empty base appends;
5. replace import (mode 1) of the main file -> base equals the main file again;
   error paths (missing file, file without values, the dashboard itself) leave
   the base untouched; synthetic sources (December 2025 appended before the
   base, reordered columns, other spelling of a name, duplicate date rows,
   unmerged "Дата" header, text values) and a small replace (no stray cells);
6. ClearDatabaseSilent -> empty state again;
7. the import log on «Настройки» has one row per import, newest first.

Cell addresses are discovered from the workbook (openpyxl), so the test does
not depend on the generated grid.  All mismatches are collected and printed;
the run fails if there is any.

Run:  ``python3 -I tests/test_integration.py``  (or via pytest).  Set
``LHOS_E2E_OUT=<dir>`` to keep the working copy and a JSON with mismatches,
``LHOS_XLSM=<path>`` to test another build of the workbook.
"""

from __future__ import annotations

import datetime as dt
import importlib.util
import json
import os
import re
import shutil
import sys
import tempfile
import time
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))

HAVE_LO = importlib.util.find_spec("uno") is not None and shutil.which("soffice") is not None
if not HAVE_LO and "pytest" in sys.modules:  # pragma: no cover
    import pytest

    pytest.skip("LibreOffice/uno not available", allow_module_level=True)

import openpyxl
import oracle as O

XLSM = Path(os.environ.get("LHOS_XLSM") or ROOT / "dist" / "Дашборд_ЛХОС.xlsm")
MAIN = ROOT / "dist" / "Мокап_ЛХОС_2026.xlsx"
UPD = ROOT / "dist" / "Мокап_ЛХОС_2026_дополнение.xlsx"

EPOCH = dt.date(1899, 12, 30)
EPOCH_DT = dt.datetime(1899, 12, 30)  # noqa: DTZ001 - Excel serials are local, naive


def now_local() -> dt.datetime:
    """Local wall-clock time, as Excel's NOW() sees it (naive on purpose)."""
    return dt.datetime.now()  # noqa: DTZ005
TOL = 1e-9


def serial(d: dt.date) -> float:
    return float((d - EPOCH).days)


def from_serial(x) -> dt.date | None:
    if isinstance(x, (int, float)) and x > 0:
        return EPOCH + dt.timedelta(days=int(x))
    return None


# ---------------------------------------------------------------------------
# Layout discovery (addresses only; formulas are not evaluated or copied)
# ---------------------------------------------------------------------------

def discover_layout(path: Path) -> dict:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, keep_vba=True)
    names = {k: v.attr_text for k, v in wb.defined_names.items()}

    def name_addr(n):
        sheet, addr = names[n].rsplit("!", 1)
        return sheet.strip("'"), addr.replace("$", "")

    ws = wb["Сводка"]
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                cells.setdefault(c.value, []).append(c.coordinate)

    def find(pred, what):
        hits = [a for f, addrs in cells.items() if pred(f) for a in addrs]
        if not hits:
            raise AssertionError(f"layout: no cell for {what}")
        return hits

    dash = {
        "today": find(lambda f: f == "=TODAY()", "today")[0],
        "updated": name_addr("ui_Updated")[1],
        "data_date": find(lambda f: "c_DataDate" in f, "data date")[0],
        "loaded": find(lambda f: f.startswith("=IF(N(sys_LastImport)=0,\"—\""), "loaded")[0],
        "file": find(lambda f: f.startswith("=IF(sys_LastFile=\"\",\"—\""), "file")[0],
        "state_text": find(lambda f: f == "=c_StateText", "state text")[0],
        "state_sub": find(lambda f: f == "=c_StateSub", "state sub")[0],
        "max": find(lambda f: "c_MaxVal" in f, "max")[0],
        "max_name": find(lambda f: f == "=c_MaxName", "max name")[0],
        "avg": find(lambda f: "c_AvgVal" in f, "avg")[0],
        "avg_sub": find(lambda f: "c_CntValued" in f and "по " in f, "avg sub")[0],
        "green": find(lambda f: f == "=c_CntGreen", "green")[0],
        "yellow": find(lambda f: f == "=c_CntYellow", "yellow")[0],
        "red": find(lambda f: f == "=c_CntRed", "red")[0],
        "total": min(find(lambda f: f == "=c_ObjCount", "total"),
                     key=lambda a: int(re.sub(r"\D", "", a))),
        "donut_total": max(find(lambda f: f == "=c_ObjCount", "total"),
                           key=lambda a: int(re.sub(r"\D", "", a))),
        "total_sub": find(lambda f: "c_CntNoData" in f, "no data")[0],
        "title": find(lambda f: "c_MonthLabel" in f, "title")[0],
        "note": find(lambda f: "c_StaleCount" in f, "note")[0],
        "footer": find(lambda f: "c_Footer" in f, "footer")[0],
        "source": find(lambda f: "Источник данных" in f, "source")[0],
        "leg_g": find(lambda f: f == "=c_LegG", "legend g")[0],
        "leg_y": find(lambda f: f == "=c_LegY", "legend y")[0],
        "leg_r": find(lambda f: f == "=c_LegR", "legend r")[0],
        "zone_r": find(lambda f: f == "=c_ZoneR", "zone r")[0],
        "zone_y": find(lambda f: f == "=c_ZoneY", "zone y")[0],
        "zone_g": find(lambda f: f == "=c_ZoneG", "zone g")[0],
        "sel": name_addr("sel_Month")[1],
    }

    def col_of(addr):
        return re.match(r"[A-Z]+", addr).group(0)

    def row_of(addr):
        return int(re.search(r"\d+", addr).group(0))

    r0 = 92  # first sorted row on «Расчет» (columns used by the dashboard rows)
    first = {
        "num": find(lambda f: f == f"=IF('Расчет'!$C${r0}=\"\",\"\",1)", "row num")[0],
        "name": find(lambda f: f == f"='Расчет'!$C${r0}", "row name")[0],
        "value": find(lambda f: f.startswith(f"=IF('Расчет'!$C${r0}=\"\",\"\",IF('Расчет'!$D${r0}"), "row value")[0],
        "status": find(lambda f: f"$E${r0}" in f and "НОРМА" in f, "row status")[0],
        "zone": find(lambda f: f"$E${r0}" in f and "c_ZoneG" in f, "row zone")[0],
        "change": find(lambda f: f"$F${r0}" in f and "↑" not in f, "row change")[0],
        "arrow": find(lambda f: f"$F${r0}" in f and "↑" in f, "row arrow")[0],
        "zone_h": find(lambda f: f == f"='Расчет'!$E${r0}", "zone helper")[0],
        "stale_h": find(lambda f: f == f"='Расчет'!$G${r0}", "stale helper")[0],
    }
    cols = {k: col_of(a) for k, a in first.items()}
    rows = []
    for k in range(O.TABLE_ROWS):
        a = find(lambda f, r=r0 + k: f == f"='Расчет'!$C${r}", f"table row {k}")[0]
        rows.append(row_of(a))

    # conditional formats of the value column (stale "*" number format)
    cf = []
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        for rule in rules:
            num = rule.dxf.numFmt.formatCode if rule.dxf is not None and rule.dxf.numFmt is not None else None
            cf.append((str(rng.sqref), rule.formula, num))

    st = wb["Настройки"]
    log_cols = [name_addr(f"log_C{j}")[1] for j in range(1, 10)]
    jcol = openpyxl.utils.column_index_from_string(col_of(log_cols[0]))
    hdr_row = row_of(log_cols[0])
    hdr_last = max((m.max_row for m in st.merged_cells.ranges
                    if m.min_row == hdr_row and m.min_col == jcol), default=hdr_row)
    log_rows = sorted(m.min_row for m in st.merged_cells.ranges
                      if m.min_col == jcol and m.min_row > hdr_last)
    info = {}
    for row in st.iter_rows():
        for c in row:
            if c.value == "=COUNTA(db_Names)":
                info["objects"] = c.coordinate
            elif c.value == "=COUNT(db_Vals)":
                info["values"] = c.coordinate
            elif isinstance(c.value, str) and c.value.startswith("=IF(c_FirstDate=\"\""):
                info["period"] = c.coordinate
    return {
        "names": names, "dash": dash, "cols": cols, "rows": rows, "cf": cf,
        "log_cols": [col_of(a) for a in log_cols], "log_rows": log_rows, "info": info,
        "cfg": {n: name_addr(n)[1] for n in ("cfg_Green", "cfg_Yellow", "cfg_Limit")},
        "sys": {n: name_addr(n)[1] for n in ("sys_LastImport", "sys_LastFile", "sys_LastMode")},
        "calc": {n[2:]: name_addr(n)[1] for n in names if n.startswith("c_")
                 and name_addr(n)[0] == "Расчет"},
    }


# ---------------------------------------------------------------------------
# Mismatch collector
# ---------------------------------------------------------------------------

class Checker:
    def __init__(self):
        self.fails: list[dict] = []
        self.count = 0
        self.ctx = ""

    def eq(self, label, actual, expected, tol=TOL):
        self.count += 1
        ok = actual == expected
        if not ok and isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
            ok = abs(actual - expected) <= tol
        if not ok:
            self.fails.append({"ctx": self.ctx, "what": label, "actual": actual, "expected": expected})
        return ok

    def true(self, label, cond, detail=""):
        self.count += 1
        if not cond:
            self.fails.append({"ctx": self.ctx, "what": label, "actual": detail, "expected": True})
        return cond


# ---------------------------------------------------------------------------
# Reading the workbook through UNO
# ---------------------------------------------------------------------------

def cell_value(doc, sheet, addr):
    """None (empty / ""), float, str or ("ERR", code)."""
    c = doc.cell(sheet, addr)
    if c.getError():
        return ("ERR", c.getError())
    v = doc.get(sheet, addr)
    if v == "":
        return None
    return v


def read_block(doc, sheet, rng):
    """2-D list of cell_value for a range (errors as ("ERR", code))."""
    s = doc.sheet(sheet)
    r = s.getCellRangeByName(rng)
    data = r.getDataArray()
    out = []
    for i, row in enumerate(data):
        orow = []
        for j, v in enumerate(row):
            if v is None or v == "":  # getDataArray gives None for error results
                cell = r.getCellByPosition(j, i)
                err = cell.getError()
                orow.append(("ERR", err) if err else None)
            else:
                orow.append(v)
        out.append(orow)
    return out


def calc_param(doc, L, name):
    return cell_value(doc, "Расчет", L["calc"][name])


# ---------------------------------------------------------------------------
# Comparisons
# ---------------------------------------------------------------------------

def check_view(doc, L, v: O.View, ck: Checker, real_today: bool = True):
    """Compare «Расчет» and the visible texts of «Сводка» with the oracle view."""
    P = lambda n: calc_param(doc, L, n)
    ck.eq("c_MonthStart", from_serial(P("MonthStart")), v.month_start)
    ck.eq("c_MonthEnd", from_serial(P("MonthEnd")), v.month_end)
    ck.eq("c_RefDate", from_serial(P("RefDate")), v.ref_date)
    ck.eq("c_IsCurrent", bool(P("IsCurrent")), v.is_current)
    ck.eq("c_MonthLabel", P("MonthLabel"), v.label)
    ck.eq("c_ObjCount", P("ObjCount"), float(v.obj_count))
    ck.eq("c_DataDate", from_serial(P("DataDate")), v.data_date)
    ck.eq("c_CntValued", P("CntValued"), float(v.cnt_valued))
    ck.eq("c_CntGreen", P("CntGreen"), float(v.cnt_green))
    ck.eq("c_CntYellow", P("CntYellow"), float(v.cnt_yellow))
    ck.eq("c_CntRed", P("CntRed"), float(v.cnt_red))
    ck.eq("c_CntOver", P("CntOver"), float(v.cnt_over))
    ck.eq("c_CntNoData", P("CntNoData"), float(v.cnt_nodata))
    ck.eq("c_MaxVal", P("MaxVal"), v.max_val)
    ck.eq("c_MaxName", P("MaxName"), v.max_name or None)
    ck.eq("c_AvgVal", P("AvgVal"), v.avg_val, tol=1e-9)
    ck.eq("c_State", P("State"), float(v.state))
    ck.eq("c_StaleCount", P("StaleCount"), float(v.stale_count))
    ck.eq("c_LastMaxDate", from_serial(P("LastMaxDate")), v.last_max_date)

    # objects in base order (rows 46..85)
    blk = read_block(doc, "Расчет", "A46:K85")
    for i in range(O.MAX_OBJ):
        row = blk[i]
        if i < len(v.objects):
            o = v.objects[i]
            tag = f"obj[{i + 1}] {o.name}"
            ck.eq(tag + " name", row[1], o.name)
            ck.eq(tag + " group", row[2], o.group or None)
            ck.eq(tag + " value date", from_serial(row[3]), o.value_date)
            ck.eq(tag + " value", row[4], o.value)
            ck.eq(tag + " prev date", from_serial(row[5]), o.prev_date)
            ck.eq(tag + " prev value", row[6], o.prev_value)
            ck.eq(tag + " change", row[7], o.change, tol=1e-6)
            ck.eq(tag + " zone", row[8], float(o.zone))
            ck.eq(tag + " stale", row[10], 1.0 if o.stale else 0.0)
        else:
            ck.eq(f"obj[{i + 1}] empty name", row[1], None)

    # sorted table + trends (rows 92..131, trend L..AP)
    blk = read_block(doc, "Расчет", "A92:AP131")
    for k in range(O.MAX_OBJ):
        row = blk[k]
        if k < len(v.sorted):
            o = v.sorted[k]
            tag = f"sorted[{k + 1}]"
            ck.eq(tag + " name", row[2], o.name)
            ck.eq(tag + " value", row[3], o.value)
            ck.eq(tag + " zone", row[4], float(o.zone))
            ck.eq(tag + " change", row[5], o.change, tol=1e-6)
            ck.eq(tag + " stale", row[6], 1.0 if o.stale else 0.0)
            trend = [None if isinstance(x, tuple) else x for x in row[11:42]]
            errs_ok = all(isinstance(x, tuple) for x, e in zip(row[11:42], v.trend[k]) if e is None)
            ck.eq(tag + f" trend {o.name}", trend, v.trend[k])
            ck.true(tag + " trend gaps are #N/A", errs_ok, str(row[11:42]))
        else:
            ck.eq(f"sorted[{k + 1}] empty name", row[2], None)
            ck.true(f"sorted[{k + 1}] trend all #N/A",
                    all(isinstance(x, tuple) for x in row[11:42]), str(row[11:42])[:120])

    # day dates row 91 and max-per-day row 134, last point row 138
    days = read_block(doc, "Расчет", "L91:AP91")[0]
    exp_days = [v.days[k] if k < len(v.days) else None for k in range(31)]
    ck.eq("day dates", [from_serial(x) for x in days], exp_days)
    mx = read_block(doc, "Расчет", "L134:AP134")[0]
    ck.eq("max per day", [None if isinstance(x, tuple) else x for x in mx], v.maxday)
    last = read_block(doc, "Расчет", "L138:AP138")[0]
    exp_last = [v.maxday[k] if (v.last_max_date and k < len(v.days) and v.days[k] == v.last_max_date)
                else None for k in range(31)]
    ck.eq("last point", [None if isinstance(x, tuple) else x for x in last], exp_last)
    lim = read_block(doc, "Расчет", "L135:AP137")
    for r, val, nm in ((0, v.cfg.limit, "limit"), (1, v.cfg.yellow, "yellow"), (2, v.cfg.green, "green")):
        ck.eq(f"threshold line {nm}", [None if isinstance(x, tuple) else x for x in lim[r]],
              [val if k < len(v.days) else None for k in range(31)])

    # month list
    ml = read_block(doc, "Расчет", "B142:C201")
    got = [(from_serial(a), b) for a, b in ml if b]
    ck.eq("month list", got, v.months)

    # visible texts on «Сводка»
    exp = O.dash_texts(v)
    for key, want in exp.items():
        addr = L["dash"][key]
        ck.eq(f"Сводка {key} ({addr})", doc.get_string("Сводка", addr), want)
    if real_today:  # the header always shows the real date (TODAY())
        ck.eq("Сводка today", doc.get_string("Сводка", L["dash"]["today"]), v.today.strftime("%d.%m.%Y"))
    texts = O.table_texts(v)
    c = L["cols"]
    for k, r in enumerate(L["rows"]):
        got = tuple(doc.get_string("Сводка", f"{c[n]}{r}")
                    for n in ("num", "name", "value", "status", "zone", "change", "arrow"))
        ck.eq(f"Сводка table row {k + 1} (row {r})", got, texts[k])
        o = v.sorted[k] if k < len(v.sorted) else None
        zh = cell_value(doc, "Сводка", f"{c['zone_h']}{r}")
        sh = cell_value(doc, "Сводка", f"{c['stale_h']}{r}")
        if o is not None:
            ck.eq(f"Сводка row {k + 1} zone helper", zh, float(o.zone))
            ck.eq(f"Сводка row {k + 1} stale helper", sh, 1.0 if o.stale else 0.0)


def read_db(doc):
    """(objects, groups, {(date, name): value}, dates list) from «БД»."""
    s = doc.sheet("БД")
    hdr = s.getCellRangeByName("B1:CV2").getDataArray()
    names = []
    for x in hdr[0]:
        if x == "":
            break
        names.append(x)
    groups = {n: hdr[1][i] for i, n in enumerate(names)}
    col = s.getCellRangeByName("A3:A3702").getDataArray()
    dates = []
    for (x,) in col:
        if x == "":
            break
        dates.append(from_serial(x))
    vals = {}
    if names and dates:
        from openpyxl.utils import get_column_letter
        last = get_column_letter(1 + len(names))
        data = s.getCellRangeByName(f"B3:{last}{2 + len(dates)}").getDataArray()
        for i, row in enumerate(data):
            for j, x in enumerate(row):
                if x != "":
                    vals[(dates[i], names[j])] = x
    # anything outside the block (rows below, header/values right of the objects)?
    extra = s.getCellRangeByName(f"A{3 + len(dates)}:CV{3 + len(dates) + 5}").getDataArray()
    stray = any(x != "" for row in extra for x in row)
    right = s.getCellRangeByName(f"{_col(2 + len(names))}1:CV{max(3, 2 + len(dates))}").getDataArray()
    stray = stray or any(x != "" for row in right for x in row)
    return names, groups, vals, dates, stray


def _col(n: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(n)


def check_db(doc, db: O.Db, ck: Checker):
    names, groups, vals, dates, stray = read_db(doc)
    ck.eq("БД objects", names, db.objects)
    ck.eq("БД groups", groups, {n: db.groups.get(n, "") for n in db.objects})
    ck.eq("БД dates contiguous first..last", dates, db.dates())
    ck.eq("БД value count", len(vals), len(db.values))
    diff = {k: (vals.get(k), db.values.get(k)) for k in set(vals) | set(db.values)
            if vals.get(k) is None or db.values.get(k) is None or abs(vals[k] - db.values[k]) > 1e-12}
    ck.true("БД values equal oracle", not diff,
            str(sorted(((d.isoformat(), n), a, e) for (d, n), (a, e) in diff.items())[:10]))
    ck.true("БД no stray cells below / right of the data", not stray)


def parse_report(text: str) -> dict:
    out = {}
    for line in text.replace("\r", "").split("\n"):
        if ":" in line:
            k, _, val = line.partition(":")
            out[k.strip(" •")] = val.strip()
    return out


def num(s: str) -> int:
    return int(re.sub(r"[^\d]", "", s.split("(")[0]) or -1)


def check_report(res: str, rep: O.ImportReport, src: O.Source, ck: Checker):
    ck.true("import returns OK", res.startswith("OK\n"), res[:200])
    r = parse_report(res)
    ck.eq("report Файл", r.get("Файл"), src.file_name)
    ck.eq("report Значений в файле", num(r.get("Значений в файле", "")), rep.total)
    ck.eq("report Объектов в базе", num(r.get("Объектов в базе", "")), rep.obj_count)
    ck.eq("report Период", r.get("Период данных в базе"),
          f"{O.fmt_date(rep.first)} – {O.fmt_date(rep.last)}")
    ck.eq("report Листов с данными", r.get("Листов с данными"),
          f"{len(src.sheets_with_values)} ({', '.join(src.sheets_with_values)})")
    ck.eq("report Листов без значений", num(r.get("Листов без значений", "0")), src.empty_sheets)
    ck.eq("report Пропущено нечисловых", num(r.get("Пропущено нечисловых значений", "0")),
          src.text_skipped)
    if rep.mode == O.MODE_REPLACE:
        ck.true("report mode replace", "полная загрузка" in r.get("Режим", ""), r.get("Режим"))
        ck.eq("report Загружено", num(r.get("Загружено значений", "")), rep.added)
    else:
        ck.true("report mode append", "добавление" in r.get("Режим", ""), r.get("Режим"))
        ck.eq("report Добавлено", num(r.get("Добавлено новых значений", "")), rep.added)
        ck.eq("report Уже были", num(r.get("Уже были в базе", "")), rep.same)
        ck.eq("report conflicts", num(r.get("Отличаются от базы (оставлены как в базе)", "0")),
              rep.conflicts)
        for nm, d, old, new in rep.conflict_list[:5]:
            key = f"{nm}, {d.strftime('%d.%m.%Y')}"
            line = r.get(key, "")
            ck.true(f"report conflict example {key}",
                    O.fmt_num(old, 1, 4, ".") in line.replace(",", ".") and
                    O.fmt_num(new, 1, 4, ".") in line.replace(",", "."), line)
        if rep.new_objects:
            ck.eq("report Новых объектов", r.get("Новых объектов"),
                  f"{len(rep.new_objects)} ({', '.join(rep.new_objects)})")
        else:
            ck.eq("report no new objects line", r.get("Новых объектов"), None)


def check_sys(doc, L, ck: Checker, *, file_name, mode_text, since: dt.datetime | None):
    s = L["sys"]
    last = cell_value(doc, "Настройки", s["sys_LastImport"])
    if since is None:
        ck.eq("sys_LastImport empty", last, None)
    else:
        ok = isinstance(last, float)
        if ok:
            ts = EPOCH_DT + dt.timedelta(days=last)
            ok = since - dt.timedelta(minutes=2) <= ts <= now_local() + dt.timedelta(minutes=2)
        ck.true("sys_LastImport ~ now", ok, str(last))
    ck.eq("sys_LastFile", cell_value(doc, "Настройки", s["sys_LastFile"]), file_name)
    ck.eq("sys_LastMode", cell_value(doc, "Настройки", s["sys_LastMode"]), mode_text)
    d = L["dash"]
    if since is None:
        ck.eq("Сводка updated text", doc.get_string("Сводка", d["updated"]), "Данные еще не загружались")
        ck.eq("Сводка loaded", doc.get_string("Сводка", d["loaded"]), "—")
        ck.eq("Сводка file", doc.get_string("Сводка", d["file"]), "—")
        ck.eq("Сводка source", doc.get_string("Сводка", d["source"]),
              "Данные не загружены — нажмите «Импорт данных»")
    else:
        ck.true("Сводка updated text", doc.get_string("Сводка", d["updated"]).startswith("Обновление: "),
                doc.get_string("Сводка", d["updated"]))
        ck.true("Сводка updated just now",
                doc.get_string("Сводка", d["updated"]) in ("Обновление: только что",
                                                           "Обновление: 1 мин назад"),
                doc.get_string("Сводка", d["updated"]))
        ck.true("Сводка loaded dd.mm.yyyy hh:mm",
                re.fullmatch(r"\d\d\.\d\d\.\d{4} \d\d:\d\d", doc.get_string("Сводка", d["loaded"])) is not None,
                doc.get_string("Сводка", d["loaded"]))
        ck.eq("Сводка file", doc.get_string("Сводка", d["file"]), file_name)
        ck.eq("Сводка source", doc.get_string("Сводка", d["source"]), "Источник данных: " + file_name)


def check_info(doc, L, db: O.Db, ck: Checker):
    i = L["info"]
    ck.eq("Настройки objects in base", cell_value(doc, "Настройки", i["objects"]), float(len(db.objects)))
    ck.eq("Настройки values in base", cell_value(doc, "Настройки", i["values"]), float(len(db.values)))
    want = "—" if db.empty else f"{O.fmt_date(db.first)} – {O.fmt_date(db.last)}"
    ck.eq("Настройки period", doc.get_string("Настройки", i["period"]), want)


def read_log(doc, L):
    rows = []
    for r in L["log_rows"]:
        vals = [cell_value(doc, "Настройки", f"{c}{r}") for c in L["log_cols"]]
        if all(v is None for v in vals):
            continue
        rows.append(vals)
    return rows


def set_month(doc, L, label):
    doc.cell("Сводка", L["dash"]["sel"]).setString(label)
    doc.recalc()


def set_cfg(doc, L, cfg: O.Cfg):
    for n, v in (("cfg_Green", cfg.green), ("cfg_Yellow", cfg.yellow), ("cfg_Limit", cfg.limit)):
        doc.cell("Настройки", L["cfg"][n]).setValue(v)
    doc.recalc()


def elapsed_text(delta: dt.timedelta) -> str:
    """Expected «Обновление: …» text for the time since the last import."""
    mins = int(delta.total_seconds() // 60)
    if delta < dt.timedelta(minutes=1):
        return "Обновление: только что"
    if delta < dt.timedelta(hours=1):
        return f"Обновление: {mins} мин назад"
    if delta < dt.timedelta(days=1):
        return f"Обновление: {mins // 60} ч {mins % 60} мин назад"
    return f"Обновление: {delta.days} дн {int(delta.total_seconds() // 3600) % 24} ч назад"


def check_elapsed(doc, L, ck: Checker):
    addr = L["sys"]["sys_LastImport"]
    keep = doc.cell("Настройки", addr).getValue()
    for delta in (dt.timedelta(seconds=20), dt.timedelta(minutes=5, seconds=20),
                  dt.timedelta(hours=2, minutes=7, seconds=20),
                  dt.timedelta(days=3, hours=4, minutes=10, seconds=20)):
        stamp = now_local() - delta
        doc.cell("Настройки", addr).setValue(
            (stamp - EPOCH_DT).total_seconds() / 86400)
        doc.recalc()
        ck.eq(f"elapsed label for {delta}", doc.get_string("Сводка", L["dash"]["updated"]),
              elapsed_text(delta))
        el = cell_value(doc, "Сводка", L["names"]["ui_Elapsed"].split("!")[1].replace("$", ""))
        ck.eq(f"ui_Elapsed for {delta}", el, delta.total_seconds() / 86400, tol=5 / 86400)
    doc.cell("Настройки", addr).setValue(keep)
    doc.recalc()


def check_pdf_values(doc, v: O.View, ck: Checker, work: Path, tag: str):
    """Value column as rendered (conditional number format adds "*" when stale)."""
    import subprocess
    if not shutil.which("pdftotext"):
        return
    pdf = doc.export_pdf(work / f"{tag}.pdf")
    txt = subprocess.run(["pdftotext", "-layout", "-f", "1", "-l", "1", str(pdf), "-"],
                         capture_output=True, text=True, check=True).stdout
    for k, o in enumerate(v.sorted[:O.TABLE_ROWS]):
        want = "—" if o.value is None else O.fmt_num(o.value) + ("*" if o.stale else "")
        m = re.search(rf"(?m)^\s*(?:\S+\s+)?{k + 1}\s+{re.escape(o.name)}\s+(\S+)", txt)
        ck.eq(f"rendered value row {k + 1} {o.name}", m.group(1) if m else None, want)


def simulate_today(doc, L, day: dt.date | None):
    """Replace c_Today (=TODAY()) by a constant date in the working copy."""
    c = doc.cell("Расчет", L["calc"]["Today"])
    if day is None:
        c.setFormula("=TODAY()")
    else:
        c.setValue(serial(day))
    doc.recalc()


def write_source(path: Path, sheets: list):
    """sheets: [(title, header_kind, groups, names, rows)] where header_kind is
    "merged" (Дата in B4:B5, groups row 4, names row 5) or "single" (Дата and
    names in row 3); rows = [(date, [values...])]."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, kind, groups, names, rows in sheets:
        ws = wb.create_sheet(title)
        if kind is None:
            ws["A2"] = "Показатель"
            ws["B3"] = 1.5
            continue
        if kind == "merged":
            ws.merge_cells("B4:B5")
            ws["B4"] = "Дата"
            hdr, grow = 5, 4
        else:
            ws["B3"] = "Дата"
            hdr, grow = 3, None
        for j, nm in enumerate(names):
            ws.cell(hdr, 3 + j, nm)
            if grow and groups:
                ws.cell(grow, 3 + j, groups[j])
        for i, (d, vals) in enumerate(rows):
            ws.cell(hdr + 1 + i, 2, EPOCH_DT + (d - EPOCH)).number_format = "dd.mm.yyyy"
            for j, x in enumerate(vals):
                if x is not None:
                    ws.cell(hdr + 1 + i, 3 + j, x)
    wb.save(path)
    return path


def edge_sources(work: Path) -> dict:
    D = dt.date
    a = write_source(work / "edge_append.xlsx", [
        ("Декабрь 2025", "merged", ["Нефть тип А", "Новые"], ["СИКН №301", "Новый-1"], [
            (D(2025, 12, 29), [0.4, 1.9]),
            (D(2025, 12, 30), [" 1,25 ", "н/д"]),
            (D(2025, 12, 31), [0.9, "-"]),
            (D(2025, 12, 31), [0.45, 2.6]),          # duplicate date row: last wins
        ]),
        ("Свод", None, None, None, None),
        # other column order, other spelling of an existing object, a value that
        # equals the base and one that differs from it
        ("Январь", "single", None, ["Новый-1", "сикн  №304", "СИКН №301"], [
            (D(2026, 1, 1), [0.3, "__SAME304__", "__DIFF301__"]),
            (D(2026, 1, 2), [0.35, None, None]),
        ]),
    ])
    b = write_source(work / "edge_small.xlsx", [
        ("Август", "merged", ["Г1", "Г1", "Г2"], ["Узел-А", "Узел-Б", "Узел-В"], [
            (D(2026, 8, 10), [0.5, 1.6, 2.4]),
            (D(2026, 8, 11), [0.6, 1.7, 3.1]),
            (D(2026, 8, 12), [0.7, None, 0.2]),
        ]),
    ])
    c = write_source(work / "edge_novalues.xlsx", [
        ("Ноябрь", "merged", ["Г"], ["Объект"], [(D(2026, 11, 1), [None]), (D(2026, 11, 2), ["н/д"])]),
    ])
    return {"append": a, "small": b, "novalues": c}


def patch_edge_append(path: Path, base: O.Db):
    """Fill the placeholders with the base value (same) and base + 0.5 (conflict)."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Январь"]
    for row in ws.iter_rows():
        for c in row:
            if c.value == "__SAME304__":
                c.value = base.values[(dt.date(2026, 1, 1), "СИКН №304")]
            elif c.value == "__DIFF301__":
                c.value = round(base.values[(dt.date(2026, 1, 1), "СИКН №301")] + 0.5, 2)
    wb.save(path)


# ---------------------------------------------------------------------------
# Scenario
# ---------------------------------------------------------------------------

def run(out_dir: Path | None = None) -> Checker:
    from lo_harness import LibreOffice

    work = Path(out_dir) if out_dir else Path(tempfile.mkdtemp(prefix="lhos_e2e_"))
    work.mkdir(parents=True, exist_ok=True)
    book = work / "e2e_dashboard.xlsm"
    shutil.copy(XLSM, book)
    L = discover_layout(book)
    ck = Checker()
    today = dt.date.today()  # noqa: DTZ011 - must match TODAY() in the workbook
    src_main = O.read_source(MAIN)
    src_upd = O.read_source(UPD)
    default = O.Cfg()
    log_expect = []   # newest first: (mode text, file, total, added, same, conflicts, objects)

    with LibreOffice(base_dir=work) as lo:
        lo.set_config("/org.openoffice.Office.Calc/Formula/Load", OOXMLRecalcMode=0)
        lo.set_config("/org.openoffice.Setup/L10N", ooSetupSystemLocale="ru-RU")
        doc = lo.open(book)
        doc.recalc()

        # 1. empty state ------------------------------------------------------
        ck.ctx = "empty"
        empty = O.Db()
        check_view(doc, L, O.compute(empty, today), ck)
        check_db(doc, empty, ck)
        check_sys(doc, L, ck, file_name=None, mode_text=None, since=None)
        check_info(doc, L, empty, ck)
        ck.eq("log empty", read_log(doc, L), [])

        # 2. first import, mode 0 --------------------------------------------
        ck.ctx = "import main mode 0"
        t0 = now_local().replace(microsecond=0)
        res = doc.run_macro("modImport", "ImportFileSilent", (str(MAIN), 0))
        doc.recalc()
        db, rep = O.import_source(None, src_main, O.MODE_AUTO)
        check_report(res, rep, src_main, ck)
        check_db(doc, db, ck)
        check_sys(doc, L, ck, file_name=src_main.file_name, mode_text="Полная загрузка", since=t0)
        check_info(doc, L, db, ck)
        log_expect.insert(0, ("Полная загрузка", src_main.file_name, rep.total, rep.added, 0, 0,
                              rep.obj_count))
        exp_sel = O.month_label(today if db.first <= today <= db.last else db.last)
        ck.eq("sel_Month after import", doc.get_string("Сводка", L["dash"]["sel"]), exp_sel)
        ck.ctx = f"main, today={today}, sel={exp_sel}"
        check_view(doc, L, O.compute(db, today, O.month_start(today if db.first <= today <= db.last
                                                               else db.last)), ck)

        check_pdf_values(doc, O.compute(db, today, O.month_start(today if db.first <= today <= db.last
                                                                  else db.last)), ck, work, "oct")
        ck.ctx = "elapsed label"
        check_elapsed(doc, L, ck)

        # 3. other months via sel_Month --------------------------------------
        for m in (2, 3, 9, 1, 10):
            ms = dt.date(2026, m, 1)
            set_month(doc, L, O.month_label(ms))
            ck.ctx = f"main, sel={O.month_label(ms)}"
            check_view(doc, L, O.compute(db, today, ms), ck)
        set_month(doc, L, "Сентябрь 2026")
        ck.ctx = "main, sel=Сентябрь 2026 (rendered)"
        check_pdf_values(doc, O.compute(db, today, dt.date(2026, 9, 1)), ck, work, "sep")
        # an invalid selection falls back to the default month
        set_month(doc, L, "Ноябрь 2026")
        ck.ctx = "main, sel=Ноябрь 2026 (not in list)"
        check_view(doc, L, O.compute(db, today, None), ck)
        set_month(doc, L, O.month_label(today))

        # 3b. other "today" dates (c_Today replaced by a constant in the copy) --
        for fake, sel in ((dt.date(2026, 10, 1), 10), (dt.date(2026, 10, 2), 10),
                          (dt.date(2026, 10, 6), 10), (dt.date(2026, 10, 20), 10),
                          (dt.date(2026, 11, 1), 10), (dt.date(2026, 11, 1), None),
                          (dt.date(2026, 3, 17), 3), (dt.date(2026, 3, 18), 3),
                          (dt.date(2026, 3, 1), 3), (dt.date(2026, 2, 10), 2)):
            simulate_today(doc, L, fake)
            set_month(doc, L, O.month_label(dt.date(2026, sel, 1)) if sel else "")
            ck.ctx = f"main, simulated today={fake}, sel={sel}"
            check_view(doc, L, O.compute(db, fake, dt.date(2026, sel, 1) if sel else None), ck,
                       real_today=False)
        simulate_today(doc, L, None)
        set_month(doc, L, O.month_label(today))

        # 4. thresholds -------------------------------------------------------
        for cfg in (O.Cfg(0.5, 1.0, 2.0), O.Cfg(0.7, 0.9, 1.2), O.Cfg(1.8, 2.5, 2.5)):
            set_cfg(doc, L, cfg)
            for m in (10, 3):
                ms = dt.date(2026, m, 1)
                set_month(doc, L, O.month_label(ms))
                ck.ctx = f"main, cfg={cfg}, sel={O.month_label(ms)}"
                check_view(doc, L, O.compute(db, today, ms, cfg), ck)
        set_cfg(doc, L, default)
        set_month(doc, L, O.month_label(today))

        # 5. append import (mode 2) -------------------------------------------
        ck.ctx = "append update mode 2"
        t0 = now_local().replace(microsecond=0)
        res = doc.run_macro("modImport", "ImportFileSilent", (str(UPD), 2))
        doc.recalc()
        db2, rep2 = O.import_source(db, src_upd, O.MODE_APPEND)
        check_report(res, rep2, src_upd, ck)
        ck.eq("oracle: one conflict (СИКН №301 03.10)", [(c[0], c[1]) for c in rep2.conflict_list],
              [("СИКН №301", dt.date(2026, 10, 3))])
        check_db(doc, db2, ck)
        _n, _g, vals, _d, _s = read_db(doc)
        ck.eq("conflict kept base value", vals.get((dt.date(2026, 10, 3), "СИКН №301")),
              src_main.values[(dt.date(2026, 10, 3), "СИКН №301")])
        check_sys(doc, L, ck, file_name=src_upd.file_name, mode_text="Добавление новых", since=t0)
        check_info(doc, L, db2, ck)
        log_expect.insert(0, ("Добавление новых", src_upd.file_name, rep2.total, rep2.added,
                              rep2.same, rep2.conflicts, rep2.obj_count))
        sel2 = today if db2.first <= today <= db2.last else db2.last
        ck.eq("sel_Month after append", doc.get_string("Сводка", L["dash"]["sel"]), O.month_label(sel2))
        ck.ctx = f"after append, today={today}"
        check_view(doc, L, O.compute(db2, today, O.month_start(sel2)), ck)
        set_month(doc, L, "Март 2026")
        ck.ctx = "after append, sel=Март 2026"
        check_view(doc, L, O.compute(db2, today, dt.date(2026, 3, 1)), ck)
        set_month(doc, L, O.month_label(sel2))

        # 6. mode 0 on a non-empty base appends -------------------------------
        ck.ctx = "update again mode 0"
        res = doc.run_macro("modImport", "ImportFileSilent", (str(UPD), 0))
        doc.recalc()
        db3, rep3 = O.import_source(db2, src_upd, O.MODE_AUTO)
        check_report(res, rep3, src_upd, ck)
        ck.eq("mode 0 on non-empty base adds nothing", rep3.added, 0)
        check_db(doc, db3, ck)
        log_expect.insert(0, ("Добавление новых", src_upd.file_name, rep3.total, rep3.added,
                              rep3.same, rep3.conflicts, rep3.obj_count))

        # 7. replace import (mode 1) ------------------------------------------
        ck.ctx = "replace main mode 1"
        t0 = now_local().replace(microsecond=0)
        res = doc.run_macro("modImport", "ImportFileSilent", (str(MAIN), 1))
        doc.recalc()
        db4, rep4 = O.import_source(db3, src_main, O.MODE_REPLACE)
        check_report(res, rep4, src_main, ck)
        check_db(doc, db4, ck)
        _n, _g, vals4, _d, _s = read_db(doc)
        ck.eq("replace: base equals main file",
              (_n, len(vals4)), (src_main.objects, len(src_main.values)))
        check_sys(doc, L, ck, file_name=src_main.file_name, mode_text="Полная загрузка", since=t0)
        check_info(doc, L, db4, ck)
        log_expect.insert(0, ("Полная загрузка", src_main.file_name, rep4.total, rep4.added, 0, 0,
                              rep4.obj_count))
        ck.ctx = f"after replace, today={today}"
        sel4 = today if db4.first <= today <= db4.last else db4.last
        check_view(doc, L, O.compute(db4, today, O.month_start(sel4)), ck)

        # 8. errors do not touch the base ------------------------------------
        ck.ctx = "error paths"
        res = doc.run_macro("modImport", "ImportFileSilent", (str(work / "nope.xlsx"), 2))
        ck.true("missing file -> ERR", str(res).startswith("ERR"), str(res)[:200])
        check_db(doc, db4, ck)
        edges = edge_sources(work)
        res = doc.run_macro("modImport", "ImportFileSilent", (str(edges["novalues"]), 1))
        ck.true("file without values -> ERR", str(res).startswith("ERR"), str(res)[:200])
        check_db(doc, db4, ck)
        res = doc.run_macro("modImport", "ImportFileSilent", (str(book), 1))
        ck.true("the dashboard itself -> ERR", str(res).startswith("ERR"), str(res)[:200])
        check_db(doc, db4, ck)

        # 8b. edge-case source files -------------------------------------------
        ck.ctx = "edge append (Dec 2025, reordered columns, spelling, duplicates)"
        patch_edge_append(edges["append"], db4)
        src_e = O.read_source(edges["append"])
        res = doc.run_macro("modImport", "ImportFileSilent", (str(edges["append"]), 2))
        doc.recalc()
        db5, rep5 = O.import_source(db4, src_e, O.MODE_APPEND)
        check_report(res, rep5, src_e, ck)
        check_db(doc, db5, ck)
        log_expect.insert(0, ("Добавление новых", src_e.file_name, rep5.total, rep5.added,
                              rep5.same, rep5.conflicts, rep5.obj_count))
        for sel in (dt.date(2025, 12, 1), dt.date(2026, 1, 1)):
            set_month(doc, L, O.month_label(sel))
            ck.ctx = f"edge append, sel={O.month_label(sel)}"
            check_view(doc, L, O.compute(db5, today, sel), ck)

        ck.ctx = "edge replace with a small file"
        src_s = O.read_source(edges["small"])
        res = doc.run_macro("modImport", "ImportFileSilent", (str(edges["small"]), 1))
        doc.recalc()
        db6, rep6 = O.import_source(db5, src_s, O.MODE_REPLACE)
        check_report(res, rep6, src_s, ck)
        check_db(doc, db6, ck)
        log_expect.insert(0, ("Полная загрузка", src_s.file_name, rep6.total, rep6.added, 0, 0,
                              rep6.obj_count))
        sel6 = today if db6.first <= today <= db6.last else db6.last
        ck.eq("sel_Month after small replace", doc.get_string("Сводка", L["dash"]["sel"]),
              O.month_label(sel6))
        ck.ctx = "edge small, default month"
        check_view(doc, L, O.compute(db6, today, O.month_start(sel6)), ck)
        for fake in (dt.date(2026, 8, 11), dt.date(2026, 8, 12), dt.date(2026, 8, 13)):
            simulate_today(doc, L, fake)
            ck.ctx = f"edge small, simulated today={fake}"
            check_view(doc, L, O.compute(db6, fake, dt.date(2026, 8, 1)), ck, real_today=False)
        simulate_today(doc, L, None)

        # 9. clear -------------------------------------------------------------
        ck.ctx = "clear"
        doc.run_macro("modImport", "ClearDatabaseSilent")
        doc.recalc()
        check_db(doc, O.Db(), ck)
        check_view(doc, L, O.compute(O.Db(), today), ck)
        check_sys(doc, L, ck, file_name=None, mode_text=None, since=None)
        check_info(doc, L, O.Db(), ck)
        ck.eq("sel_Month cleared", doc.get_string("Сводка", L["dash"]["sel"]), "")
        log_expect.insert(0, ("Очистка базы", None, 0, 0, 0, 0, 0))

        # 10. import log -------------------------------------------------------
        ck.ctx = "import log"
        log = read_log(doc, L)
        ck.eq("log row count", len(log), len(log_expect))
        for i, (want, got) in enumerate(zip(log_expect, log)):
            mode_text, fname, total, added, same, conf, objs = want
            ck.eq(f"log[{i}] mode", got[2], mode_text)
            ck.eq(f"log[{i}] file", got[1], fname)
            ck.eq(f"log[{i}] numbers", got[3:8], [float(x) for x in (total, added, same, conf, objs)])
            ck.true(f"log[{i}] timestamp", isinstance(got[0], float), str(got[0]))
        doc.close()

    if out_dir:
        (work / "mismatches.json").write_text(
            json.dumps(ck.fails, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    else:
        shutil.rmtree(work, ignore_errors=True)
    return ck


def summarize(ck: Checker, limit: int = 400) -> str:
    lines = [f"{ck.count} checks, {len(ck.fails)} mismatches"]
    for f in ck.fails[:limit]:
        lines.append(f"[{f['ctx']}] {f['what']}: actual={f['actual']!r} expected={f['expected']!r}")
    return "\n".join(lines)


def test_integration():
    out = os.environ.get("LHOS_E2E_OUT")
    ck = run(Path(out) if out else None)
    assert not ck.fails, summarize(ck)


if __name__ == "__main__":
    if not HAVE_LO:
        print("SKIP: LibreOffice/uno not available")
        sys.exit(0)
    t = time.time()
    out = os.environ.get("LHOS_E2E_OUT")
    ck = run(Path(out) if out else None)
    print(summarize(ck))
    print(f"{time.time() - t:.1f} s")
    sys.exit(1 if ck.fails else 0)
