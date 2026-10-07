"""Independent reference model ("oracle") of the LHOS dashboard.

Re-implements, in plain Python and *without* looking at the workbook's
formulas, what the dashboard must show:

* reading a source workbook in the user's format (one sheet per month, a
  "Дата" header cell, object names in the bottom header row, product groups in
  the row above, dates below, values in ppm; "0,45" typed as text is a number,
  "<0,05" counts as 0.05, "н/д" is skipped; sheets without "Дата" are
  ignored; a name may sit in a cell merged over both header rows (then it has
  no product group); service columns such as «Среднее», «Итого», «Макс» are not
  objects; the table ends at the last date / last header cell);
* the database after an import: first import fills it, later imports either
  replace everything (mode 1) or only add values that are not in the base yet
  (mode 2; differing values are conflicts and the base keeps its value);
  the same date of an object on several sheets is "last sheet wins" and is
  reported;
* the import report text (summary first, details after a blank line);
* the dashboard for a given TODAY and selected month:
    - current month: latest value on/before TODAY (if TODAY has none, the
      previous day(s) of the month; on the 1st the previous day may be in the
      previous month);
    - past month: last value of that month;
    - zones are decided on the value rounded to 2 decimals (what the table
      shows): v <= green -> 1, <= yellow -> 2, <= limit -> 3, > limit -> 4;
    - change (24h) = displayed value minus the displayed previous measured
      value of the object (both rounded to 2 decimals), rounded to 2 decimals;
    - table sorted by value (desc; values equal to 6 decimals tie), ties in
      base order, objects without value last; "*" when an object's value is
      older than the newest value date;
    - KPIs (max + object, average, zone counts, total, overall state);
    - trend: daily values of the selected month up to the reference date;
    - max-per-day series for the same days;
    - month list: months from the first to the last month with data, newest
      first.

Only the standard library and openpyxl are used.
"""

from __future__ import annotations

import calendar
import datetime as dt
import math
import os
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_EVEN, ROUND_HALF_UP, Decimal

import openpyxl

MONTHS_RU = ("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август",
             "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
MAX_OBJ = 40             # objects the dashboard handles
TABLE_ROWS = 24          # rows of the table on «Сводка»
MAX_EXAMPLES = 3         # conflict examples in the import report
MAX_NEW_LISTED = 5       # new object names listed in the import report
MAX_SHEETS_LISTED = 12   # sheet names listed in the import report

MODE_AUTO, MODE_REPLACE, MODE_APPEND = 0, 1, 2

STATE_TEXT = ("НЕТ ДАННЫХ", "НОРМА", "ВНИМАНИЕ", "РИСК", "ПРЕВЫШЕНИЕ")
ROW_STATUS = ("●  НЕТ ДАННЫХ", "●  НОРМА", "●  ВНИМАНИЕ", "●  РИСК", "●  ПРЕВЫШЕНИЕ")

# Dates outside 2000-01-01 .. 2099-12-31 are not dates of the data.
DATE_MIN, DATE_MAX = dt.date(2000, 1, 1), dt.date(2099, 12, 31)
EPOCH = dt.date(1899, 12, 30)

# Column headers that are summaries, not objects (compared in lower case).
SERVICE_PREFIXES = ("среднее", "средн.", "итог", "всего", "сумма", "максимум", "минимум",
                    "max", "min", "примечан")
SERVICE_NAMES = ("макс", "мин")


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def clean(v) -> str:
    if v is None:
        return ""
    s = str(v).replace("\r", " ").replace("\n", " ").replace("\t", " ").replace("\xa0", " ")
    return " ".join(s.split())


def xround(x: float, n: int) -> float:
    """Excel ROUND (half away from zero, on the shortest decimal representation)."""
    q = Decimal(repr(abs(x))).quantize(Decimal(1).scaleb(-n), rounding=ROUND_HALF_UP)
    return float(-q if x < 0 else q)


def as_date(v) -> dt.date | None:
    """A date in the date column: a date/datetime, a date serial or "dd.mm.yyyy"."""
    d = None
    if isinstance(v, dt.datetime):
        d = v.date()
    elif isinstance(v, dt.date):
        d = v
    elif isinstance(v, (int, float)) and not isinstance(v, bool):
        if math.isfinite(v) and 36526 <= v < 73051:
            d = EPOCH + dt.timedelta(days=math.floor(v))
    elif isinstance(v, str):
        m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", clean(v))
        if m:
            try:
                d = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                d = None
    if d is None or not DATE_MIN <= d <= DATE_MAX:
        return None
    return d


_PLAIN_NUMBER = re.compile(r"[+-]?(?=\.?\d)\d*\.?\d*")


def parse_number(v):
    """(kind, value): kind 1 numeric, 0 empty, -1 non-numeric (text, date, bool, error)."""
    if v is None:
        return 0, None
    if isinstance(v, (bool, dt.date, dt.datetime, dt.time)):
        return -1, None
    if isinstance(v, (int, float)):
        return (1, float(v)) if math.isfinite(v) else (-1, None)
    s = clean(v).replace(" ", "")
    if not s:
        return 0, None
    s = s.replace(",", ".")
    s = s.removeprefix("<")        # "<0,05" — below the detection limit: the limit
    if not _PLAIN_NUMBER.fullmatch(s):
        return -1, None
    return 1, float(s)


def is_service_header(name: str) -> bool:
    s = name.lower()
    return s.startswith(SERVICE_PREFIXES) or s in SERVICE_NAMES


def month_start(d: dt.date) -> dt.date:
    return d.replace(day=1)


def month_end(d: dt.date) -> dt.date:
    return d.replace(day=calendar.monthrange(d.year, d.month)[1])


def month_label(d: dt.date) -> str:
    return f"{MONTHS_RU[d.month - 1]} {d.year}"


def add_months(d: dt.date, n: int) -> dt.date:
    y, m = divmod(d.year * 12 + d.month - 1 + n, 12)
    return dt.date(y, m + 1, 1)


def fmt_num(x: float, min_dec: int = 1, max_dec: int = 2, comma: str = ",") -> str:
    """Excel number format like ``0.0#`` (round half away from zero)."""
    q = Decimal(repr(abs(x))).quantize(Decimal(1).scaleb(-max_dec), rounding=ROUND_HALF_UP)
    s = f"{q:.{max_dec}f}"
    ip, fp = s.split(".")
    while len(fp) > min_dec and fp.endswith("0"):
        fp = fp[:-1]
    out = ip + ((comma + fp) if fp else "")
    neg = x < 0 and q != 0
    return ("-" if neg else "") + out


def fmt_val(x: float, comma: str = ",") -> str:
    """Excel number format ``0.0?``: 1-2 decimals, a missing 2nd decimal is a space."""
    s = fmt_num(x, 1, 2, comma)
    return s + " " if len(s.split(comma)[1]) == 1 else s


def fmt_change(x: float, comma: str = ",") -> str:
    """Excel number format ``+0.0?;-0.0?;0.0?`` (section chosen by the sign)."""
    if x > 0:
        return "+" + fmt_val(x, comma)
    if x < 0:
        return "-" + fmt_val(-x, comma)
    return fmt_val(0.0, comma)


def fmt_fixed(x: float, dec: int, comma: str = ",") -> str:
    """FIXED(x, dec) for small numbers (no thousands)."""
    q = Decimal(repr(abs(x))).quantize(Decimal(1).scaleb(-dec), rounding=ROUND_HALF_UP)
    s = f"{q:.{dec}f}".replace(".", comma)
    return ("-" if x < 0 and q != 0 else "") + s


def fmt_threshold(x: float) -> str:
    """A threshold with 1 decimal, or 2 when it has a 2nd decimal (1,5 / 1,25)."""
    return fmt_fixed(x, 1 if xround(x, 1) == x else 2)


def fmt_date(d: dt.date | None) -> str:
    return d.strftime("%d.%m.%Y") if d else "—"


def num_text(v: float) -> str:
    """A value as the import report writes it: rounded to 4 decimals (half to even),
    shortest form, decimal comma."""
    q = Decimal(repr(v)).quantize(Decimal("0.0001"), rounding=ROUND_HALF_EVEN)
    s = format(q.normalize(), "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    if s in ("-0", ""):
        s = "0"
    return s.replace(".", ",")


# ---------------------------------------------------------------------------
# Source file
# ---------------------------------------------------------------------------

@dataclass
class Source:
    file_name: str
    objects: list = field(default_factory=list)          # first-seen order
    groups: dict = field(default_factory=dict)
    values: dict = field(default_factory=dict)           # (date, name) -> float; last wins
    sheets_with_values: list = field(default_factory=list)
    empty_sheets: int = 0
    text_skipped: int = 0
    dup_count: int = 0                                    # values repeated on another sheet
    dup_example: tuple | None = None                      # (earlier sheet, later sheet, date)
    order: dict = field(default_factory=dict)             # name -> first-seen index

    @property
    def total(self) -> int:
        return len(self.values)

    @property
    def first(self) -> dt.date:
        return min(d for d, _ in self.values)

    @property
    def last(self) -> dt.date:
        return max(d for d, _ in self.values)


class _Merges:
    def __init__(self, ws):
        self.ws = ws
        self.ranges = list(ws.merged_cells.ranges)

    def area(self, row: int, col: int):
        """(min_row, min_col, max_row, max_col) of the merge containing the cell."""
        for m in self.ranges:
            if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
                return m.min_row, m.min_col, m.max_row, m.max_col
        return row, col, row, col

    def top_left(self, row: int, col: int):
        r, c, _r, _c = self.area(row, col)
        return self.ws.cell(r, c).value


def _find_date_header(ws):
    for row in ws.iter_rows(min_row=1, max_row=40, max_col=26):
        for c in row:
            if isinstance(c.value, str) and clean(c.value).lower() in ("дата", "дата:"):
                return c
    return None


def read_source(path) -> Source:
    wb = openpyxl.load_workbook(path, data_only=True)
    src = Source(os.path.basename(str(path)))
    keys: dict[str, str] = {}      # upper-case key -> canonical name
    setter: dict = {}              # (date, name) -> sheet title that set it last
    for ws in wb.worksheets:
        hdr = _find_date_header(ws)
        if hdr is None:
            continue
        mg = _Merges(ws)
        top, date_col, bottom, _right = mg.area(hdr.row, hdr.column)
        name_row, group_row = bottom, (top if bottom > top else None)
        # the table ends at the last date and at the last header cell
        last_row = max((r for r in range(1, ws.max_row + 1)
                        if ws.cell(r, date_col).value is not None), default=1)
        last_col = 1
        for hr in (name_row, group_row):
            if hr is None:
                continue
            c = max((c for c in range(1, ws.max_column + 1) if ws.cell(hr, c).value is not None),
                    default=1)
            last_col = max(last_col, mg.area(hr, c)[3])
        if last_row <= name_row or last_col <= date_col:
            src.empty_sheets += 1
            continue
        cols = {}
        for col in range(date_col + 1, last_col + 1):
            name_area = mg.area(name_row, col)
            nm = clean(ws.cell(name_area[0], name_area[1]).value) if name_area[1] == col else ""
            if not nm or is_service_header(nm):
                continue
            grp = ""
            if group_row is not None:
                group_area = mg.area(group_row, col)
                if group_area != name_area:      # a name merged over both rows has no group
                    grp = clean(ws.cell(group_area[0], group_area[1]).value)
            key = nm.upper()
            if key not in keys:
                keys[key] = nm
                src.order[nm] = len(src.objects)
                src.objects.append(nm)
                src.groups[nm] = grp
            elif not src.groups.get(keys[key]):
                src.groups[keys[key]] = grp
            cols[col] = keys[key]
        n = 0
        for r in range(name_row + 1, last_row + 1):
            d = as_date(ws.cell(r, date_col).value)
            if d is None:
                continue
            for col, nm in cols.items():
                kind, x = parse_number(ws.cell(r, col).value)
                if kind == 1:
                    k = (d, nm)
                    prev = setter.get(k)
                    if prev is not None and prev != ws.title:
                        src.dup_count += 1
                        if src.dup_example is None:
                            src.dup_example = (prev, ws.title, d)
                    setter[k] = ws.title
                    src.values[k] = x
                    n += 1
                elif kind == -1:
                    src.text_skipped += 1
        if n:
            src.sheets_with_values.append(ws.title)
        else:
            src.empty_sheets += 1
    return src


# ---------------------------------------------------------------------------
# Database and import
# ---------------------------------------------------------------------------

@dataclass
class Db:
    objects: list = field(default_factory=list)
    groups: dict = field(default_factory=dict)
    values: dict = field(default_factory=dict)

    def copy(self) -> Db:
        return Db(list(self.objects), dict(self.groups), dict(self.values))

    @property
    def empty(self) -> bool:
        return not self.values

    @property
    def first(self) -> dt.date | None:
        return min((d for d, _ in self.values), default=None)

    @property
    def last(self) -> dt.date | None:
        return max((d for d, _ in self.values), default=None)

    def dates(self) -> list:
        if self.empty:
            return []
        n = (self.last - self.first).days + 1
        return [self.first + dt.timedelta(days=i) for i in range(n)]

    def get(self, d: dt.date, name: str):
        return self.values.get((d, name))

    def without_date(self, d: dt.date) -> Db:
        """The base with every value of one date removed."""
        out = self.copy()
        out.values = {k: v for k, v in self.values.items() if k[0] != d}
        return out


@dataclass
class ImportReport:
    mode: int
    total: int
    added: int
    same: int
    conflicts: int
    conflict_list: list
    new_objects: list
    obj_count: int
    first: dt.date
    last: dt.date


def import_source(db: Db | None, src: Source, mode: int):
    db = db or Db()
    if mode == MODE_AUTO:
        mode = MODE_REPLACE if db.empty else MODE_APPEND
    if mode == MODE_REPLACE:
        new = Db(list(src.objects), dict(src.groups), dict(src.values))
        rep = ImportReport(mode, src.total, src.total, 0, 0, [], [], len(new.objects),
                           new.first, new.last)
        return new, rep
    new = db.copy()
    keys = {o.upper(): o for o in new.objects}
    new_objects = []
    canon = {}
    for nm in src.objects:
        if nm.upper() in keys:
            canon[nm] = keys[nm.upper()]
            if not new.groups.get(canon[nm]):
                new.groups[canon[nm]] = src.groups.get(nm, "")
        else:
            new.objects.append(nm)
            new.groups[nm] = src.groups.get(nm, "")
            keys[nm.upper()] = nm
            canon[nm] = nm
            new_objects.append(nm)
    added = same = 0
    conflicts = []
    # by date, then in the order the objects appear in the file
    for (d, nm), x in sorted(src.values.items(), key=lambda kv: (kv[0][0], src.order[kv[0][1]])):
        k = (d, canon[nm])
        if k not in new.values:
            new.values[k] = x
            added += 1
        elif abs(new.values[k] - x) < 1e-7:
            same += 1
        else:
            conflicts.append((canon[nm], d, new.values[k], x))
    rep = ImportReport(mode, src.total, added, same, len(conflicts), conflicts, new_objects,
                       len(new.objects), new.first, new.last)
    return new, rep


def _grouped(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def report_text(rep: ImportReport, src: Source) -> str:
    """The import report (lines joined by "\\n"; thousands grouped with a space)."""
    lines = [f"Файл: {src.file_name}"]
    if rep.mode == MODE_REPLACE:
        lines += ["Режим: полная загрузка (база заполнена заново)",
                  f"Загружено значений: {_grouped(rep.added)}"]
    else:
        lines += ["Режим: добавление новых значений",
                  f"Добавлено новых значений: {_grouped(rep.added)}",
                  f"Уже были в базе: {_grouped(rep.same)}"]
        if rep.conflicts:
            lines.append(f"Отличаются от базы (оставлены как в базе): {rep.conflicts}")
        if rep.new_objects:
            lines.append(f"Новых объектов: {len(rep.new_objects)}")
    lines.append(f"Объектов в базе: {rep.obj_count}")
    lines.append(f"Период данных в базе: {fmt_date(rep.first)} – {fmt_date(rep.last)}")
    if rep.obj_count > MAX_OBJ:
        lines.append(f"Внимание: на дашборд выводятся первые {MAX_OBJ} объектов из {rep.obj_count}.")
    if src.text_skipped:
        lines.append(f"Пропущено нечисловых значений: {src.text_skipped}")
    if src.dup_count:
        a, b, d = src.dup_example
        lines.append(f"Даты повторяются на разных листах: {src.dup_count} знач. "
                     f"(например, «{a}» и «{b}», {fmt_date(d)}); взяты значения с последнего листа.")
    lines.append("")
    sheets = src.sheets_with_values
    s = f"Значений в файле: {_grouped(rep.total)}; листов с данными: {len(sheets)}"
    if 0 < len(sheets) <= MAX_SHEETS_LISTED:
        s += f" ({', '.join(sheets)})"
    if src.empty_sheets:
        s += f"; листов без значений: {src.empty_sheets}"
    lines.append(s)
    if rep.mode == MODE_APPEND:
        if rep.conflicts:
            lines.append("Расхождения с базой:")
            for nm, d, old, new in rep.conflict_list[:MAX_EXAMPLES]:
                lines.append(f"   • {nm}, {fmt_date(d)}: в базе {num_text(old)}, в файле {num_text(new)}")
            if rep.conflicts > MAX_EXAMPLES:
                lines.append(f"   … и еще {rep.conflicts - MAX_EXAMPLES}")
            lines.append("Чтобы заменить их, выполните импорт с полной перезагрузкой.")
        if rep.new_objects:
            more = " …" if len(rep.new_objects) > MAX_NEW_LISTED else ""
            lines.append("Новые объекты: " + ", ".join(rep.new_objects[:MAX_NEW_LISTED]) + more)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Dashboard view
# ---------------------------------------------------------------------------

@dataclass
class Cfg:
    green: float = 1.5
    yellow: float = 2.3
    limit: float = 3.0


@dataclass
class ObjView:
    idx: int                 # 1-based base column
    name: str
    group: str
    value_date: dt.date | None
    value: float | None
    prev_date: dt.date | None
    prev_value: float | None
    change: float | None
    zone: int
    stale: bool


@dataclass
class View:
    today: dt.date
    month_start: dt.date
    month_end: dt.date
    is_current: bool
    ref_date: dt.date
    label: str
    objects: list
    sorted: list
    data_date: dt.date | None
    cnt_valued: int
    cnt_green: int
    cnt_yellow: int
    cnt_red: int            # red + over-limit
    cnt_over: int
    cnt_nodata: int
    max_val: float | None
    max_name: str
    avg_val: float | None
    state: int
    obj_count: int
    db_count: int
    stale_count: int
    days: list               # dates of the selected month (all days)
    trend: list              # per sorted row: list of 31 (value|None)
    maxday: list             # 31 (value|None)
    last_max_date: dt.date | None
    months: list             # [(start, label)] newest first
    cfg: Cfg


def zone_of(v: float | None, cfg: Cfg) -> int:
    """Zone of a value as displayed (rounded to 2 decimals)."""
    if v is None:
        return 0
    r = xround(v, 2)
    if r <= cfg.green:
        return 1
    if r <= cfg.yellow:
        return 2
    if r <= cfg.limit:
        return 3
    return 4


def change_of(value: float, prev: float) -> float:
    """Change between the displayed (2-decimal) values, rounded to 2 decimals."""
    return xround(xround(value, 2) - xround(prev, 2), 2)


def month_list(db: Db) -> list:
    if db.empty:
        return []
    out = []
    m = month_start(db.last)
    while m >= month_start(db.first):
        out.append((m, month_label(m)))
        m = add_months(m, -1)
    return out


def default_month(db: Db, today: dt.date) -> dt.date:
    """Month the dashboard shows when nothing (valid) is selected."""
    return month_start(db.last) if not db.empty else month_start(today)


def compute(db: Db, today: dt.date, sel: dt.date | None = None, cfg: Cfg | None = None) -> View:
    cfg = cfg or Cfg()
    ms = month_start(sel) if sel else default_month(db, today)
    me = month_end(ms)
    is_current = ms <= today <= me
    ref = today if is_current else me
    # current month: today's value, or the previous day's (which on the 1st of a
    # month lies in the previous month); a past month: its last value
    low = min(ms, ref - dt.timedelta(days=1)) if is_current else ms
    objs = []
    names = db.objects[:MAX_OBJ]
    for i, nm in enumerate(names, 1):
        vd = None
        d = ref
        while d >= low:
            if db.get(d, nm) is not None:
                vd = d
                break
            d -= dt.timedelta(days=1)
        val = db.get(vd, nm) if vd else None
        pd = pv = chg = None
        if vd is not None:
            earlier = [d for (d, n) in db.values if n == nm and d < vd]
            if earlier:
                pd = max(earlier)
                pv = db.get(pd, nm)
                chg = change_of(val, pv)
        objs.append(ObjView(i, nm, db.groups.get(nm, ""), vd, val, pd, pv, chg,
                            zone_of(val, cfg), False))
    data_date = max((o.value_date for o in objs if o.value_date), default=None)
    for o in objs:
        o.stale = o.value is not None and data_date is not None and o.value_date < data_date
    valued = [o for o in objs if o.value is not None]
    nodata = [o for o in objs if o.value is None]
    srt = sorted(valued, key=lambda o: (-xround(o.value, 6), o.idx)) + nodata
    cnt = {z: sum(1 for o in objs if o.zone == z) for z in range(5)}
    if not valued:
        state = 0
    elif cnt[4]:
        state = 4
    elif cnt[3]:
        state = 3
    elif cnt[2]:
        state = 2
    else:
        state = 1
    days = [ms + dt.timedelta(days=i) for i in range((me - ms).days + 1)]
    trend = []
    for o in srt:
        row = []
        for k in range(31):
            if k < len(days) and days[k] <= ref:
                row.append(db.get(days[k], o.name))
            else:
                row.append(None)
        trend.append(row)
    maxday = []
    for k in range(31):
        if k < len(days) and days[k] <= ref:
            vals = [db.get(days[k], nm) for nm in names]
            vals = [v for v in vals if v is not None]
            maxday.append(max(vals) if vals else None)
        else:
            maxday.append(None)
    last_max = max((days[k] for k in range(len(days)) if maxday[k] is not None), default=None)
    return View(
        today=today, month_start=ms, month_end=me, is_current=is_current, ref_date=ref,
        label=month_label(ms), objects=objs, sorted=srt, data_date=data_date,
        cnt_valued=len(valued), cnt_green=cnt[1], cnt_yellow=cnt[2], cnt_red=cnt[3] + cnt[4],
        cnt_over=cnt[4], cnt_nodata=cnt[0], max_val=max((o.value for o in valued), default=None),
        max_name=srt[0].name if valued else "",
        avg_val=(sum(o.value for o in valued) / len(valued)) if valued else None,
        state=state, obj_count=len(names), db_count=len(db.objects), stale_count=sum(1 for o in objs if o.stale),
        days=days, trend=trend, maxday=maxday, last_max_date=last_max, months=month_list(db),
        cfg=cfg)


# ---------------------------------------------------------------------------
# Expected visible texts on the dashboard (ru-RU number formatting)
# ---------------------------------------------------------------------------

def zone_label(z: int, cfg: Cfg) -> str:
    g, y, lim = (fmt_threshold(x) for x in (cfg.green, cfg.yellow, cfg.limit))
    return ("—", f"0 – {g} ppm", f"{g} – {y} ppm", f"> {y} ppm", f"> {lim} ppm")[z]


def state_sub(v: View) -> str:
    if v.state == 0:
        if v.obj_count == 0:
            return "Загрузите файл с данными"
        if v.is_current:
            return ("Нет значений за сегодня и вчера" if v.today.day == 1
                    else "Нет значений с начала месяца")
        return "Нет значений за " + v.label.lower()
    return ("", "Превышений не зафиксировано",
            f"В желтой зоне: {v.cnt_yellow} СИКН", f"В красной зоне: {v.cnt_red} СИКН",
            f"Выше норматива: {v.cnt_over} СИКН")[v.state]


def footer(v: View) -> str:
    return "СОСТОЯНИЕ: " + ("НЕТ ДАННЫХ" if v.state == 0 else
                            "ВСЕ В НОРМЕ" if v.state == 1 else "ТРЕБУЕТСЯ ВНИМАНИЕ")


def legend(count: int, valued: int) -> str:
    pct = Decimal(100 * count / max(1, valued)).quantize(Decimal(1), rounding=ROUND_HALF_UP)
    return f"{count} ({pct}%)"


def table_note(v: View) -> str:
    stale = "* — значение за предыдущую дату"
    if v.obj_count == 0:
        return "База пуста — нажмите «ИМПОРТ ДАННЫХ»"
    if v.db_count > TABLE_ROWS:
        return f"Показаны {TABLE_ROWS} из {v.db_count} СИКН" + ("; " + stale if v.stale_count else "")
    return stale if v.stale_count else ""


def arrow_of(change: float | None) -> str:
    if change is None:
        return ""
    return "↑" if change > 0.00001 else "↓" if change < -0.00001 else "—"


def table_texts(v: View) -> list:
    """24 rows x (№, name, value, status, zone, change, arrow) as displayed
    (the value without the "*" that the conditional format adds)."""
    rows = []
    for k in range(TABLE_ROWS):
        if k >= len(v.sorted):
            rows.append(("", "", "", "", "", "", ""))
            continue
        o = v.sorted[k]
        # формат «0.0?_*»: место под «*» занято пробелом и у значений без отметки
        val = "—" if o.value is None else fmt_val(o.value) + " "
        chg = "" if o.change is None else fmt_change(o.change)
        rows.append((str(k + 1), o.name, val, ROW_STATUS[o.zone], zone_label(o.zone, v.cfg),
                     chg, arrow_of(o.change)))
    return rows


def rendered_value(o: ObjView) -> str:
    """The value cell as rendered (conditional number format adds "*" when stale)."""
    if o.value is None:
        return "—"
    return fmt_val(o.value) + ("*" if o.stale else "")


def dash_texts(v: View) -> dict:
    """Expected strings of the KPI / header cells, keyed by a logical id."""
    return {
        "state_text": STATE_TEXT[v.state],
        "state_sub": state_sub(v),
        "max": "—" if v.max_val is None else fmt_num(v.max_val),
        "max_unit": "" if v.max_val is None else "ppm",
        "max_name": v.max_name,
        "avg": "—" if v.avg_val is None else fmt_num(v.avg_val),
        "avg_unit": "" if v.avg_val is None else "ppm",
        "avg_sub": "" if v.cnt_valued == 0 else f"по {v.cnt_valued} СИКН",
        "green": str(v.cnt_green),
        "yellow": str(v.cnt_yellow),
        "red": str(v.cnt_red),
        "total": str(v.obj_count),
        "total_sub": f"нет данных: {v.cnt_nodata}" if v.cnt_nodata > 0 else "",
        "data_date": fmt_date(v.data_date),
        "title": "СОДЕРЖАНИЕ ЛХОС ПО СИКН — " + v.label.upper(),
        "note": table_note(v),
        "footer": footer(v),
        "leg_g": legend(v.cnt_green, v.cnt_valued),
        "leg_y": legend(v.cnt_yellow, v.cnt_valued),
        "leg_r": legend(v.cnt_red, v.cnt_valued),
        "donut_total": str(v.obj_count),
        "zone_r": zone_label(3, v.cfg),
        "zone_y": zone_label(2, v.cfg),
        "zone_g": zone_label(1, v.cfg),
        "limit": f"≤ {fmt_num(v.cfg.limit)} ppm",
        "scale_limit": fmt_num(v.cfg.limit),
        "scale_yellow": fmt_num(v.cfg.yellow),
        "scale_green": fmt_num(v.cfg.green),
    }


def file_label(name: str | None) -> str:
    """The file name in the control bar (long names shortened)."""
    if not name:
        return "—"
    return name[:29] + "…" if len(name) > 31 else name


def source_label(name: str | None) -> str:
    """The footer text about the data source (long names shortened)."""
    if not name:
        return "Данные не загружены — нажмите «Импорт данных»"
    return "Источник данных: " + (name[:42] + "…" if len(name) > 44 else name)


def elapsed_text(delta: dt.timedelta | None) -> str:
    """«Обновление: …» for the time since the last import (None: never imported)."""
    if delta is None:
        return "Данные еще не загружались"
    mins = int(delta.total_seconds() // 60)
    if delta < dt.timedelta(minutes=1):
        return "Обновление: только что"
    if delta < dt.timedelta(hours=1):
        return f"Обновление: {mins} мин назад"
    if delta < dt.timedelta(days=1):
        return f"Обновление: {mins // 60} ч {mins % 60} мин назад"
    return f"Обновление: {delta.days} дн {int(delta.total_seconds() // 3600) % 24} ч назад"


def freshness(delta: dt.timedelta | None) -> str | None:
    """Colour of the freshness dot: None (never imported), green < 1 day,
    yellow < 3 days, red otherwise."""
    if delta is None:
        return None
    if delta < dt.timedelta(days=1):
        return "green"
    if delta < dt.timedelta(days=3):
        return "yellow"
    return "red"


if __name__ == "__main__":  # quick manual check
    import sys
    s = read_source(sys.argv[1])
    db, rep = import_source(None, s, 0)
    print(report_text(rep, s))
    v = compute(db, dt.date.today())  # noqa: DTZ011 - Excel TODAY() is local
    print(v.label, v.state, v.max_val, v.max_name, v.avg_val, v.data_date)
    for row in table_texts(v):
        print(row)
