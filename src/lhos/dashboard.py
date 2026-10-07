"""Сборка книги-дашборда ЛХОС (xlsxwriter).

Листы (кодовые имена VBA в скобках):
  Сводка (shDash)       — дашборд, как на макете;
  СИКН, Тренды, Тревоги, Отчеты (shObj, shTrend, shAlarm, shReport) — разделы-заглушки;
  Настройки (shSettings) — пороги зон, подписи, сведения о загрузках, журнал;
  БД (shDB)             — база данных, заполняется макросом импорта;
  Расчет (shCalc)       — скрытый лист с формулами.

Все показатели считаются формулами по листу «БД», поэтому дашборд работает и
пересчитывается без макросов; макросы нужны только для импорта.
"""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from .canvas import Canvas
from .theme import C, FONT, MONTHS_RU, ZONE_COLOR

# ----------------------------------------------------------------------------
# Константы структуры
# ----------------------------------------------------------------------------
MAX_OBJ = 40            # объектов, обрабатываемых формулами
DB_ROWS = 3700          # дней в базе (~10 лет)
TABLE_ROWS = 24         # строк в таблице дашборда
LOG_ROWS = 50           # строк журнала загрузок (совпадает с VBA)

S_DASH, S_OBJ, S_TREND, S_ALARM, S_REPORT, S_SET, S_DB, S_CALC = (
    "Сводка", "СИКН", "Тренды", "Тревоги", "Отчеты", "Настройки", "БД", "Расчет")

NAV = [  # (подпись, лист, иконка, кодовое имя)
    ("СВОДКА", S_DASH, "nav_dashboard", "shDash"),
    ("СИКН", S_OBJ, "nav_sikn", "shObj"),
    ("ТРЕНДЫ", S_TREND, "nav_trends", "shTrend"),
    ("ТРЕВОГИ", S_ALARM, "nav_alarms", "shAlarm"),
    ("ОТЧЕТЫ", S_REPORT, "nav_reports", "shReport"),
    ("НАСТРОЙКИ", S_SET, "nav_settings", "shSettings"),
]

# Лист «Расчет»: строки (1-based)
OBJ_R0 = 46                      # объекты 46..85
DAYNUM_R, DAYDATE_R = 90, 91     # номер дня / дата дня месяца (столбцы L..AP)
SORT_R0 = 92                     # сортированные строки 92..131
MAXDAY_R, LIM_R, YEL_R, GRN_R, LAST_R, CAT_R = 134, 135, 136, 137, 138, 139
MONTH_R0, MONTH_N = 142, 60      # список месяцев 142..201
TR_C0 = 11                       # столбец L (0-based) — первый день тренда
TR_C1 = TR_C0 + 30               # столбец AP — 31-й день

W, H = 1390, 1006                # размер макета дашборда, px


def _q(sheet: str) -> str:
    return f"'{sheet}'"


def _calc(ref: str) -> str:
    return f"{_q(S_CALC)}!{ref}"


def _months_choose(expr: str) -> str:
    names = ",".join(f'"{m}"' for m in MONTHS_RU)
    return f"CHOOSE(MONTH({expr}),{names})"


@dataclass
class BuildResult:
    path: str
    names: dict


# ----------------------------------------------------------------------------
# Общие элементы оформления
# ----------------------------------------------------------------------------
class Assets:
    def __init__(self, icons_dir: str):
        self.dir = icons_dir

    def button(self, name: str) -> str:
        return self.icon("btn_" + name)

    def icon(self, name: str) -> str:
        p = os.path.join(self.dir, name + ".png")
        if not os.path.exists(p):
            raise FileNotFoundError(p)
        return p


def draw_header(cv: Canvas, assets: Assets) -> None:
    cv.fill((0, 0, cv.width, 80), C["header"])
    cv.hline(0, cv.width, 80, C["divider"])
    cv.image(18, 17, assets.icon("logo_mark"), scale=0.48)
    cv.put((66, 20, 196, 62), "=cfg_Company", bold=True, size=17, color=C["logo"],
           italic=True)
    cv.vline(206, 16, 72, C["divider"])
    cv.put((224, 14, 1080, 46), "=cfg_Title", bold=True, size=16, color=C["text"])
    cv.put((224, 48, 1080, 70), "=cfg_Subtitle", bold=True, size=10, color=C["green_soft"])
    cv.put((1110, 12, 1370, 44), "=TODAY()", bold=True, size=20, color=C["text"],
           align="right", num="dd.mm.yyyy")


def draw_nav(cv: Canvas, assets: Assets, active: int) -> None:
    cv.fill((10, 92, 110, 948), C["nav"])
    cv.frame((10, 92, 110, 948), C["border"])
    for k, (label, sheet, icon, _code) in enumerate(NAV):
        y0 = 100 + k * 90
        url = f"internal:{_q(sheet)}!A1"
        if k == active:
            cv.fill((16, y0, 104, y0 + 82), C["nav_active"])
            cv.frame((16, y0, 104, y0 + 82), C["divider"])
        cv.put((16, y0 + 50, 104, y0 + 74), label, bold=True, size=8.5, align="center",
               color=C["text"] if k == active else C["text2"], url=url,
               url_tip=label.capitalize())
        if k < len(NAV) - 1:
            cv.hline(26, 94, y0 + 86, C["line"])
        variant = "_active" if k == active else "_normal"
        cv.image(44, y0 + 12, assets.icon(icon + variant), scale=0.5, url=url,
                 tip=label.capitalize())


def button(cv: Canvas, x: int, y: int, path: str, macro: str) -> None:
    """Кнопка-картинка с макросом: атрибут macro проставляется при постобработке,
    а при открытии книги макрос дополнительно назначается из VBA по тексту btn:…"""
    cv.image(x, y, path, scale=0.5, descr="btn:" + macro)


# ----------------------------------------------------------------------------
# Лист «Расчет»
# ----------------------------------------------------------------------------
PARAMS = [  # (имя, подпись, формула, массив?, формат)
    ("c_Today", "Сегодня", "=TODAY()", False, "dd.mm.yyyy"),
    ("c_FirstDate", "Первая дата в базе", '=IF(COUNT(db_Dates)=0,"",MIN(db_Dates))', False, "dd.mm.yyyy"),
    ("c_LastDate", "Последняя дата в базе", '=IF(COUNT(db_Dates)=0,"",MAX(db_Dates))', False, "dd.mm.yyyy"),
    ("c_LastMonth", "Последний месяц", '=IF(c_LastDate="","",DATE(YEAR(c_LastDate),MONTH(c_LastDate),1))', False, "dd.mm.yyyy"),
    ("c_FirstMonth", "Первый месяц", '=IF(c_FirstDate="","",DATE(YEAR(c_FirstDate),MONTH(c_FirstDate),1))', False, "dd.mm.yyyy"),
    ("c_SelIdx", "Индекс выбранного месяца",
     '=IF(ISNUMBER(sel_Month),0,IF(sel_Month="",0,IFERROR(MATCH(sel_Month,m_Labels,0),0)))', False, "0"),
    ("c_DefMonth", "Месяц по умолчанию",
     '=IF(c_LastMonth="",DATE(YEAR(c_Today),MONTH(c_Today),1),c_LastMonth)', False, "dd.mm.yyyy"),
    ("c_MonthStart", "Начало выбранного месяца",
     "=IF(ISNUMBER(sel_Month),DATE(YEAR(sel_Month),MONTH(sel_Month),1),"
     "IF(c_SelIdx>0,IF(ISNUMBER(INDEX(m_Starts,c_SelIdx)),INDEX(m_Starts,c_SelIdx),c_DefMonth),c_DefMonth))",
     False, "dd.mm.yyyy"),
    ("c_MonthEnd", "Конец выбранного месяца", "=EOMONTH(c_MonthStart,0)", False, "dd.mm.yyyy"),
    ("c_IsCurrent", "Выбран текущий месяц", "=AND(c_Today>=c_MonthStart,c_Today<=c_MonthEnd)", False, "General"),
    ("c_RefDate", "Дата отчета", "=IF(c_IsCurrent,c_Today,c_MonthEnd)", False, "dd.mm.yyyy"),
    ("c_LowDate", "Нижняя граница поиска значения",
     "=IF(c_IsCurrent,MIN(c_MonthStart,c_RefDate-1),c_MonthStart)", False, "dd.mm.yyyy"),
    ("c_MonthLabel", "Подпись месяца", "=" + _months_choose("c_MonthStart") + '&" "&YEAR(c_MonthStart)', False, "@"),
    ("c_ObjCount", "Объектов", f"=MIN({MAX_OBJ},COUNTA(db_Names))", False, "0"),
    ("c_DataDate", "Дата данных", f'=IF(MAX(D{OBJ_R0}:D{OBJ_R0 + MAX_OBJ - 1})=0,"",MAX(D{OBJ_R0}:D{OBJ_R0 + MAX_OBJ - 1}))', False, "dd.mm.yyyy"),
    ("c_CntValued", "Объектов со значением", f'=COUNTIF(I{OBJ_R0}:I{OBJ_R0 + MAX_OBJ - 1},">0")', False, "0"),
    ("c_CntGreen", "Зеленая зона", f"=COUNTIF(I{OBJ_R0}:I{OBJ_R0 + MAX_OBJ - 1},1)", False, "0"),
    ("c_CntYellow", "Желтая зона", f"=COUNTIF(I{OBJ_R0}:I{OBJ_R0 + MAX_OBJ - 1},2)", False, "0"),
    ("c_CntRed", "Красная зона",
     f"=COUNTIF(I{OBJ_R0}:I{OBJ_R0 + MAX_OBJ - 1},3)+COUNTIF(I{OBJ_R0}:I{OBJ_R0 + MAX_OBJ - 1},4)", False, "0"),
    ("c_CntOver", "Выше норматива", f"=COUNTIF(I{OBJ_R0}:I{OBJ_R0 + MAX_OBJ - 1},4)", False, "0"),
    ("c_CntNoData", "Нет данных", f"=COUNTIF(I{OBJ_R0}:I{OBJ_R0 + MAX_OBJ - 1},0)", False, "0"),
    ("c_MaxVal", "Максимум", f'=IF(c_CntValued=0,"",MAX(E{OBJ_R0}:E{OBJ_R0 + MAX_OBJ - 1}))', False, "0.0#"),
    ("c_MaxName", "Объект с максимумом", f'=IF(c_CntValued=0,"",C{SORT_R0})', False, "@"),
    ("c_AvgVal", "Среднее", f'=IF(c_CntValued=0,"",AVERAGE(E{OBJ_R0}:E{OBJ_R0 + MAX_OBJ - 1}))', False, "0.0#"),
    ("c_State", "Состояние (0-4)",
     "=IF(c_CntValued=0,0,IF(c_CntOver>0,4,IF(c_CntRed>0,3,IF(c_CntYellow>0,2,1))))", False, "0"),
    ("c_StateText", "Состояние, текст",
     '=CHOOSE(c_State+1,"НЕТ ДАННЫХ","НОРМА","ВНИМАНИЕ","РИСК","ПРЕВЫШЕНИЕ")', False, "@"),
    ("c_StateSub", "Состояние, пояснение",
     '=CHOOSE(c_State+1,"Загрузите файл с данными","Превышений не зафиксировано",'
     '"В желтой зоне: "&c_CntYellow&" СИКН","В красной зоне: "&c_CntRed&" СИКН",'
     '"Выше норматива: "&c_CntOver&" СИКН")', False, "@"),
    ("c_ZoneG", "Подпись зеленой зоны", '="0 – "&FIXED(cfg_Green,1)&" ppm"', False, "@"),
    ("c_ZoneY", "Подпись желтой зоны", '=FIXED(cfg_Green,1)&" – "&FIXED(cfg_Yellow,1)&" ppm"', False, "@"),
    ("c_ZoneR", "Подпись красной зоны", '="> "&FIXED(cfg_Yellow,1)&" ppm"', False, "@"),
    ("c_ZoneX", "Подпись превышения", '="> "&FIXED(cfg_Limit,1)&" ppm"', False, "@"),
    ("c_MonthCount", "Месяцев в списке", f"=COUNT(B{MONTH_R0}:B{MONTH_R0 + MONTH_N - 1})", False, "0"),
    ("c_LastMaxDate", "Последний день с максимумом",
     f"{{=MAX(IF(ISNUMBER(L{MAXDAY_R}:AP{MAXDAY_R}),L{DAYDATE_R}:AP{DAYDATE_R},0))}}", True, "dd.mm.yyyy"),
    ("c_Footer", "Строка состояния",
     '=IF(c_State=0,"НЕТ ДАННЫХ",IF(c_State=1,"ВСЕ В НОРМЕ","ТРЕБУЕТСЯ ВНИМАНИЕ"))', False, "@"),
    ("c_StaleCount", "Значений за более раннюю дату", f"=COUNTIF(K{OBJ_R0}:K{OBJ_R0 + MAX_OBJ - 1},1)", False, "0"),
]


def build_calc_sheet(wb, ws) -> None:
    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 16)
    ws.set_column(2, 2, 22)
    ws.set_column(3, 10, 12)
    ws.set_column(TR_C0, TR_C1, 9)
    bold = wb.add_format({"bold": True})
    date_f = wb.add_format({"num_format": "dd.mm.yyyy"})
    fmts: dict[str, object] = {}

    def f(num):
        if num not in fmts:
            fmts[num] = wb.add_format({"num_format": num})
        return fmts[num]

    ws.write(0, 0, "Служебный лист расчетов дашборда (формулы, не редактировать)", bold)
    for i, (name, label, formula, is_array, num) in enumerate(PARAMS):
        r = 1 + i  # строки 2..
        ws.write(r, 0, label)
        if is_array:
            ws.write_array_formula(r, 1, r, 1, formula, f(num))
        else:
            ws.write_formula(r, 1, formula, f(num))
        wb.define_name(name, f"={_q(S_CALC)}!$B${r + 1}")
    assert 1 + len(PARAMS) < OBJ_R0 - 3

    # Данные для кольцевой диаграммы и легенды (D2:F7)
    ws.write(0, 3, "Зона", bold)
    ws.write(0, 4, "Кол-во", bold)
    ws.write(0, 5, "Подпись", bold)
    donut = [("Зеленая", "=c_CntGreen", "c_LegG"), ("Желтая", "=c_CntYellow", "c_LegY"),
             ("Красная", "=c_CntRed", "c_LegR"), ("Нет данных", "=IF(c_ObjCount=0,1,c_CntNoData)", None)]
    for i, (lbl, formula, leg) in enumerate(donut):
        ws.write(1 + i, 3, lbl)
        ws.write_formula(1 + i, 4, formula)
        if leg:
            cnt = formula[1:]
            ws.write_formula(1 + i, 5, f'={cnt}&" ("&ROUND(100*{cnt}/MAX(1,c_CntValued),0)&"%)"')
            wb.define_name(leg, f"={_q(S_CALC)}!$F${2 + i}")
    ws.write(6, 3, "Всего")
    ws.write_formula(6, 4, "=c_ObjCount")
    wb.define_name("c_Total", f"={_q(S_CALC)}!$E$7")

    # Объекты в порядке базы
    hdr = ["№", "Объект", "Вид продукции", "Дата значения", "Значение", "Пред. дата",
           "Пред. значение", "Изменение", "Зона", "Ключ сортировки", "Ранняя дата"]
    for c, h in enumerate(hdr):
        ws.write(OBJ_R0 - 2, c, h, bold)
    for i in range(MAX_OBJ):
        r = OBJ_R0 + i  # 1-based
        ri = r - 1
        ws.write_number(ri, 0, i + 1)
        ws.write_formula(ri, 1, f'=IF($A{r}>c_ObjCount,"",INDEX(db_Names,1,$A{r})&"")')
        ws.write_formula(ri, 2, f'=IF($B{r}="","",INDEX(db_Groups,1,$A{r})&"")')
        ws.write_array_formula(ri, 3, ri, 3,
            f'{{=IF($B{r}="",0,MAX(IF((db_Dates>=c_LowDate)*(db_Dates<=c_RefDate)'
            f'*ISNUMBER(INDEX(db_Vals,0,$A{r})),db_Dates,0)))}}', date_f)
        ws.write_formula(ri, 4, f'=IF($D{r}=0,"",INDEX(db_Vals,MATCH($D{r},db_Dates,0),$A{r}))')
        ws.write_array_formula(ri, 5, ri, 5,
            f'{{=IF($E{r}="",0,MAX(IF((db_Dates<$D{r})*ISNUMBER(INDEX(db_Vals,0,$A{r})),db_Dates,0)))}}',
            date_f)
        ws.write_formula(ri, 6, f'=IF($F{r}=0,"",INDEX(db_Vals,MATCH($F{r},db_Dates,0),$A{r}))')
        ws.write_formula(ri, 7, f'=IF(OR($E{r}="",$G{r}=""),"",ROUND($E{r}-$G{r},4))')
        ws.write_formula(ri, 8, f'=IF($B{r}="",-1,IF($E{r}="",0,IF($E{r}<=cfg_Green,1,'
                                f'IF($E{r}<=cfg_Yellow,2,IF($E{r}<=cfg_Limit,3,4)))))')
        ws.write_formula(ri, 9, f'=IF($B{r}="",-2-$A{r}/1000000,IF($E{r}="",-1-$A{r}/1000000,'
                                f'$E{r}-$A{r}/10000000))')
        ws.write_formula(ri, 10, f'=IF(OR($E{r}="",c_DataDate=""),0,IF($D{r}<c_DataDate,1,0))')

    # Сортировка по убыванию значения + тренды по дням выбранного месяца
    o0, o1 = OBJ_R0, OBJ_R0 + MAX_OBJ - 1
    hdr2 = ["Место", "№ объекта", "Объект", "Значение", "Зона", "Изменение", "Ранняя дата",
            "Дата значения"]
    for c, h in enumerate(hdr2):
        ws.write(SORT_R0 - 3, c, h, bold)
    ws.write(DAYNUM_R - 1, TR_C0 - 1, "День", bold)
    ws.write(DAYDATE_R - 1, TR_C0 - 1, "Дата", bold)
    ddmm = wb.add_format({"num_format": "dd.mm"})
    for d in range(31):
        c = TR_C0 + d
        col = xl_col_to_name(c)
        ws.write_number(DAYNUM_R - 1, c, d + 1)
        ws.write_formula(DAYDATE_R - 1, c,
                         f'=IF({col}${DAYNUM_R}<=DAY(c_MonthEnd),c_MonthStart+{col}${DAYNUM_R}-1,"")',
                         ddmm)
    for k in range(MAX_OBJ):
        r = SORT_R0 + k
        ri = r - 1
        ws.write_number(ri, 0, k + 1)
        ws.write_formula(ri, 1, f"=MATCH(LARGE($J${o0}:$J${o1},$A{r}),$J${o0}:$J${o1},0)")
        ws.write_formula(ri, 2, f"=INDEX($B${o0}:$B${o1},$B{r})")
        ws.write_formula(ri, 3, f"=INDEX($E${o0}:$E${o1},$B{r})")
        ws.write_formula(ri, 4, f"=INDEX($I${o0}:$I${o1},$B{r})")
        ws.write_formula(ri, 5, f"=INDEX($H${o0}:$H${o1},$B{r})")
        ws.write_formula(ri, 6, f"=INDEX($K${o0}:$K${o1},$B{r})")
        ws.write_formula(ri, 7, f"=INDEX($D${o0}:$D${o1},$B{r})", date_f)
        for d in range(31):
            c = TR_C0 + d
            col = xl_col_to_name(c)
            day = f"{col}${DAYDATE_R}"
            look = f"INDEX(db_Vals,MATCH({day},db_Dates,0),$B{r})"
            ws.write_formula(ri, c, f'=IFERROR(IF(OR($C{r}="",{day}="",{day}>c_RefDate),NA(),'
                                    f'IF(ISNUMBER({look}),{look},NA())),NA())')

    # Динамика максимального значения
    labels = {MAXDAY_R: "Максимум за день", LIM_R: "Норматив", YEL_R: "Граница желтой зоны",
              GRN_R: "Граница зеленой зоны", LAST_R: "Последняя точка", CAT_R: "Подпись дня"}
    for r, lbl in labels.items():
        ws.write(r - 1, TR_C0 - 1, lbl, bold)
    for d in range(31):
        c = TR_C0 + d
        col = xl_col_to_name(c)
        day = f"{col}${DAYDATE_R}"
        row = f"INDEX(db_Vals,MATCH({day},db_Dates,0),0)"
        ws.write_formula(MAXDAY_R - 1, c, f'=IFERROR(IF(OR({day}="",{day}>c_RefDate),NA(),'
                                          f'IF(COUNT({row})=0,NA(),MAX({row}))),NA())')
        ws.write_formula(LIM_R - 1, c, f'=IF({day}="",NA(),cfg_Limit)')
        ws.write_formula(YEL_R - 1, c, f'=IF({day}="",NA(),cfg_Yellow)')
        ws.write_formula(GRN_R - 1, c, f'=IF({day}="",NA(),cfg_Green)')
        ws.write_formula(LAST_R - 1, c, f"=IF(ISNUMBER({col}{MAXDAY_R}),"
                                        f"IF({day}=c_LastMaxDate,{col}{MAXDAY_R},NA()),NA())")
        # подписи оси — каждые 7 дней (1, 8, 15, 22, 29), остальные пустые
        ws.write_formula(CAT_R - 1, c, f'=IF(OR({day}="",MOD({col}${DAYNUM_R}-1,7)<>0),"",'
                                       f'RIGHT("0"&DAY({day}),2)&"."&RIGHT("0"&MONTH({day}),2))')

    # Список месяцев (последний — первым)
    ws.write(MONTH_R0 - 2, 0, "Месяцы с данными", bold)
    for k in range(MONTH_N):
        r = MONTH_R0 + k
        ri = r - 1
        ws.write_number(ri, 0, k + 1)
        ws.write_formula(ri, 1, f'=IF(c_LastMonth="","",IF(EDATE(c_LastMonth,1-$A{r})<c_FirstMonth,"",'
                                f'EDATE(c_LastMonth,1-$A{r})))', date_f)
        ws.write_formula(ri, 2, f'=IF($B{r}="","",{_months_choose(f"$B{r}")}&" "&YEAR($B{r}))')
    m0, m1 = MONTH_R0, MONTH_R0 + MONTH_N - 1
    wb.define_name("m_Starts", f"={_q(S_CALC)}!$B${m0}:$B${m1}")
    wb.define_name("m_Labels", f"={_q(S_CALC)}!$C${m0}:$C${m1}")
    wb.define_name("lst_Months", f"=OFFSET({_q(S_CALC)}!$C${m0},0,0,MAX(1,c_MonthCount),1)")


# ----------------------------------------------------------------------------
# Лист «БД»
# ----------------------------------------------------------------------------
def build_db_sheet(wb, ws) -> None:
    last_col = xl_col_to_name(MAX_OBJ)  # AO
    last_row = 2 + DB_ROWS
    wb.define_name("db_N", f"=MAX(1,COUNT({_q(S_DB)}!$A$3:$A${last_row}))")
    wb.define_name("db_Names", f"={_q(S_DB)}!$B$1:${last_col}$1")
    wb.define_name("db_Groups", f"={_q(S_DB)}!$B$2:${last_col}$2")
    wb.define_name("db_Dates", f"={_q(S_DB)}!$A$3:INDEX({_q(S_DB)}!$A$3:$A${last_row},db_N)")
    wb.define_name("db_Vals", f"={_q(S_DB)}!$B$3:INDEX({_q(S_DB)}!$B$3:${last_col}${last_row},"
                              f"db_N,{MAX_OBJ})")
    hdr = wb.add_format({"bold": True, "font_name": FONT, "font_size": 9, "bg_color": C["panel_hdr"],
                         "font_color": C["text"], "border": 1, "border_color": C["border"],
                         "align": "center", "valign": "vcenter", "text_wrap": True})
    sub = wb.add_format({"font_name": FONT, "font_size": 8, "bg_color": C["panel"],
                         "font_color": C["muted"], "border": 1, "border_color": C["border"],
                         "align": "center", "valign": "vcenter", "text_wrap": True})
    date_f = wb.add_format({"num_format": "dd.mm.yyyy", "font_name": FONT, "font_size": 9,
                            "align": "center"})
    val_f = wb.add_format({"font_name": FONT, "font_size": 9, "align": "center"})
    ws.set_column(0, 0, 12, date_f)
    ws.set_column(1, 200, 13, val_f)
    ws.set_row(0, 30, hdr)
    ws.set_row(1, 24, sub)
    ws.write(0, 0, "Дата \\ Объект", hdr)
    ws.write(1, 0, "Вид продукции", sub)
    ws.freeze_panes(2, 1)
    ws.write_comment(0, 0, "База данных дашборда. Заполняется кнопкой «Импорт данных» "
                           "(строка 1 — объекты, строка 2 — вид продукции, далее — по строке на день). "
                           "Ручное редактирование не требуется.",
                     {"x_scale": 2, "y_scale": 1.2})


# ----------------------------------------------------------------------------
# Лист «Сводка»
# ----------------------------------------------------------------------------
def build_dash_sheet(wb, ws, assets: Assets) -> dict:
    cv = Canvas(wb, ws, W, H)
    draw_header(cv, assets)
    cv.put((1080, 48, 1352, 70), None, name="ui_Updated")  # формула ниже (NOW)
    cv.put((1352, 48, 1372, 70), "●", size=9, align="center", color=C["dim"])
    draw_nav(cv, assets, active=0)

    # --- панель управления: период, дата данных, загрузка, импорт
    lbl = dict(bold=True, size=8.5, color=C["muted"])
    cv.put((126, 92, 194, 126), "ПЕРИОД:", **lbl)
    cv.fill((194, 92, 398, 126), C["input"])
    cv.frame((194, 92, 398, 126), C["divider"])
    cv.put((194, 92, 372, 126), "", bold=True, size=11, color=C["text"], indent=1, num="@",
           name="sel_Month",
           validation={"validate": "list", "source": "=lst_Months",
                       "input_title": "Период отображения",
                       "input_message": "Выберите месяц из списка",
                       "error_title": "Период", "error_message": "Выберите месяц из списка.",
                       "error_type": "stop"})
    cv.put((372, 92, 398, 126), "▼", size=8, color=C["muted"], align="center")
    cv.put((422, 92, 512, 126), "ДАННЫЕ НА:", **lbl)
    cv.put((512, 92, 612, 126), '=IF(c_DataDate="","—",c_DataDate)', bold=True, size=11,
           num="dd.mm.yyyy", align="left")
    cv.put((636, 92, 732, 126), "ЗАГРУЖЕНО:", **lbl)
    cv.put((732, 92, 888, 126), '=IF(N(sys_LastImport)=0,"—",sys_LastImport)', size=10,
           num="dd.mm.yyyy hh:mm", align="left")
    cv.put((906, 92, 952, 126), "ФАЙЛ:", **lbl)
    cv.put((952, 92, 1186, 126), '=IF(sys_LastFile="","—",sys_LastFile)', size=9,
           color=C["text2"], shrink=True)
    button(cv, 1196, 92, assets.button("import"), "ImportData")

    # --- KPI: общее состояние
    card1 = (126, 138, 400, 238)
    cv.fill(card1, C["panel"])
    cv.frame(card1, C["border"])
    cv.put((140, 146, 392, 166), "ОБЩЕЕ СОСТОЯНИЕ", bold=True, size=8.5, color=C["text2"])
    # значок-щит: пять картинок, видимую выбирает макрос UpdateStateIcon по c_State
    for st in range(5):
        cv.image(142, 168, assets.icon(f"state_{st}"), scale=0.5, descr=f"state:{st}")
    cv.put((212, 170, 392, 200), "=c_StateText", bold=True, size=16, color=C["green"],
           align="left")
    cv.put((212, 200, 392, 222), "=c_StateSub", size=8.5, color=C["text2"], align="left")

    # --- KPI: значения и количество
    card2 = (412, 138, 1380, 238)
    cv.fill(card2, C["panel"])
    cv.frame(card2, C["border"])
    klabel = dict(bold=True, size=8.5, color=C["text2"])
    big = dict(bold=True, size=24, valign="bottom")
    unit = dict(size=11, color=C["text2"], valign="bottom")
    cv.put((426, 146, 622, 166), "МАКСИМАЛЬНОЕ ЗНАЧЕНИЕ", **klabel)
    cv.put((426, 168, 512, 206), '=IF(c_MaxVal="","—",c_MaxVal)', align="right", num="0.0#", **big)
    cv.put((512, 168, 600, 206), "ppm", indent=1, **unit)
    cv.put((426, 208, 622, 230), "=c_MaxName", size=9, color=C["muted"])
    cv.vline(628, 150, 228, C["divider"])
    cv.put((642, 146, 806, 166), "СРЕДНЕЕ ЗНАЧЕНИЕ", **klabel)
    cv.put((642, 168, 714, 206), '=IF(c_AvgVal="","—",c_AvgVal)', align="right", num="0.0#", **big)
    cv.put((714, 168, 800, 206), "ppm", indent=1, **unit)
    cv.put((642, 208, 806, 230), '=IF(c_CntValued=0,"","по "&c_CntValued&" СИКН")', size=9,
           color=C["muted"])
    cv.vline(812, 150, 228, C["divider"])
    cv.put((826, 146, 1166, 166), "КОЛИЧЕСТВО СИКН", **klabel)
    zones = [(826, 936, "ЗЕЛЕНАЯ ЗОНА", "=c_CntGreen", C["green"]),
             (944, 1050, "ЖЕЛТАЯ ЗОНА", "=c_CntYellow", C["yellow"]),
             (1058, 1166, "КРАСНАЯ ЗОНА", "=c_CntRed", C["red"])]
    for i, (x0, x1, t, frm, col) in enumerate(zones):
        cv.put((x0, 170, x1, 188), t, bold=True, size=8, color=col)
        cv.put((x0, 190, x1, 230), frm, bold=True, size=24, color=col, num="0")
        if i:
            cv.vline(x0 - 6, 172, 228, C["line"])
    cv.vline(1172, 150, 228, C["divider"])
    cv.put((1186, 146, 1370, 166), "ВСЕГО СИКН", **klabel)
    cv.put((1186, 168, 1370, 206), "=c_ObjCount", num="0", align="left", **big)
    cv.put((1186, 208, 1370, 230), '=IF(c_CntNoData>0,"нет данных: "&c_CntNoData,"")', size=9,
           color=C["muted"])

    # --- таблица
    tp = (126, 250, 1036, 948)
    cv.fill(tp, C["panel"])
    cv.frame(tp, C["border"])
    cv.put((140, 250, 740, 284), '="СОДЕРЖАНИЕ ЛХОС ПО СИКН — "&UPPER(c_MonthLabel)', bold=True,
           size=10, color=C["text2"])
    cv.put((740, 250, 1022, 284),
           '=IF(c_ObjCount=0,"База пуста — нажмите «ИМПОРТ ДАННЫХ»",'
           'IF(c_StaleCount>0,"* — значение за предыдущую дату",""))',
           size=8, color=C["muted"], align="right")
    cols = [  # (x0, x1, заголовок, выравнивание)
        (140, 180, "№", "center"),
        (180, 360, "СИКН", "left"),
        (360, 470, "ЗНАЧЕНИЕ, ppm", "center"),
        (470, 610, "СТАТУС", "left"),
        (610, 740, "ЗОНА", "center"),
        (740, 880, "ИЗМЕНЕНИЕ (24Ч)", "center"),
        (880, 1022, "ТРЕНД", "center"),
    ]
    hy0, hy1 = 284, 314
    cv.fill((140, hy0, 1022, hy1), C["panel_hdr"])
    for x0, x1, t, al in cols:
        cv.put((x0, hy0, x1, hy1), t, bold=True, size=8, color=C["text2"], align=al,
               indent=1 if al == "left" else 0)
    cv.hline(140, 1022, hy1, C["divider"])
    row_h = 26
    rows = []
    for k in range(TABLE_ROWS):
        y0 = hy1 + k * row_h
        y1 = y0 + row_h
        rows.append((y0, y1))
        r = SORT_R0 + k
        nm = _calc(f"$C${r}")
        val = _calc(f"$D${r}")
        zone = _calc(f"$E${r}")
        chg = _calc(f"$F${r}")
        cv.put((140, y0, 180, y1), f'=IF({nm}="","",{k + 1})', size=9, color=C["text2"],
               align="center")
        cv.put((180, y0, 360, y1), f"={nm}", size=9.5, color=C["text"], indent=1)
        cv.put((360, y0, 470, y1), f'=IF({nm}="","",IF({val}="","—",{val}))', bold=True,
               size=10, color=C["green"], align="center", num="0.0#")
        cv.put((470, y0, 610, y1),
               f'=IF({nm}="","",CHOOSE({zone}+1,"●  НЕТ ДАННЫХ","●  НОРМА","●  ВНИМАНИЕ",'
               f'"●  РИСК","●  ПРЕВЫШЕНИЕ"))', bold=True, size=8.5, color=C["green"], indent=1)
        cv.put((610, y0, 740, y1),
               f'=IF({nm}="","",CHOOSE({zone}+1,"—",c_ZoneG,c_ZoneY,c_ZoneR,c_ZoneX))',
               size=9, color=C["green"], align="center")
        cv.put((740, y0, 834, y1), f'=IF(OR({nm}="",{chg}=""),"",{chg})', size=9.5,
               color=C["text"], align="right", num="+0.0#;-0.0#;0.0")
        cv.put((834, y0, 872, y1),
               f'=IF(OR({nm}="",{chg}=""),"",IF({chg}>0.00001,"↑",IF({chg}<-0.00001,"↓","—")))',
               bold=True, size=10, color=C["green"], align="center")
        cv.sparkline((880, y0, 1022, y1), {
            "range": f"{_q(S_CALC)}!{xl_col_to_name(TR_C0)}{r}:{xl_col_to_name(TR_C1)}{r}",
            "series_color": C["green"], "weight": 1.0, "empty_cells": "gaps",
        })
        cv.hline(140, 1022, y1, C["line"])
    for x in (180, 360, 470, 610, 740, 880):
        cv.vline(x, hy0, rows[-1][1], C["line"])

    # --- норматив и зоны
    p1 = (1050, 250, 1380, 560)
    cv.fill(p1, C["panel"])
    cv.frame(p1, C["border"])
    cv.put((1064, 250, 1370, 280), "НОРМАТИВ И ЗОНЫ", bold=True, size=10, color=C["text2"])
    cv.put((1064, 282, 1226, 312), "ПРЕДЕЛЬНОЕ ЗНАЧЕНИЕ", bold=True, size=8, color=C["muted"])
    cv.put((1226, 282, 1366, 312), "=cfg_Limit", bold=True, size=13, align="right",
           num='"≤ "0.0" ppm"')
    blocks = [(322, 392, C["red_block"]), (392, 462, C["yellow_block"]), (462, 532, C["green_block"])]
    for y0, y1, col in blocks:
        cv.fill((1124, y0, 1168, y1), col)
    scale = [(314, "=cfg_Limit"), (384, "=cfg_Yellow"), (454, "=cfg_Green"), (524, 0)]
    for y, v in scale:
        cv.put((1064, y, 1116, y + 16), v, size=8.5, color=C["muted"], align="right",
               num="0.0" if isinstance(v, str) else "0")
    ztxt = [
        (326, "=c_ZoneR", "КРАСНАЯ ЗОНА", "Повышенный риск", C["red"]),
        (398, "=c_ZoneY", "ЖЕЛТАЯ ЗОНА", "Внимание", C["yellow"]),
        (468, "=c_ZoneG", "ЗЕЛЕНАЯ ЗОНА", "Норма", C["green"]),
    ]
    for y, frm, t1, t2, col in ztxt:
        cv.put((1180, y, 1366, y + 20), frm, bold=True, size=11, color=col)
        cv.put((1180, y + 20, 1366, y + 38), t1, bold=True, size=8, color=col)
        cv.put((1180, y + 38, 1366, y + 56), t2, size=8.5, color=C["text"])
    cv.hline(1168, 1366, 392, C["line"])
    cv.hline(1168, 1366, 462, C["line"])
    cv.put((1064, 541, 1366, 558), "Все значения указаны в ppm", size=7.5, color=C["muted"],
           align="center")

    # --- распределение по зонам
    p2 = (1050, 572, 1380, 752)
    cv.fill(p2, C["panel"])
    cv.frame(p2, C["border"])
    cv.put((1064, 572, 1370, 602), "РАСПРЕДЕЛЕНИЕ ПО ЗОНАМ", bold=True, size=10, color=C["text2"])
    donut = wb.add_chart({"type": "doughnut"})
    donut.add_series({
        "name": "Зоны",
        "categories": f"={_q(S_CALC)}!$D$2:$D$5",
        "values": f"={_q(S_CALC)}!$E$2:$E$5",
        "points": [
            {"fill": {"color": C["green_block"]}, "border": {"color": C["panel"], "width": 1.5}},
            {"fill": {"color": C["yellow_block"]}, "border": {"color": C["panel"], "width": 1.5}},
            {"fill": {"color": C["red_block"]}, "border": {"color": C["panel"], "width": 1.5}},
            {"fill": {"color": C["nodata"]}, "border": {"color": C["panel"], "width": 1.5}},
        ],
    })
    donut.set_hole_size(72)
    donut.set_legend({"none": True})
    donut.set_title({"none": True})
    donut.set_chartarea({"fill": {"none": True}, "border": {"none": True}})
    donut.set_plotarea({"fill": {"none": True}, "border": {"none": True},
                        "layout": {"x": 0.08, "y": 0.06, "width": 0.84, "height": 0.84}})
    cv.chart((1060, 604, 1196, 748), donut)
    # число в центре кольца — в ячейках под прозрачной диаграммой
    cv.put((1090, 652, 1166, 682), "=c_ObjCount", bold=True, size=20, align="center", num="0")
    cv.put((1090, 682, 1166, 698), "ВСЕГО", bold=True, size=7, color=C["muted"], align="center",
           valign="top")
    legend = [(622, "Зеленая зона", "=c_LegG", C["green_block"]),
              (652, "Желтая зона", "=c_LegY", C["yellow_block"]),
              (682, "Красная зона", "=c_LegR", C["red_block"])]
    for y, t, frm, col in legend:
        cv.put((1202, y, 1216, y + 24), "■", size=11, color=col, align="center")
        cv.put((1216, y, 1310, y + 24), t, size=9, color=C["text"], indent=1)
        cv.put((1310, y, 1368, y + 24), frm, size=9, color=C["text"], align="right")

    # --- динамика максимального значения
    p3 = (1050, 764, 1380, 948)
    cv.fill(p3, C["panel"])
    cv.frame(p3, C["border"])
    cv.put((1064, 764, 1370, 794), "ДИНАМИКА МАКСИМАЛЬНОГО ЗНАЧЕНИЯ", bold=True, size=10,
           color=C["text2"])
    cv.chart((1054, 794, 1376, 944), max_chart(wb))

    # --- подвал
    cv.fill((0, 958, W, H), C["footer"])
    cv.hline(0, W, 958, C["divider"])
    for st in range(5):  # значок состояния в подвале, переключается вместе со щитом
        cv.image(26, 967, assets.icon(f"foot_{st}"), scale=0.5, descr=f"foot:{st}")
    cv.put((58, 962, 470, 996), '="СОСТОЯНИЕ: "&c_Footer', bold=True, size=9.5, color=C["green"])
    cv.image(540, 972, assets.icon("ui_doc"), scale=20 / 48)
    cv.put((566, 962, 1010, 996),
           '=IF(sys_LastFile="","Данные не загружены — нажмите «Импорт данных»",'
           '"Источник данных: "&sys_LastFile)', size=9, color=C["text2"])
    cv.image(1040, 972, assets.icon("ui_shield"), scale=20 / 48)
    cv.put((1066, 962, 1122, 996), "Доступ:", size=9, color=C["muted"])
    cv.put((1122, 962, 1248, 996), "=cfg_Role", bold=True, size=9, color=C["text"])
    cv.vline(1256, 964, 994, C["divider"])
    button(cv, 1268, 965, assets.button("exit"), "ExitApp")

    cv.build()

    # --- формулы, которые удобнее писать после построения сетки
    upd_r, upd_c, _, _ = cv.cell_range((1080, 48, 1352, 70))
    helper_c = cv.n_cols + 1  # скрытые служебные столбцы справа от макета
    elapsed = xl_rowcol_to_cell(upd_r, helper_c, row_abs=True, col_abs=True)
    ws.write_formula(upd_r, helper_c, '=IF(N(sys_LastImport)=0,-1,NOW()-sys_LastImport)')
    wb.define_name("ui_Elapsed", f"={_q(S_DASH)}!{elapsed}")
    upd_fmt = cv.fmt({"bg": C["header"], "size": 9, "color": C["muted"], "align": "right"})
    ws.write_formula(
        upd_r, upd_c,
        '=IF(N(sys_LastImport)=0,"Данные еще не загружались",'
        '"Обновление: "&IF(NOW()-sys_LastImport<1/1440,"только что",'
        'IF(NOW()-sys_LastImport<1/24,INT((NOW()-sys_LastImport)*1440)&" мин назад",'
        'IF(NOW()-sys_LastImport<1,INT((NOW()-sys_LastImport)*24)&" ч "&'
        'MOD(INT((NOW()-sys_LastImport)*1440),60)&" мин назад",'
        'INT(NOW()-sys_LastImport)&" дн "&MOD(INT((NOW()-sys_LastImport)*24),24)&" ч назад"))))',
        upd_fmt)

    # служебные ячейки: зона/«ранняя дата» строк таблицы, состояние
    state_r = cv.cell_range(card1)[0]
    ws.write_formula(state_r, helper_c, "=c_State")
    state_ref = xl_rowcol_to_cell(state_r, helper_c, row_abs=True, col_abs=True)
    zone_c, stale_c = helper_c, helper_c + 1
    table_cells = []
    for k, (y0, y1) in enumerate(rows):
        r = SORT_R0 + k
        rr = cv.cell_range((140, y0, 180, y1))[0]
        if k == 0 and rr == state_r:
            raise AssertionError("helper collision")
        ws.write_formula(rr, zone_c, f"={_calc(f'$E${r}')}")
        ws.write_formula(rr, stale_c, f"={_calc(f'$G${r}')}")
        table_cells.append(rr)
    for c in range(helper_c - 1, helper_c + 2):
        ws.set_column(c, c, 8, None, {"hidden": True})
    ws.set_column(helper_c + 2, 16383, 9, cv.fmt({"bg": C["page"]}))

    # --- условное форматирование
    first, last = table_cells[0], table_cells[-1]

    def col_of(x0, y0=rows[0][0], x1=None, y1=rows[0][1]):
        return cv.cell_range((x0, y0, x1 or x0 + 1, y1))[1]

    def rng(c):
        return f"{xl_rowcol_to_cell(first, c)}:{xl_rowcol_to_cell(last, c)}"

    zone_abs = xl_col_to_name(zone_c)
    stale_abs = xl_col_to_name(stale_c)
    zr = f"${zone_abs}{first + 1}"
    sr = f"${stale_abs}{first + 1}"
    value_c = cv.cell_range((360, rows[0][0], 470, rows[0][1]))[1]
    status_c = cv.cell_range((470, rows[0][0], 610, rows[0][1]))[1]
    zonecol_c = cv.cell_range((610, rows[0][0], 740, rows[0][1]))[1]
    arrow_c = cv.cell_range((834, rows[0][0], 872, rows[0][1]))[1]
    fmt_cache = {}

    def cf_fmt(color, num=None):
        key = (color, num)
        if key not in fmt_cache:
            p = {"font_color": color}
            if num:
                p["num_format"] = num
            fmt_cache[key] = wb.add_format(p)
        return fmt_cache[key]

    for z in (1, 2, 3, 4):
        for s, num in ((0, "0.0#"), (1, '0.0#"*"')):
            ws.conditional_format(rng(value_c), {
                "type": "formula", "criteria": f"=AND({zr}={z},{sr}={s})",
                "format": cf_fmt(ZONE_COLOR[z], num)})
    for c in (value_c, status_c, zonecol_c, arrow_c):
        ws.conditional_format(rng(c), {"type": "formula", "criteria": f"={zr}=0",
                                       "format": cf_fmt(C["muted"])})
    for c in (status_c, zonecol_c, arrow_c):
        for z in (2, 3, 4):
            ws.conditional_format(rng(c), {"type": "formula", "criteria": f"={zr}={z}",
                                           "format": cf_fmt(ZONE_COLOR[z])})

    # состояние: цвет текста
    st_r, st_c, _, _ = cv.cell_range((212, 170, 392, 200))
    ft_r, ft_c, _, _ = cv.cell_range((58, 962, 470, 996))
    for (r, c) in ((st_r, st_c), (ft_r, ft_c)):
        for val, color in ((0, C["muted"]), (2, C["yellow"]), (3, C["red"]), (4, C["red"])):
            ws.conditional_format(r, c, r, c, {"type": "formula", "criteria": f"={state_ref}={val}",
                                               "format": cf_fmt(color)})
    # индикатор свежести данных
    dot_r, dot_c, _, _ = cv.cell_range((1352, 48, 1372, 70))
    el = elapsed
    for crit, color in ((f"=AND({el}>=0,{el}<1)", C["green"]), (f"=AND({el}>=1,{el}<3)", C["yellow"]),
                        (f"={el}>=3", C["red"])):
        ws.conditional_format(dot_r, dot_c, dot_r, dot_c, {"type": "formula", "criteria": crit,
                                                           "format": cf_fmt(color)})

    ws.hide_gridlines(2)
    ws.hide_row_col_headers()
    ws.set_zoom(85)
    print_setup(ws, cv)
    sel_r, sel_c, _, _ = cv.cell_range((194, 92, 372, 126))
    ws.set_selection(sel_r, sel_c, sel_r, sel_c)
    ws.set_tab_color(C["green"])
    return {"sel_cell": (sel_r, sel_c), "table_rows": table_cells}


def print_setup(ws, cv: Canvas) -> None:
    """Печать макета на один лист A4 (альбомная ориентация)."""
    ws.set_landscape()
    ws.set_paper(9)
    ws.set_margins(left=0.2, right=0.2, top=0.2, bottom=0.2)
    ws.print_area(0, 0, cv.n_rows - 1, cv.n_cols - 1)
    ws.fit_to_pages(1, 1)
    ws.center_horizontally()


def max_chart(wb):
    cats = f"={_q(S_CALC)}!${xl_col_to_name(TR_C0)}${CAT_R}:${xl_col_to_name(TR_C1)}${CAT_R}"

    def vals(r):
        return f"={_q(S_CALC)}!${xl_col_to_name(TR_C0)}${r}:${xl_col_to_name(TR_C1)}${r}"

    ch = wb.add_chart({"type": "line"})
    for r, name, color, dash in ((LIM_R, "Норматив", C["red"], "dash"),
                                 (YEL_R, "Желтая зона", C["yellow"], "dash"),
                                 (GRN_R, "Зеленая зона", "#3E6B52", "round_dot")):
        ch.add_series({"name": name, "categories": cats, "values": vals(r),
                       "line": {"color": color, "width": 1.0, "dash_type": dash},
                       "marker": {"type": "none"}})
    ch.add_series({"name": "Максимум", "categories": cats, "values": vals(MAXDAY_R),
                   "line": {"color": C["green"], "width": 1.75},
                   "marker": {"type": "circle", "size": 4,
                              "fill": {"color": C["green"]}, "border": {"color": C["green"]}}})
    ch.add_series({"name": "Последнее", "categories": cats, "values": vals(LAST_R),
                   "line": {"none": True},
                   "marker": {"type": "circle", "size": 6,
                              "fill": {"color": C["green"]}, "border": {"color": "#FFFFFF"}},
                   "data_labels": {"value": True, "position": "above", "num_format": "0.0#",
                                   "font": {"name": FONT, "size": 9, "bold": True,
                                            "color": C["green"]}}})
    axis_font = {"name": FONT, "size": 7.5, "color": C["muted"], "rotation": 0}
    ch.set_x_axis({"num_font": axis_font, "text_axis": True,
                   "line": {"color": C["divider"]}, "major_tick_mark": "none",
                   "interval_unit": 7, "label_position": "low"})
    ch.set_y_axis({"min": 0, "num_format": "0.0", "num_font": axis_font,
                   "major_gridlines": {"visible": False}, "line": {"none": True},
                   "major_tick_mark": "none", "name": "ppm",
                   "name_font": {"name": FONT, "size": 7.5, "color": C["muted"], "bold": False,
                                 "rotation": 0},
                   "name_layout": {"x": 0.02, "y": 0.02}})
    ch.set_legend({"none": True})
    ch.set_title({"none": True})
    ch.set_chartarea({"fill": {"none": True}, "border": {"none": True}})
    ch.set_plotarea({"fill": {"none": True}, "border": {"none": True}})
    ch.show_blanks_as("gap")
    return ch


# ----------------------------------------------------------------------------
# Заглушки разделов
# ----------------------------------------------------------------------------
def build_placeholder(wb, ws, assets: Assets, active: int, title: str, text: str) -> None:
    cv = Canvas(wb, ws, W, 960)
    draw_header(cv, assets)
    draw_nav(cv, assets, active)
    p = (126, 92, 1380, 948)
    cv.fill(p, C["panel"])
    cv.frame(p, C["border"])
    cv.put((150, 110, 1360, 150), title, bold=True, size=14, color=C["text"])
    cv.hline(150, 1360, 150, C["divider"])
    cv.image(722, 380, assets.icon(NAV[active][2] + "_normal"), scale=1.5)
    cv.put((300, 490, 1206, 530), "Раздел в разработке", bold=True, size=16, color=C["text"],
           align="center")
    cv.put((300, 530, 1206, 600), text, size=10, color=C["muted"], align="center", wrap=True,
           valign="top")
    cv.put((560, 620, 746, 652), "← К СВОДКЕ", bold=True, size=9.5, color=C["green"],
           align="center", url=f"internal:{_q(S_DASH)}!A1", url_tip="Перейти на сводку")
    cv.build()
    ws.hide_gridlines(2)
    ws.hide_row_col_headers()
    ws.set_zoom(85)
    print_setup(ws, cv)
    ws.set_column(cv.n_cols, 16383, 9, cv.fmt({"bg": C["page"]}))


# ----------------------------------------------------------------------------
# Настройки
# ----------------------------------------------------------------------------
def build_settings(wb, ws, assets: Assets) -> None:
    h = 548 + 34 + 30 + LOG_ROWS * 22 + 24
    cv = Canvas(wb, ws, W, h)
    draw_header(cv, assets)
    draw_nav(cv, assets, active=5)

    lab = dict(size=9.5, color=C["text2"])
    inp = dict(size=10.5, bold=True, color=C["text"], bg=C["input"], align="center", locked=False)

    p = (126, 92, 700, 340)
    cv.fill(p, C["panel"])
    cv.frame(p, C["border"])
    cv.put((140, 96, 690, 128), "ПОРОГОВЫЕ ЗНАЧЕНИЯ, ppm", bold=True, size=10, color=C["text2"])
    rows = [("Зеленая зона — до (включительно)", "cfg_Green", 1.5, C["green"]),
            ("Желтая зона — до (включительно)", "cfg_Yellow", 2.3, C["yellow"]),
            ("Норматив (предельное значение)", "cfg_Limit", 3.0, C["red"])]
    for i, (t, name, v, col) in enumerate(rows):
        y = 136 + i * 40
        cv.put((140, y, 150, y + 30), "■", size=10, color=col)
        cv.put((152, y, 480, y + 30), t, **lab)
        cv.frame((490, y + 2, 580, y + 28), C["divider"])
        cv.put((490, y + 2, 580, y + 28), v, num="0.0#", name=name,
               validation={"validate": "decimal", "criteria": "between", "minimum": 0,
                           "maximum": 1000, "error_message": "Введите число от 0 до 1000."},
               **inp)
        cv.put((586, y, 640, y + 30), "ppm", size=9, color=C["muted"])
    cv.put((140, 262, 690, 334),
           "Красная зона — значения выше границы желтой зоны. Значения выше норматива "
           "получают статус «ПРЕВЫШЕНИЕ». Изменения сразу применяются на сводке.",
           size=8.5, color=C["muted"], wrap=True, valign="top")

    p = (712, 92, 1380, 340)
    cv.fill(p, C["panel"])
    cv.frame(p, C["border"])
    cv.put((726, 96, 1370, 128), "ОФОРМЛЕНИЕ", bold=True, size=10, color=C["text2"])
    texts = [("Название организации", "cfg_Company", "TATNEFT"),
             ("Заголовок", "cfg_Title", "СОДЕРЖАНИЕ ЛЕГКОЛЕТУЧИХ ХЛОРОРГАНИЧЕСКИХ СОЕДИНЕНИЙ ПО СИКН"),
             ("Подзаголовок", "cfg_Subtitle", "ОПЕРАТИВНАЯ СВОДКА"),
             ("Доступ (роль)", "cfg_Role", "ДИСПЕТЧЕР")]
    for i, (t, name, v) in enumerate(texts):
        y = 136 + i * 40
        cv.put((726, y, 900, y + 30), t, **lab)
        cv.frame((900, y + 2, 1366, y + 28), C["divider"])
        cv.put((900, y + 2, 1366, y + 28), v, name=name, indent=1,
               **{**inp, "align": "left", "size": 9.5})

    p = (126, 352, 1380, 536)
    cv.fill(p, C["panel"])
    cv.frame(p, C["border"])
    cv.put((140, 356, 1000, 388), "ДАННЫЕ", bold=True, size=10, color=C["text2"])
    info = [
        ("Последняя загрузка", None, "sys_LastImport", "dd.mm.yyyy hh:mm"),
        ("Файл", None, "sys_LastFile", "@"),
        ("Режим загрузки", None, "sys_LastMode", "@"),
        ("Объектов в базе", "=COUNTA(db_Names)", None, "0"),
        ("Значений в базе", "=COUNT(db_Vals)", None, "#,##0"),
        ("Период данных", '=IF(c_FirstDate="","—",DAY(c_FirstDate)&"."&RIGHT("0"&MONTH(c_FirstDate),2)'
                          '&"."&YEAR(c_FirstDate)&" – "&DAY(c_LastDate)&"."&RIGHT("0"&MONTH(c_LastDate),2)'
                          '&"."&YEAR(c_LastDate))', None, "@"),
    ]
    for i, (t, frm, name, num) in enumerate(info):
        col, row = divmod(i, 3)
        x = 140 + col * 440
        y = 396 + row * 30
        cv.put((x, y, x + 170, y + 28), t, size=9, color=C["muted"])
        extra = {"name": name} if name else {}
        cv.put((x + 170, y, x + 430, y + 28), frm, num=num, bold=True, size=9.5,
               color=C["text"], **extra)
    button(cv, 1182, 398, assets.button("import"), "ImportData")
    button(cv, 1182, 444, assets.button("clear"), "ClearDatabase")
    cv.put((140, 490, 1366, 530),
           "Первый импорт заполняет базу. При следующих импортах программа спросит: загрузить "
           "все данные заново (старые удаляются) или добавить только новые значения. "
           "Для работы кнопок разрешите макросы.", size=8.5, color=C["muted"], wrap=True,
           valign="top")

    y0 = 548
    p = (126, y0, 1380, h - 12)
    cv.fill(p, C["panel"])
    cv.frame(p, C["border"])
    cv.put((140, y0 + 2, 1000, y0 + 32), "ЖУРНАЛ ЗАГРУЗОК", bold=True, size=10, color=C["text2"])
    lcols = [(140, 270, "Дата и время", "dd.mm.yyyy hh:mm", "center"),
             (270, 560, "Файл", "@", "left"),
             (560, 700, "Режим", "@", "left"),
             (700, 790, "В файле", "#,##0", "center"),
             (790, 880, "Добавлено", "#,##0", "center"),
             (880, 970, "Совпало", "#,##0", "center"),
             (970, 1070, "Отличаются", "#,##0", "center"),
             (1070, 1150, "Объектов", "0", "center"),
             (1150, 1366, "Период в базе", "@", "center")]
    hy = y0 + 34
    cv.fill((140, hy, 1366, hy + 30), C["panel_hdr"])
    for j, (x0, x1, t, num, al) in enumerate(lcols):
        cv.put((x0, hy, x1, hy + 30), t, bold=True, size=8, color=C["text2"], align="center",
               name=f"log_C{j + 1}")
        for i in range(LOG_ROWS):
            y = hy + 30 + i * 22
            cv.put((x0, y, x1, y + 22), None, size=8.5, color=C["text"], num=num, align=al,
                   indent=1 if al == "left" else 0)
    for i in range(LOG_ROWS):
        cv.hline(140, 1366, hy + 30 + (i + 1) * 22, C["line"])
    cv.build()
    ws.hide_gridlines(2)
    ws.hide_row_col_headers()
    ws.set_zoom(85)
    ws.set_column(cv.n_cols, 16383, 9, cv.fmt({"bg": C["page"]}))


# ----------------------------------------------------------------------------
# Постобработка: макросы фигур-кнопок
# ----------------------------------------------------------------------------
_SP_RE = re.compile(r'<xdr:(sp|pic) macro=""( textlink="[^"]*")?>(\s*<xdr:nv(?:Sp|Pic)Pr>\s*'
                    r'<xdr:cNvPr [^>]*descr="btn:(\w+)")')


def _patch_drawing(xml: str) -> str:
    def repl(m):
        tag, textlink, rest, macro = m.group(1), m.group(2) or "", m.group(3), m.group(4)
        return f'<xdr:{tag} macro="[0]!{macro}"{textlink}>{rest}'
    xml = _SP_RE.sub(repl, xml)
    # имена фигур: btn<Макрос>, imgState<N> — по ним VBA находит кнопки и значки
    xml = re.sub(r'name="[^"]*"( descr="btn:(\w+)")', r'name="btn\2"\1', xml)
    xml = re.sub(r'name="[^"]*"( descr="state:(\d)")', r'name="imgState\2"\1', xml)
    xml = re.sub(r'name="[^"]*"( descr="foot:(\d)")', r'name="imgFoot\2"\1', xml)
    # значки состояния: по умолчанию виден только «нет данных» (…:0)
    xml = re.sub(r'(<xdr:cNvPr [^>]*descr="(?:state|foot):[1-9]")', r'\1 hidden="1"', xml)
    # кнопки — со скругленными углами
    xml = re.sub(r'(descr="btn:\w+"(?:(?!</xdr:sp>).)*?<a:prstGeom prst=")rect(")',
                 r"\1roundRect\2", xml, flags=re.S)
    return xml


def postprocess(path: str) -> None:
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/drawings/drawing") and item.filename.endswith(".xml"):
                data = _patch_drawing(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


# ----------------------------------------------------------------------------
# Сборка книги
# ----------------------------------------------------------------------------
def build_workbook(path: str, icons_dir: str, vba_bin: str | None = None) -> BuildResult:
    assets = Assets(icons_dir)
    wb = xlsxwriter.Workbook(path, {"strings_to_numbers": False, "strings_to_formulas": False})
    wb.set_properties({"title": "ЛХОС по СИКН — дашборд", "subject": "Оперативная сводка",
                       "comments": "Сформировано генератором build.py"})
    wb.set_calc_mode("auto")

    sheets = {}
    for name, code in ((S_DASH, "shDash"), (S_OBJ, "shObj"), (S_TREND, "shTrend"),
                       (S_ALARM, "shAlarm"), (S_REPORT, "shReport"), (S_SET, "shSettings"),
                       (S_DB, "shDB"), (S_CALC, "shCalc")):
        ws = wb.add_worksheet(name)
        sheets[name] = ws
        if vba_bin:
            ws.set_vba_name(code)
    if vba_bin:
        wb.set_vba_name("ThisWorkbook")
        wb.add_vba_project(vba_bin)

    build_calc_sheet(wb, sheets[S_CALC])
    build_db_sheet(wb, sheets[S_DB])
    info = build_dash_sheet(wb, sheets[S_DASH], assets)
    texts = {
        1: ("СИКН — карточки объектов",
            "Здесь будут карточки СИКН: история значений, статистика по объекту, вид продукции."),
        2: ("Тренды", "Здесь будут графики содержания ЛХОС по выбранным СИКН за произвольный период."),
        3: ("Тревоги", "Здесь будет журнал выходов за границы зон и превышений норматива."),
        4: ("Отчеты", "Здесь будут печатные формы и выгрузки (месячный и годовой свод)."),
    }
    for idx, (title, text) in texts.items():
        build_placeholder(wb, sheets[NAV[idx][1]], assets, idx, title, text)
    build_settings(wb, sheets[S_SET], assets)

    sheets[S_CALC].hide()
    sheets[S_DASH].activate()
    sheets[S_DASH].set_first_sheet()
    for n in (S_OBJ, S_TREND, S_ALARM, S_REPORT):
        sheets[n].set_tab_color(C["divider"])
    sheets[S_SET].set_tab_color(C["muted"])
    sheets[S_DB].set_tab_color(C["yellow"])
    wb.close()
    postprocess(path)
    return BuildResult(path, info)
