Option Explicit
'
' Импорт исходных данных ЛХОС в базу данных дашборда (лист «БД»).
'
' Исходный файл: на каждый месяц отдельный лист. На листе ищется ячейка
' «Дата»; ниже неё идут даты, в нижней строке заголовка (той же, где кончается
' ячейка «Дата») — наименования объектов (СИКН), строкой выше — вид продукции.
' Листы без ячейки «Дата» (например, «Свод») пропускаются.
'
' База (лист «БД»): строка 1 — объекты, строка 2 — вид продукции, со строки 3 —
' даты (по строке на каждый календарный день без пропусков) и значения, ppm.
' Дашборд считает всё формулами по базе, макросы нужны только для загрузки.

Public Const MODE_ASK As Long = 0
Public Const MODE_REPLACE As Long = 1
Public Const MODE_APPEND As Long = 2

Private Const DB_FIRST_ROW As Long = 3
Private Const MAX_DASH_OBJECTS As Long = 40
Private Const MAX_DB_DAYS As Long = 3700
Private Const LOG_COLS As Long = 9
Private Const LOG_MAX_ROWS As Long = 50
Private Const MAX_EXAMPLES As Long = 5

' --- результаты разбора исходного файла
Private srcNames() As String
Private srcGroups() As String
Private srcKeys As Collection
Private srcObjCount As Long
Private recDate() As Long
Private recObj() As Long
Private recVal() As Double
Private recCount As Long
Private recCap As Long
Private srcSheets As String
Private srcSheetCount As Long
Private srcEmptySheets As Long
Private srcTextSkipped As Long
Private srcFileName As String

' ============================================================================
' Точки входа
' ============================================================================

' Кнопка «Импорт данных»: выбор файла и загрузка.
Public Sub ImportData()
    Dim f As Variant
    f = Application.GetOpenFilename( _
        "Файлы Excel (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", , _
        "Выберите файл с данными ЛХОС")
    If VarType(f) = vbBoolean Then Exit Sub
    RunImport CStr(f), MODE_ASK, True
End Sub

' Загрузка без диалогов (автоматизация, проверки).
' mode: 0 — авто (первая загрузка — полная, иначе добавление), 1 — заново, 2 — добавить.
' Возвращает "OK" или "ERR", перевод строки и текст отчета.
Public Function ImportFileSilent(ByVal path As String, ByVal mode As Long) As String
    If mode = MODE_ASK Then
        If DbValueCount() = 0 Then mode = MODE_REPLACE Else mode = MODE_APPEND
    End If
    ImportFileSilent = RunImport(path, mode, False)
End Function

' Кнопка «Очистить базу» (лист «Настройки»).
Public Sub ClearDatabase()
    If MsgBox("Удалить все загруженные данные из базы?" & vbCrLf & _
              "Дашборд будет пустым до следующего импорта.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Очистка базы") <> vbYes Then Exit Sub
    ClearDatabaseSilent
    MsgBox "База данных очищена.", vbInformation, "Очистка базы"
End Sub

Public Sub ClearDatabaseSilent()
    Dim calcMode As Long
    calcMode = GetCalc()
    SetCalc xlCalculationManual
    ClearDbArea
    SetNamedValue "sys_LastImport", Empty
    SetNamedValue "sys_LastFile", Empty
    SetNamedValue "sys_LastMode", Empty
    SetNamedValue "sel_Month", Empty
    AddLogRow "Очистка базы", "", 0, 0, 0, 0, 0, ""
    SetCalc calcMode
    Application.Calculate
    UpdateStateIcon
End Sub

' ============================================================================
' Основной сценарий
' ============================================================================

Private Function RunImport(ByVal path As String, ByVal mode As Long, _
                           ByVal interactive As Boolean) As String
    Dim calcMode As Long, wbSrc As Workbook, openedHere As Boolean
    Dim stage As String, report As String, ans As VbMsgBoxResult
    Dim dbVals As Long, dbObj As Long, dbFirst As Long, dbLast As Long

    calcMode = GetCalc()
    On Error GoTo Fail

    stage = "открытие файла"
    If StrComp(path, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 513, , "Выбран сам файл дашборда. Укажите файл с исходными данными."
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    SetCalc xlCalculationManual
    Set wbSrc = OpenSource(path, openedHere)

    stage = "чтение файла"
    ParseWorkbook wbSrc
    If openedHere Then wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing
    If recCount = 0 Then
        Err.Raise vbObjectError + 514, , "В файле не найдено ни одного значения." & vbCrLf & _
            "Ожидается: на каждом листе месяца ячейка «Дата», под ней даты, " & _
            "справа от заголовка — наименования объектов и значения."
    End If

    stage = "проверка базы"
    DbStats dbVals, dbObj, dbFirst, dbLast
    If mode = MODE_ASK Then
        If dbVals = 0 Then
            mode = MODE_REPLACE
        Else
            Application.ScreenUpdating = True
            ans = MsgBox("В базе уже есть данные: " & dbObj & " объект(ов), " & _
                    Format$(dbVals, "#,##0") & " знач., период " & _
                    DateText(dbFirst) & " – " & DateText(dbLast) & "." & vbCrLf & _
                    "В файле «" & srcFileName & "»: " & Format$(recCount, "#,##0") & _
                    " знач., период " & DateText(MinRecDate()) & " – " & _
                    DateText(MaxRecDate()) & "." & vbCrLf & vbCrLf & _
                    "Как загрузить данные?" & vbCrLf & vbCrLf & _
                    "ДА — загрузить все данные заново (старые данные будут удалены);" & vbCrLf & _
                    "НЕТ — добавить только новые значения (имеющиеся не изменятся);" & vbCrLf & _
                    "ОТМЕНА — ничего не загружать.", _
                    vbYesNoCancel + vbQuestion + vbDefaultButton2, "Импорт данных")
            If ans = vbCancel Then
                RestoreApp calcMode
                RunImport = "ERR" & vbLf & "Импорт отменен пользователем."
                Exit Function
            End If
            If ans = vbYes Then mode = MODE_REPLACE Else mode = MODE_APPEND
            Application.ScreenUpdating = False
        End If
    End If

    stage = "запись в базу"
    report = MergeAndWrite(mode)

    stage = "обновление дашборда"
    RestoreApp calcMode
    Application.Calculate
    UpdateStateIcon
    If interactive Then
        On Error Resume Next
        shDash.Activate
        On Error GoTo 0
        MsgBox report, vbInformation, "Импорт завершен"
    End If
    RunImport = "OK" & vbLf & report
    Exit Function

Fail:
    Dim errText As String
    errText = "Ошибка на этапе «" & stage & "»:" & vbCrLf & Err.Description
    On Error Resume Next
    If openedHere Then
        If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    End If
    RestoreApp calcMode
    If interactive Then MsgBox errText, vbCritical, "Импорт данных"
    RunImport = "ERR" & vbLf & errText
End Function

Private Function OpenSource(ByVal path As String, ByRef openedHere As Boolean) As Workbook
    Dim wb As Workbook, shortName As String
    shortName = Mid$(path, InStrRev(Replace(path, "/", "\"), "\") + 1)
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            openedHere = False
            Set OpenSource = wb
            Exit Function
        End If
        If StrComp(wb.Name, shortName, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 515, , "Уже открыта другая книга с именем «" & shortName & _
                "». Закройте ее и повторите импорт."
        End If
    Next
    Set OpenSource = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True)
    openedHere = True
End Function

Private Sub RestoreApp(ByVal calcMode As Long)
    On Error Resume Next
    SetCalc calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ============================================================================
' Разбор исходного файла
' ============================================================================

Private Sub ResetParse()
    srcObjCount = 0
    ReDim srcNames(1 To 1)
    ReDim srcGroups(1 To 1)
    Set srcKeys = New Collection
    recCount = 0
    recCap = 1024
    ReDim recDate(1 To recCap)
    ReDim recObj(1 To recCap)
    ReDim recVal(1 To recCap)
    srcSheets = ""
    srcSheetCount = 0
    srcEmptySheets = 0
    srcTextSkipped = 0
    srcFileName = ""
End Sub

Private Sub ParseWorkbook(ByVal wb As Workbook)
    Dim ws As Worksheet, nVals As Long
    ResetParse
    srcFileName = wb.Name
    For Each ws In wb.Worksheets
        nVals = 0
        If ParseSheet(ws, nVals) Then
            If nVals > 0 Then
                srcSheetCount = srcSheetCount + 1
                If Len(srcSheets) > 0 Then srcSheets = srcSheets & ", "
                srcSheets = srcSheets & ws.Name
            Else
                srcEmptySheets = srcEmptySheets + 1
            End If
        End If
    Next ws
End Sub

' Ищет ячейку «Дата» в левом верхнем углу листа.
Private Function FindDateHeader(ByVal ws As Worksheet) As Range
    Dim a As Variant, r As Long, c As Long, s As String
    a = ws.Range("A1:Z40").Value
    For r = 1 To 40
        For c = 1 To 26
            If VarType(a(r, c)) = vbString Then
                s = LCase$(CleanText(a(r, c)))
                If s = "дата" Or s = "дата:" Then
                    Set FindDateHeader = ws.Cells(r, c)
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

' Разбирает лист месяца. Возвращает False, если лист не похож на лист данных.
Private Function ParseSheet(ByVal ws As Worksheet, ByRef nVals As Long) As Boolean
    Dim hdr As Range, area As Range, used As Range
    Dim nameRow As Long, groupRow As Long, dateCol As Long
    Dim firstRow As Long, lastRow As Long, lastCol As Long
    Dim hdrNames As Variant, data As Variant, objMap() As Long
    Dim nc As Long, r As Long, c As Long, d As Long, v As Double
    Dim nm As String, grp As String

    Set hdr = FindDateHeader(ws)
    If hdr Is Nothing Then Exit Function
    ParseSheet = True

    Set area = hdr.MergeArea
    nameRow = area.Row + area.Rows.Count - 1
    If area.Rows.Count > 1 Then groupRow = area.Row Else groupRow = 0
    dateCol = area.Column
    firstRow = nameRow + 1

    Set used = ws.UsedRange
    lastRow = used.Row + used.Rows.Count - 1
    lastCol = used.Column + used.Columns.Count - 1
    If lastRow < firstRow Or lastCol <= dateCol Then Exit Function

    nc = lastCol - dateCol + 1
    hdrNames = ws.Range(ws.Cells(nameRow, dateCol), ws.Cells(nameRow, lastCol)).Value
    data = ws.Range(ws.Cells(firstRow, dateCol), ws.Cells(lastRow, lastCol)).Value
    If Not IsArray(data) Then Exit Function

    ReDim objMap(1 To nc)
    For c = 2 To nc
        nm = CleanText(hdrNames(1, c))
        If Len(nm) > 0 Then
            grp = ""
            If groupRow > 0 Then
                grp = CleanText(ws.Cells(groupRow, dateCol + c - 1).MergeArea.Cells(1, 1).Value)
            End If
            objMap(c) = RegisterObject(nm, grp)
        End If
    Next c

    For r = 1 To UBound(data, 1)
        d = ToDateSerial(data(r, 1))
        If d > 0 Then
            For c = 2 To nc
                If objMap(c) > 0 Then
                    Select Case ParseValue(data(r, c), v)
                        Case 1
                            AddRecord d, objMap(c), v
                            nVals = nVals + 1
                        Case -1
                            srcTextSkipped = srcTextSkipped + 1
                    End Select
                End If
            Next c
        End If
    Next r
End Function

Private Function RegisterObject(ByVal nm As String, ByVal grp As String) As Long
    Dim key As String, idx As Long
    key = UCase$(nm)
    idx = KeyIndex(srcKeys, key)
    If idx = 0 Then
        srcObjCount = srcObjCount + 1
        idx = srcObjCount
        ReDim Preserve srcNames(1 To idx)
        ReDim Preserve srcGroups(1 To idx)
        srcNames(idx) = nm
        srcGroups(idx) = grp
        srcKeys.Add idx, key
    ElseIf Len(srcGroups(idx)) = 0 Then
        srcGroups(idx) = grp
    End If
    RegisterObject = idx
End Function

Private Sub AddRecord(ByVal d As Long, ByVal o As Long, ByVal v As Double)
    If recCount >= recCap Then
        recCap = recCap * 2
        ReDim Preserve recDate(1 To recCap)
        ReDim Preserve recObj(1 To recCap)
        ReDim Preserve recVal(1 To recCap)
    End If
    recCount = recCount + 1
    recDate(recCount) = d
    recObj(recCount) = o
    recVal(recCount) = v
End Sub

Private Function MinRecDate() As Long
    Dim i As Long, m As Long
    For i = 1 To recCount
        If m = 0 Or recDate(i) < m Then m = recDate(i)
    Next i
    MinRecDate = m
End Function

Private Function MaxRecDate() As Long
    Dim i As Long, m As Long
    For i = 1 To recCount
        If recDate(i) > m Then m = recDate(i)
    Next i
    MaxRecDate = m
End Function

' ============================================================================
' Слияние с базой и запись
' ============================================================================

Private Function MergeAndWrite(ByVal mode As Long) As String
    Dim finNames() As String, finGroups() As String, finKeys As Collection, finCount As Long
    Dim dbNames As Variant, dbGroups As Variant, dbDates As Variant, dbData As Variant
    Dim dbRows As Long, dbObj As Long
    Dim srcToFin() As Long, newObjects As String, newObjCount As Long
    Dim minD As Long, maxD As Long, nDays As Long
    Dim mat() As Variant, srcMat() As Variant
    Dim i As Long, r As Long, c As Long, fc As Long, d As Long
    Dim total As Long, added As Long, same As Long, conflicts As Long
    Dim examples As String, rep As String, oldV As Double, newV As Double

    Set finKeys = New Collection
    ReDim finNames(1 To 1)
    ReDim finGroups(1 To 1)

    ' 1. Объекты: в режиме добавления сначала объекты базы (порядок сохраняется).
    If mode = MODE_APPEND Then
        ReadDb dbNames, dbGroups, dbDates, dbData, dbRows, dbObj
        For c = 1 To dbObj
            finCount = finCount + 1
            ReDim Preserve finNames(1 To finCount)
            ReDim Preserve finGroups(1 To finCount)
            finNames(finCount) = CleanText(dbNames(1, c))
            finGroups(finCount) = CleanText(dbGroups(1, c))
            AddKeySafe finKeys, UCase$(finNames(finCount)), finCount
        Next c
    End If
    ReDim srcToFin(1 To srcObjCount)
    For i = 1 To srcObjCount
        fc = KeyIndex(finKeys, UCase$(srcNames(i)))
        If fc = 0 Then
            finCount = finCount + 1
            ReDim Preserve finNames(1 To finCount)
            ReDim Preserve finGroups(1 To finCount)
            finNames(finCount) = srcNames(i)
            finGroups(finCount) = srcGroups(i)
            finKeys.Add finCount, UCase$(srcNames(i))
            fc = finCount
            If mode = MODE_APPEND Then
                newObjCount = newObjCount + 1
                If newObjCount <= 10 Then
                    If Len(newObjects) > 0 Then newObjects = newObjects & ", "
                    newObjects = newObjects & srcNames(i)
                End If
            End If
        ElseIf Len(finGroups(fc)) = 0 Then
            finGroups(fc) = srcGroups(i)
        End If
        srcToFin(i) = fc
    Next i

    ' 2. Диапазон дат.
    minD = MinRecDate()
    maxD = MaxRecDate()
    If mode = MODE_APPEND Then
        For r = 1 To dbRows
            d = ToDateSerial(dbDates(r, 1))
            If d > 0 Then
                For c = 1 To dbObj
                    If IsNumber(dbData(r, c)) Then
                        If d < minD Then minD = d
                        If d > maxD Then maxD = d
                        Exit For
                    End If
                Next c
            End If
        Next r
    End If
    nDays = maxD - minD + 1
    If nDays > MAX_DB_DAYS Then
        Err.Raise vbObjectError + 516, , "Слишком длинный период данных: " & nDays & _
            " дн. (допускается не более " & MAX_DB_DAYS & ")."
    End If

    ' 3. Матрица итоговых данных (строки — дни, столбцы — объекты).
    ReDim mat(1 To nDays, 1 To finCount)
    If mode = MODE_APPEND Then
        For r = 1 To dbRows
            d = ToDateSerial(dbDates(r, 1))
            If d > 0 Then
                For c = 1 To dbObj
                    If IsNumber(dbData(r, c)) Then mat(d - minD + 1, c) = CDbl(dbData(r, c))
                Next c
            End If
        Next r
    End If

    ' Значения файла (при повторе даты/объекта в файле берется последнее).
    ReDim srcMat(1 To nDays, 1 To srcObjCount)
    For i = 1 To recCount
        srcMat(recDate(i) - minD + 1, recObj(i)) = recVal(i)
    Next i
    For r = 1 To nDays
        For c = 1 To srcObjCount
            If Not IsEmpty(srcMat(r, c)) Then
                total = total + 1
                fc = srcToFin(c)
                newV = srcMat(r, c)
                If IsEmpty(mat(r, fc)) Then
                    mat(r, fc) = newV
                    added = added + 1
                Else
                    oldV = mat(r, fc)
                    If Abs(oldV - newV) < 0.0000001 Then
                        same = same + 1
                    Else
                        conflicts = conflicts + 1
                        If conflicts <= MAX_EXAMPLES Then
                            examples = examples & vbCrLf & "   • " & finNames(fc) & ", " & _
                                DateText(minD + r - 1) & ": в базе " & NumText(oldV) & _
                                ", в файле " & NumText(newV)
                        End If
                    End If
                End If
            End If
        Next c
    Next r

    ' 4. Запись.
    WriteDb finNames, finGroups, finCount, mat, nDays, minD

    ' 5. Отчет, журнал, выбор месяца.
    rep = "Файл: " & srcFileName & vbCrLf
    rep = rep & "Листов с данными: " & srcSheetCount
    If srcSheetCount > 0 And srcSheetCount <= 12 Then rep = rep & " (" & srcSheets & ")"
    rep = rep & vbCrLf
    If srcEmptySheets > 0 Then rep = rep & "Листов без значений: " & srcEmptySheets & vbCrLf
    rep = rep & "Значений в файле: " & Format$(total, "#,##0") & vbCrLf & vbCrLf
    If mode = MODE_REPLACE Then
        rep = rep & "Режим: полная загрузка (база заполнена заново)" & vbCrLf
        rep = rep & "Загружено значений: " & Format$(added, "#,##0") & vbCrLf
    Else
        rep = rep & "Режим: добавление новых значений" & vbCrLf
        rep = rep & "Добавлено новых значений: " & Format$(added, "#,##0") & vbCrLf
        rep = rep & "Уже были в базе: " & Format$(same, "#,##0") & vbCrLf
        If conflicts > 0 Then
            rep = rep & "Отличаются от базы (оставлены как в базе): " & conflicts & examples
            If conflicts > MAX_EXAMPLES Then rep = rep & vbCrLf & "   … и еще " & (conflicts - MAX_EXAMPLES)
            rep = rep & vbCrLf & "   Чтобы заменить их, выполните импорт с полной перезагрузкой." & vbCrLf
        End If
        If newObjCount > 0 Then rep = rep & "Новых объектов: " & newObjCount & " (" & newObjects & ")" & vbCrLf
    End If
    rep = rep & "Объектов в базе: " & finCount & vbCrLf
    rep = rep & "Период данных в базе: " & DateText(minD) & " – " & DateText(maxD)
    If srcTextSkipped > 0 Then
        rep = rep & vbCrLf & "Пропущено нечисловых значений: " & srcTextSkipped
    End If
    If finCount > MAX_DASH_OBJECTS Then
        rep = rep & vbCrLf & vbCrLf & "Внимание: на дашборд выводятся первые " & MAX_DASH_OBJECTS & _
            " объектов из " & finCount & "."
    End If

    SetNamedValue "sys_LastImport", Now
    SetNamedValue "sys_LastFile", srcFileName
    If mode = MODE_REPLACE Then
        SetNamedValue "sys_LastMode", "Полная загрузка"
    Else
        SetNamedValue "sys_LastMode", "Добавление новых"
    End If
    AddLogRow IIf(mode = MODE_REPLACE, "Полная загрузка", "Добавление новых"), srcFileName, _
        total, added, same, conflicts, finCount, DateText(minD) & " – " & DateText(maxD)
    SelectMonthFor minD, maxD

    MergeAndWrite = rep
End Function

' Чтение базы в массивы. dbNames/dbGroups — (1, 1..dbObj); dbDates — (1..dbRows, 1).
Private Sub ReadDb(ByRef dbNames As Variant, ByRef dbGroups As Variant, ByRef dbDates As Variant, _
                   ByRef dbData As Variant, ByRef dbRows As Long, ByRef dbObj As Long)
    Dim hdr As Variant, col As Variant, c As Long, r As Long
    dbRows = 0
    dbObj = 0
    hdr = shDB.Range(shDB.Cells(1, 2), shDB.Cells(1, 2 + 1023)).Value
    For c = 1 To 1024
        If Len(CleanText(hdr(1, c))) = 0 Then Exit For
        dbObj = c
    Next c
    col = shDB.Range(shDB.Cells(DB_FIRST_ROW, 1), shDB.Cells(DB_FIRST_ROW + MAX_DB_DAYS - 1, 1)).Value
    For r = 1 To MAX_DB_DAYS
        If ToDateSerial(col(r, 1)) = 0 Then Exit For
        dbRows = r
    Next r
    If dbObj = 0 Or dbRows = 0 Then
        dbObj = 0
        dbRows = 0
        Exit Sub
    End If
    dbNames = shDB.Range(shDB.Cells(1, 2), shDB.Cells(1, 1 + dbObj)).Value
    dbGroups = shDB.Range(shDB.Cells(2, 2), shDB.Cells(2, 1 + dbObj)).Value
    dbDates = shDB.Range(shDB.Cells(DB_FIRST_ROW, 1), shDB.Cells(DB_FIRST_ROW + dbRows - 1, 1)).Value
    dbData = shDB.Range(shDB.Cells(DB_FIRST_ROW, 2), shDB.Cells(DB_FIRST_ROW + dbRows - 1, 1 + dbObj)).Value
    ' одиночная ячейка возвращается не массивом — приводим к массиву
    If Not IsArray(dbNames) Then dbNames = OneCell(dbNames)
    If Not IsArray(dbGroups) Then dbGroups = OneCell(dbGroups)
    If Not IsArray(dbDates) Then dbDates = OneCell(dbDates)
    If Not IsArray(dbData) Then dbData = OneCell(dbData)
End Sub

' Количество значений, объектов и период базы.
Private Sub DbStats(ByRef nVals As Long, ByRef nObj As Long, ByRef firstD As Long, ByRef lastD As Long)
    Dim dbNames As Variant, dbGroups As Variant, dbDates As Variant, dbData As Variant
    Dim dbRows As Long, r As Long, c As Long, d As Long
    nVals = 0
    firstD = 0
    lastD = 0
    ReadDb dbNames, dbGroups, dbDates, dbData, dbRows, nObj
    For r = 1 To dbRows
        d = ToDateSerial(dbDates(r, 1))
        For c = 1 To nObj
            If IsNumber(dbData(r, c)) Then
                nVals = nVals + 1
                If firstD = 0 Or d < firstD Then firstD = d
                If d > lastD Then lastD = d
            End If
        Next c
    Next r
End Sub

Public Function DbValueCount() As Long
    Dim n As Long, o As Long, f As Long, l As Long
    DbStats n, o, f, l
    DbValueCount = n
End Function

Private Sub WriteDb(finNames() As String, finGroups() As String, ByVal finCount As Long, _
                    mat() As Variant, ByVal nDays As Long, ByVal minD As Long)
    Dim hdrN() As Variant, hdrG() As Variant, dates() As Variant, c As Long, r As Long
    ClearDbArea
    ReDim hdrN(1 To 1, 1 To finCount)
    ReDim hdrG(1 To 1, 1 To finCount)
    For c = 1 To finCount
        hdrN(1, c) = finNames(c)
        hdrG(1, c) = finGroups(c)
    Next c
    ReDim dates(1 To nDays, 1 To 1)
    For r = 1 To nDays
        dates(r, 1) = CDate(minD + r - 1)
    Next r
    With shDB
        .Range(.Cells(1, 2), .Cells(1, 1 + finCount)).Value = hdrN
        .Range(.Cells(2, 2), .Cells(2, 1 + finCount)).Value = hdrG
        .Range(.Cells(DB_FIRST_ROW, 1), .Cells(DB_FIRST_ROW + nDays - 1, 1)).NumberFormat = "dd.mm.yyyy"
        .Range(.Cells(DB_FIRST_ROW, 1), .Cells(DB_FIRST_ROW + nDays - 1, 1)).Value = dates
        .Range(.Cells(DB_FIRST_ROW, 2), .Cells(DB_FIRST_ROW + nDays - 1, 1 + finCount)).Value = mat
    End With
End Sub

Private Sub ClearDbArea()
    Dim used As Range, lastRow As Long, lastCol As Long
    With shDB
        Set used = .UsedRange
        lastRow = used.Row + used.Rows.Count - 1
        lastCol = used.Column + used.Columns.Count - 1
        If lastRow < DB_FIRST_ROW + MAX_DB_DAYS Then lastRow = DB_FIRST_ROW + MAX_DB_DAYS
        If lastCol < 2 + MAX_DASH_OBJECTS Then lastCol = 2 + MAX_DASH_OBJECTS
        .Range(.Cells(1, 2), .Cells(2, lastCol)).ClearContents
        .Range(.Cells(DB_FIRST_ROW, 1), .Cells(lastRow, lastCol)).ClearContents
    End With
End Sub

' ============================================================================
' Журнал, служебные значения, выбор месяца
' ============================================================================

Private Sub AddLogRow(ByVal modeText As String, ByVal fileName As String, ByVal total As Long, _
                      ByVal added As Long, ByVal same As Long, ByVal conflicts As Long, _
                      ByVal objCount As Long, ByVal period As String)
    Dim hdr As Range, blk As Range, a As Variant, r As Long, c As Long
    On Error GoTo Done
    Set hdr = ThisWorkbook.Names("log_Header").RefersToRange
    Set blk = hdr.Offset(1, 0).Resize(LOG_MAX_ROWS, LOG_COLS)
    a = blk.Value
    For r = LOG_MAX_ROWS To 2 Step -1
        For c = 1 To LOG_COLS
            a(r, c) = a(r - 1, c)
        Next c
    Next r
    a(1, 1) = Now
    a(1, 2) = fileName
    a(1, 3) = modeText
    a(1, 4) = total
    a(1, 5) = added
    a(1, 6) = same
    a(1, 7) = conflicts
    a(1, 8) = objCount
    a(1, 9) = period
    blk.Value = a
Done:
End Sub

Public Sub SetNamedValue(ByVal nm As String, ByVal v As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Value = v
End Sub

' Выбирает на дашборде месяц: текущий, если он есть в данных, иначе последний.
Private Sub SelectMonthFor(ByVal minD As Long, ByVal maxD As Long)
    Dim t As Long
    t = CLng(Date)
    If t > maxD Or t < minD Then t = maxD
    SetNamedValue "sel_Month", MonthLabel(CDate(t))
End Sub

Public Function MonthLabel(ByVal d As Date) As String
    MonthLabel = Choose(Month(d), "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", _
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь") & " " & Year(d)
End Function

' ============================================================================
' Вспомогательные функции
' ============================================================================

Private Function CleanText(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    s = CStr(v)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr(160), " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = s
End Function

' Дата (число, дата или текст) -> порядковый номер дня; 0, если это не дата.
Private Function ToDateSerial(ByVal v As Variant) As Long
    Dim d As Double
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbDate
            d = CDbl(v)
        Case vbString
            If IsDate(v) Then d = CDbl(CDate(v))
    End Select
    If d >= 36526 And d < 73051 Then ToDateSerial = CLng(Int(d))
End Function

' 1 — число (результат в outVal), 0 — пусто, -1 — нечисловое значение.
Private Function ParseValue(ByVal v As Variant, ByRef outVal As Double) As Long
    Dim s As String
    Select Case VarType(v)
        Case vbEmpty, vbNull
            ParseValue = 0
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal
            outVal = CDbl(v)
            ParseValue = 1
        Case vbString
            s = Replace(CleanText(v), " ", "")
            If Len(s) = 0 Then
                ParseValue = 0
            Else
                s = Replace(s, ",", ".")
                If Left$(s, 1) = "<" Then s = Mid$(s, 2)
                If IsPlainNumber(s) Then
                    outVal = Val(s)
                    ParseValue = 1
                Else
                    ParseValue = -1
                End If
            End If
        Case Else
            ParseValue = -1
    End Select
End Function

' Строка вида [+-]digits[.digits] (десятичный разделитель — точка).
Private Function IsPlainNumber(ByVal s As String) As Boolean
    Dim i As Long, ch As String, digits As Long, dots As Long
    If Left$(s, 1) = "-" Or Left$(s, 1) = "+" Then s = Mid$(s, 2)
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        Else
            Exit Function
        End If
    Next i
    IsPlainNumber = (digits > 0)
End Function

Private Function IsNumber(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal
            IsNumber = True
    End Select
End Function

Private Function KeyIndex(ByVal col As Collection, ByVal key As String) As Long
    On Error GoTo NotFound
    KeyIndex = col.Item(key)
    Exit Function
NotFound:
    KeyIndex = 0
End Function

Private Sub AddKeySafe(ByVal col As Collection, ByVal key As String, ByVal idx As Long)
    If KeyIndex(col, key) = 0 Then col.Add idx, key
End Sub

Private Function OneCell(ByVal v As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    a(1, 1) = v
    OneCell = a
End Function

Private Function DateText(ByVal d As Long) As String
    If d <= 0 Then
        DateText = "—"
    Else
        DateText = Format$(CDate(d), "dd.mm.yyyy")
    End If
End Function

Private Function NumText(ByVal v As Double) As String
    NumText = CStr(Round(v, 4))
End Function

Private Function GetCalc() As Long
    On Error Resume Next
    GetCalc = xlCalculationAutomatic
    GetCalc = Application.Calculation
End Function

Private Sub SetCalc(ByVal mode As Long)
    On Error Resume Next
    Application.Calculation = mode
End Sub
