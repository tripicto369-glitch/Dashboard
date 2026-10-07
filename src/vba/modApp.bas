Option Explicit
'
' Служебные процедуры дашборда: «часы» (обновление надписи о времени с момента
' загрузки), значок состояния, назначение макросов кнопкам, выбор месяца,
' масштаб под окно, выход.

Private nextTick As Date
Private clockOn As Boolean
Private tickProc As String
Private lastDay As Long
Private lastBucket As Long
Private shownState As Long
Private stateKnown As Boolean

' ============================================================================
' Часы: ежеминутное обновление надписи «Обновление: N мин назад»
' ============================================================================

Public Sub StartClock()
    On Error Resume Next
    StopClock
    ' полное имя процедуры — чтобы не перепутать с копией дашборда в другой книге
    tickProc = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!modApp.TickClock"
    nextTick = Now + TimeSerial(0, 1, 0)
    Application.OnTime nextTick, tickProc
    clockOn = True
End Sub

' Запускает часы, если они еще не идут (при переключении окон).
Public Sub EnsureClock()
    If Not clockOn Then StartClock
End Sub

Public Sub StopClock()
    On Error Resume Next
    If clockOn Then Application.OnTime nextTick, tickProc, , False
    clockOn = False
End Sub

Public Sub TickClock()
    On Error Resume Next
    ' «осиротевший» вызов (после сброса VBA-проекта) раньше запланированного — не множим цепочки
    If clockOn And Now < nextTick - TimeSerial(0, 0, 5) Then Exit Sub
    clockOn = False
    RefreshClock
    StartClock
End Sub

' Пересчитывает надписи со временем (ячейки с именами ui_Upd*); при смене
' суток — всю книгу: текущая дата влияет на выбор «последнего актуального значения».
Public Sub RefreshClock()
    Dim nm As Name, lastImp As Variant, age As Double, b As Long
    On Error Resume Next
    ' «корзина» свежести данных (нет / < 1 сут / < 3 сут / старше) — цвет индикатора.
    ' Value2: для ячейки в формате даты .Value вернул бы Date, а IsNumeric(Date) = False
    lastImp = ThisWorkbook.Names("sys_LastImport").RefersToRange.Value2
    If VarType(lastImp) = vbDouble Then
        If lastImp > 0 Then
            age = CDbl(Now) - lastImp
            If age < 1 Then
                b = 1
            ElseIf age < 3 Then
                b = 2
            Else
                b = 3
            End If
        End If
    End If
    If lastDay <> CLng(Date) Or b <> lastBucket Then
        lastDay = CLng(Date)
        lastBucket = b
        Application.Calculate
    Else
        For Each nm In ThisWorkbook.Names
            If LCase$(Left$(nm.Name, 6)) = "ui_upd" Then nm.RefersToRange.Calculate
        Next nm
    End If
End Sub

' ============================================================================
' Значок общего состояния
' ============================================================================

' Показывает значки, соответствующие общему состоянию (имя c_State): картинки
' imgState0..4 (щит) и imgFoot0..4 (подвал). Видимость меняется только при смене
' состояния — запись в объекты книги очищает журнал отмены (Ctrl+Z) Excel.
Public Sub UpdateStateIcon()
    Dim st As Long, i As Long, shp As Shape, prefix As Variant, want As Boolean
    On Error Resume Next
    st = CLng(ThisWorkbook.Names("c_State").RefersToRange.Value)
    If stateKnown And st = shownState Then Exit Sub
    For Each prefix In Array("imgState", "imgFoot")
        For i = 0 To 4
            Set shp = Nothing
            Set shp = shDash.Shapes(prefix & i)
            If Not shp Is Nothing Then
                want = (i = st)
                If CBool(shp.Visible) <> want Then shp.Visible = want
            End If
        Next i
    Next prefix
    shownState = st
    stateKnown = True
End Sub

' Сбрасывает запомненное состояние (после открытия книги).
Public Sub ResetStateIcon()
    stateKnown = False
End Sub

' ============================================================================
' Кнопки, месяц, масштаб
' ============================================================================

' Назначает макросы фигурам-кнопкам: фигура «btnИмяМакроса» запускает ИмяМакроса.
Public Sub AssignButtonMacros()
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            If Left$(shp.Name, 3) = "btn" Then
                If shp.OnAction <> Mid$(shp.Name, 4) Then shp.OnAction = Mid$(shp.Name, 4)
            End If
        Next shp
    Next ws
End Sub

' При открытии: текущий месяц, если по нему есть данные, иначе последний месяц с данными.
Public Sub EnsureMonthSelected()
    Dim cur As String, want As String, lst As Range, c As Range, first As String
    Dim hasWant As Boolean
    On Error Resume Next
    cur = CStr(ThisWorkbook.Names("sel_Month").RefersToRange.Value)
    want = MonthLabel(Date)
    Set lst = ThisWorkbook.Names("m_Labels").RefersToRange
    If lst Is Nothing Then Exit Sub
    For Each c In lst.Cells
        If Len(CStr(c.Value)) > 0 Then
            If Len(first) = 0 Then first = CStr(c.Value)
            If CStr(c.Value) = want Then hasWant = True
        End If
    Next c
    If Not hasWant Then want = first
    If Len(want) > 0 And want <> cur Then
        ThisWorkbook.Names("sel_Month").RefersToRange.Value = want
    End If
End Sub

' Масштаб листа «Сводка» под размер окна (макет целиком, от 40 до 100 %).
Public Sub FitDashboard()
    Dim w As Window, z As Double
    On Error GoTo Done
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If Not ActiveSheet Is shDash Then Exit Sub
    Set w = ActiveWindow
    If w Is Nothing Then Exit Sub
    If w.WindowState = xlMinimized Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ThisWorkbook.Names("ui_Canvas").RefersToRange.Select
    w.Zoom = True
    z = w.Zoom
    If z > 100 Then w.Zoom = 100
    If z < 40 Then w.Zoom = 40
    w.ScrollRow = 1
    w.ScrollColumn = 1
    Application.EnableEvents = True   ' shDash.Worksheet_SelectionChange запомнит A1
    shDash.Range("A1").Select
Done:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ============================================================================
' Выход
' ============================================================================

' Кнопка «Выйти»: закрывает книгу (о сохранении спросит Workbook_BeforeClose);
' если других открытых книг нет — закрывает Excel.
Public Sub ExitApp()
    Dim w As Window, n As Long
    On Error Resume Next
    For Each w In Application.Windows
        If w.Visible Then
            If w.Parent.Name <> ThisWorkbook.Name Then n = n + 1
        End If
    Next w
    If n = 0 Then
        Application.Quit
    Else
        ThisWorkbook.Close
    End If
End Sub

' Кнопка «Обновить»: пересчет всех формул.
Public Sub RefreshAll()
    Application.CalculateFull
End Sub
