Option Explicit

' Смена месяца в шапке — пересчет (на случай ручного режима вычислений).
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, ThisWorkbook.Names("sel_Month").RefersToRange) Is Nothing Then
        Application.Calculate
    End If
End Sub

' После каждого пересчета — значок общего состояния (меняется только при смене).
Private Sub Worksheet_Calculate()
    UpdateStateIcon
End Sub

' Щелчок по стрелке рядом с периодом выделяет ячейку выбора месяца,
' и Excel показывает кнопку раскрывающегося списка.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim arrow As Range
    On Error Resume Next
    Set arrow = ThisWorkbook.Names("ui_MonthArrow").RefersToRange.MergeArea
    If arrow Is Nothing Then Exit Sub
    If Target.Address = arrow.Address Then
        Application.EnableEvents = False
        ThisWorkbook.Names("sel_Month").RefersToRange.Select
        Application.EnableEvents = True
    End If
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculate
    FitDashboard
End Sub
