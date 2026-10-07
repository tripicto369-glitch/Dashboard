Option Explicit

' Смена месяца в шапке — пересчет (на случай ручного режима вычислений).
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, ThisWorkbook.Names("sel_Month").RefersToRange) Is Nothing Then
        Application.Calculate
    End If
End Sub

' После каждого пересчета — значок общего состояния.
Private Sub Worksheet_Calculate()
    UpdateStateIcon
End Sub
