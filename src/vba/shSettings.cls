Option Explicit

' Пороги и подписи применяются сразу, даже если Excel в ручном режиме вычислений.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Application.Calculation <> xlCalculationAutomatic Then Application.Calculate
End Sub
