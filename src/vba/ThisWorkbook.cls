Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    AssignButtonMacros
    EnsureMonthSelected
    Application.Calculate
    UpdateStateIcon
    StartClock
End Sub

Private Sub Workbook_Activate()
    On Error Resume Next
    StartClock
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    StopClock
End Sub
