Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    AssignButtonMacros
    EnsureMonthSelected
    Application.Calculate
    ResetStateIcon
    UpdateStateIcon
    FitDashboard
    StartClock
End Sub

Private Sub Workbook_Activate()
    On Error Resume Next
    EnsureClock
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
    On Error Resume Next
    FitDashboard
End Sub

' Книга с формулами СЕГОДНЯ/ТДАТА всегда считается измененной, поэтому вопрос
' о сохранении задаем сами: «Отмена» оставляет книгу открытой и часы идущими.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If Not Me.Saved And Not Me.ReadOnly Then
        Select Case MsgBox("Сохранить изменения в книге «" & Me.Name & "»?", _
                           vbYesNoCancel + vbQuestion, "Дашборд ЛХОС")
            Case vbYes
                Me.Save
            Case vbNo
                Me.Saved = True
            Case Else
                Cancel = True
                Exit Sub
        End Select
    End If
    StopClock
End Sub
