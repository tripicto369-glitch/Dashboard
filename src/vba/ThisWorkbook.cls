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
' При закрытии из другого макроса с DisplayAlerts = False — без вопросов.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    StopClock                                   ' пока имя книги в строке OnTime прежнее
    If Not Me.Saved And Application.DisplayAlerts Then
        Select Case MsgBox("Сохранить изменения в книге «" & Me.Name & "»?", _
                           vbYesNoCancel + vbQuestion, "Дашборд ЛХОС")
            Case vbYes
                If Me.ReadOnly Then
                    If Not Application.Dialogs(xlDialogSaveAs).Show Then Cancel = True
                Else
                    Me.Save
                End If
            Case vbNo
                Me.Saved = True
            Case Else
                Cancel = True
        End Select
    End If
    If Cancel Then StartClock Else StopClock
End Sub

' «Сохранить как» меняет имя книги, входящее в строку OnTime: часы перезапускаются.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    StopClock
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error Resume Next
    StartClock
End Sub
