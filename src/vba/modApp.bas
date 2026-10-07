Option Explicit
'
' Служебные процедуры дашборда: «часы» (обновление надписи о времени с момента
' загрузки), назначение макросов кнопкам, выбор месяца, выход.

Private nextTick As Date
Private clockOn As Boolean
Private lastDay As Long

' Запускает ежеминутное обновление надписи «Обновление: N мин назад».
Public Sub StartClock()
    On Error Resume Next
    StopClock
    nextTick = Now + TimeSerial(0, 1, 0)
    Application.OnTime nextTick, "TickClock"
    clockOn = True
End Sub

Public Sub StopClock()
    On Error Resume Next
    If clockOn Then Application.OnTime nextTick, "TickClock", , False
    clockOn = False
End Sub

Public Sub TickClock()
    On Error Resume Next
    clockOn = False
    RefreshClock
    StartClock
End Sub

' Пересчитывает надписи со временем; при смене суток — всю книгу
' (текущая дата влияет на выбор «последнего актуального значения»).
Public Sub RefreshClock()
    On Error Resume Next
    If lastDay <> CLng(Date) Then
        lastDay = CLng(Date)
        Application.Calculate
    Else
        ThisWorkbook.Names("ui_Elapsed").RefersToRange.Calculate
        ThisWorkbook.Names("ui_Updated").RefersToRange.Calculate
    End If
End Sub

' Показывает значок-щит, соответствующий общему состоянию (имя c_State):
' картинки imgState0..4 (щит) и imgFoot0..4 (подвал), видимы только нужные.
Public Sub UpdateStateIcon()
    Dim st As Long, i As Long, shp As Shape, prefix As Variant
    On Error Resume Next
    st = CLng(ThisWorkbook.Names("c_State").RefersToRange.Value)
    For Each prefix In Array("imgState", "imgFoot")
        For i = 0 To 4
            Set shp = Nothing
            Set shp = shDash.Shapes(prefix & i)
            If Not shp Is Nothing Then shp.Visible = (i = st)
        Next i
    Next prefix
End Sub

' Назначает макросы фигурам-кнопкам: фигура «btnИмяМакроса» запускает ИмяМакроса.
Public Sub AssignButtonMacros()
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            If Left$(shp.Name, 3) = "btn" Then shp.OnAction = Mid$(shp.Name, 4)
        Next shp
    Next ws
End Sub

' Если выбранный месяц отсутствует в списке — выбирает последний месяц с данными.
Public Sub EnsureMonthSelected()
    Dim cur As String, lst As Range, c As Range, first As String
    On Error Resume Next
    cur = CStr(ThisWorkbook.Names("sel_Month").RefersToRange.Value)
    Set lst = ThisWorkbook.Names("m_Labels").RefersToRange
    If lst Is Nothing Then Exit Sub
    For Each c In lst.Cells
        If Len(CStr(c.Value)) > 0 Then
            If Len(first) = 0 Then first = CStr(c.Value)
            If CStr(c.Value) = cur Then Exit Sub
        End If
    Next c
    If Len(first) > 0 Then
        ThisWorkbook.Names("sel_Month").RefersToRange.Value = first
    End If
End Sub

' Кнопка «Выйти»: закрывает книгу (Excel предложит сохранить изменения).
Public Sub ExitApp()
    StopClock
    ThisWorkbook.Close
End Sub

' Кнопка «Обновить»: пересчет всех формул.
Public Sub RefreshAll()
    Application.CalculateFull
End Sub
