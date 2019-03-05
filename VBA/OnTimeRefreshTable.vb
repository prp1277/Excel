Option Explicit
'------------------------------------------------------------------------------
Dim NextTick As Date

Sub UpdateClock()
'
'This macro will refresh the query tables every 30 seconds

    ThisWorkbook.Sheets(1).Range("A1") = Time
    NextTick = Now + TimeValue("00:00:30")
    Application.OnTime NextTick, "UpdateClock"
    ThisWorkbook.RefreshAll

End Sub
'------------------------------------------------------------------------------
Sub StopClock()

    On Error Resume Next
    Application.OnTime NextTick, "UpdateClock", , False

End Sub

'------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call StopClock
End Sub