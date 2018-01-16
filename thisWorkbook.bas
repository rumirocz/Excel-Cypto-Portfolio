Private Sub Workbook_BeforeClose(Cancel As Boolean)
'call stopTimeFunction when closing the workbook
    On Error Resume Next
    stopTimeFunction
End Sub

Private Sub Workbook_Open()
'Call sheetsRecalculate when opening the workbook (which also starts timer)
    On Error Resume Next
    sheetsRecalculate
End Sub
