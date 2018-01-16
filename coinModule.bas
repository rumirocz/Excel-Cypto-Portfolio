Attribute VB_Name = "coinModule"
Option Explicit
'-Change values for these variables to match your database
'-sheetToSearch is the worksheet that contains the data from the ticker created with CoinMarketCap's API
'-sheetToUpdate is for the worksheet that contains YOUR data linked to the ticker
'   This is the sheet you entered the formula cryptoInfo on
'-timeToUpdate is a time format, hh:mm:ss, to update YOUR data on the worksheet sheetToUpdate
'   Set this interval to however often you want your data to update
'   I set it to be the same interval as the ticker updates (1 minute)
Private Const sheetToSearch As String = "Ticker"
Public Const sheetToUpdate As String = "Coin Profit & Loss"
Public Const timeToUpdate As String = "00:01:00"

'Function to retrieve row number of the specified coin symbol
Private Function findSymbolRow(symbol As String) As String
    On Error GoTo errorhandler
    
    Dim x As Range
    
    Application.ScreenUpdating = False

    Set x = ThisWorkbook.Sheets(sheetToSearch).Cells.Find(what:=symbol, lookat:=xlWhole)

    findSymbolRow = x.Row
    
    Application.ScreenUpdating = True
    
    Exit Function
    
errorhandler:
    Select Case Err.Number
        Case 9, 91
            Application.ScreenUpdating = True
            findSymbolRow = ""
        Exit Function

        Case Else
            Application.ScreenUpdating = True
        Exit Function
    End Select
End Function

'Function to retrieve the column letter of the specified value
Private Function findValueColumn(value As String) As String
    On Error GoTo errorhandler
    
    Dim x As Range
    
    Application.ScreenUpdating = False
    
    Set x = ThisWorkbook.Sheets(sheetToSearch).Cells.Find(what:=value, lookat:=xlWhole)
    
    findValueColumn = Split(Cells(1, x.Column).Address(True, False), "$")(0)
    
    Application.ScreenUpdating = True
    
    Exit Function
    
errorhandler:
    Select Case Err.Number
        Case 9, 91
            Application.ScreenUpdating = True
            findValueColumn = ""
        Exit Function

        Case Else
            Application.ScreenUpdating = True
        Exit Function
    End Select
End Function

'Function to retrieve cell value of specified coin symbol and search value
Public Function cryptoInfo(symbol As String, value As String) As String
    cryptoInfo = ThisWorkbook.Sheets(sheetToSearch).Range(sheetToSearch + "!" + findValueColumn(value) + findSymbolRow(symbol)).value
End Function

'Time function to refresh your data
Sub timeFunction()
    On Error Resume Next
    ThisWorkbook.Sheets(sheetToUpdate).EnableCalculation = False
    ThisWorkbook.Sheets(sheetToUpdate).EnableCalculation = True
    
    Application.OnTime Now + TimeValue(timeToUpdate), "timeFunction"
End Sub
