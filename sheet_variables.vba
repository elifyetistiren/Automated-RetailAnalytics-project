Option Explicit

Public wbMain As Workbook
Public shtRp1 As Worksheet, shtRp2 As Worksheet, shtMain As Worksheet

Public Sub InitializeVariables()

Set wbMain = ActiveWorkbook

With wbMain
    Set shtRp1 = .Sheets("Report 1")
    Set shtRp2 = .Sheets("Report 2")
    Set shtMain = .Sheets("MAIN")
End With


End Sub
