Attribute VB_Name = "Module2"
Sub RunAnalysisOnAllSheets()

Dim ws As Worksheet
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
'Loop through all sheets in the workbook
For Each ws In ThisWorkbook.Sheets
'Activate the current sheet
ws.Activate
'Run your existing Sub Ticker
Call Ticker
Next ws
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
