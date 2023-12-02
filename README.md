# VBA-challenge
Homework 2 - VBA Scripting

'Run VBA on All Sheet
Sub RunAnalysisOnAllSheets()
Dim ws As Worksheet
'Loop through all sheets in the workbook
For Each ws In ThisWorkbook.Sheets
'Activate the current sheet
ws.Activate
'Run your existing Sub Ticker
Call Ticker
Next ws
End Sub

*Warning for Grader,if add Run VBA on All Sheet. Laptop will feel like it's slowing down and get super slow*

Sub Ticker()
'Define
Dim Ticker_Name As String
Dim Ticker_Total As Double
Ticker_Total = 0
Dim Ticker_open As Double
Dim Ticker_close As Double
Dim Yearly_Change As Double
Dim Summary_Table_Row As Integer
Dim increase As Variant
Dim Percent_Increase As Variant
Summary_Table_Row = 2

'Name
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Create & Paste Data to Row I & J
last_row = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To last_row

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker_Name = Cells(i, 1).Value
Ticker_Total = Ticker_Total + Cells(i, 7).Value

Range("I" & Summary_Table_Row).Value = Ticker_Name
Range("L" & Summary_Table_Row).Value = Ticker_Total
Summary_Table_Row = Summary_Table_Row + 1
Ticker_Total = 0
Else
Ticker_Total = Ticker_Total + Cells(i, 7).Value
End If

'Create & Paste Date to Row J & K
Next i
Summary_Table_Row = 2
For i = 2 To last_row
If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
Ticker_close = Cells(i, 6).Value
ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
Ticker_open = Cells(i, 3).Value
End If
If Ticker_open > 0 And Ticker_close > 0 Then
increase = Ticker_close - Ticker_open
Percent_Increase = increase / Ticker_open
Cells(Summary_Table_Row, 10).Value = increase
Cells(Summary_Table_Row, 11).Value = FormatPercent(Percent_Increase)
Ticker_close = 0
Ticker_open = 0
Summary_Table_Row = Summary_Table_Row + 1
End If

'Put Colors for Row J & leave space empty if blank
If IsEmpty(Cells(i, 10).Value) Then Exit For
If Cells(i, 10).Value > 0 Then
Cells(i, 10).Interior.ColorIndex = 4
Else
Cells(i, 10).Interior.ColorIndex = 3
End If
'Greatest % Increase & Decrease
Next i
max_per = WorksheetFunction.Max(ActiveSheet.Columns("k"))
min_per = WorksheetFunction.Min(ActiveSheet.Columns("k"))
max_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))
Range("Q2").Value = FormatPercent(max_per)
Range("Q3").Value = FormatPercent(min_per)
Range("Q4").Value = max_vol

Summary_Table_Row = 2
For i = 2 To last_row

If max_per = Cells(i, 11).Value Then
Range("P2").Value = Cells(i, 9).Value
ElseIf min_per = Cells(i, 11).Value Then
Range("P3").Value = Cells(i, 9).Value
ElseIf max_vol = Cells(i, 12).Value Then
Range("P4").Value = Cells(i, 9).Value
End If

Next i

End Sub
