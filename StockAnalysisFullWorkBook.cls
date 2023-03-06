VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()
For Each ws In Worksheets
Dim Ticker As String
Dim TotalVolume As Double
    TotalVolume = 0
Dim YearOpen As Double
Dim YearClose As Double
Dim YearChange As Double
Dim PercentChange As Double
Dim SummaryTableRow As Integer
    SummaryTableRow = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"
ws.Cells(2, "N").Value = "Greatest % Increase"
ws.Cells(3, "N").Value = "Greatest % Decrease"
ws.Cells(4, "N").Value = "Greatest Total Volume"
ws.Cells(1, "O").Value = "Ticker"
ws.Cells(1, "P").Value = "Value"
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1)
        ws.Range("I" & SummaryTableRow).Value = Ticker
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        Range("L" & SummaryTableRow).Value = TotalVolume
        YearClose = ws.Cells(i, 6).Value
        YearlyChange = YearClose - YearOpen
        PercentChange = YearlyChange / YearOpen
        ws.Range("K" & SummaryTableRow).Value = PercentChange
        ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
        ws.Range("J" & SummaryTableRow).Value = YearlyChange
        SummaryTableRow = SummaryTableRow + 1
        TotalVolume = 0
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        YearOpen = ws.Cells(i, 3).Value
    Else
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    End If
    If ws.Range("J" & SummaryTableRow).Value > 0 Then
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        Else
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
    End If
Next i

Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestTotalVolume As Double
Dim SummaryTableLastRow As Integer
    SummaryTableLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
GreatestPercentIncrease = WorksheetFunction.Max(ActiveSheet.Columns("K"))
GreatestPercentDecrease = WorksheetFunction.Min(ActiveSheet.Columns("K"))
GreatestTotalVolume = WorksheetFunction.Max(ActiveSheet.Columns("L"))
For i = 2 To SummaryTableLastRow
    If GreatestPercentIncrease = ws.Cells(i, 11).Value Then
        ws.Range("O2").Value = ws.Cells(i, 9).Value
        ElseIf GreatestPercentDecrease = ws.Cells(i, 11).Value Then
        ws.Range("O3").Value = ws.Cells(i, 9).Value
        ElseIf GreatestTotalVolume = ws.Cells(i, 12).Value Then
        ws.Range("O4").Value = ws.Cells(i, 9).Value
    End If
Next i
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"

Next ws
    
End Sub
