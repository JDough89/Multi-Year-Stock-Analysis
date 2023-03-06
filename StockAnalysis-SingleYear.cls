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
Dim Ticker As String
Dim TotalVolume As Double
    TotalVolume = 0
Dim YearOpen As Double
Dim YearClose As Double
Dim YearChange As Double
Dim PercentChange As Double
Dim SummaryTableRow As Integer
    SummaryTableRow = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"
Cells(2, "N").Value = "Greatest % Increase"
Cells(3, "N").Value = "Greatest % Decrease"
Cells(4, "N").Value = "Greatest Total Volume"
Cells(1, "O").Value = "Ticker"
Cells(1, "P").Value = "Value"
For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1)
        Range("I" & SummaryTableRow).Value = Ticker
        TotalVolume = TotalVolume + Cells(i, 7).Value
        Range("L" & SummaryTableRow).Value = TotalVolume
        YearClose = Cells(i, 6).Value
        YearlyChange = YearClose - YearOpen
        PercentChange = YearlyChange / YearOpen
        Range("K" & SummaryTableRow).Value = PercentChange
        Range("K" & SummaryTableRow).NumberFormat = "0.00%"
        Range("J" & SummaryTableRow).Value = YearlyChange
        SummaryTableRow = SummaryTableRow + 1
        TotalVolume = 0
    ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        YearOpen = Cells(i, 3).Value
    Else
        TotalVolume = TotalVolume + Cells(i, 7).Value
    End If
    If Range("J" & SummaryTableRow).Value > 0 Then
        Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        Else
        Range("J" & SummaryTableRow).Interior.ColorIndex = 3
    End If
Next i

Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestTotalVolume As Double
Dim SummaryTableLastRow As Integer
    SummaryTableLastRow = Cells(Rows.Count, 9).End(xlUp).Row
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
GreatestPercentIncrease = WorksheetFunction.Max(ActiveSheet.Columns("K"))
GreatestPercentDecrease = WorksheetFunction.Min(ActiveSheet.Columns("K"))
GreatestTotalVolume = WorksheetFunction.Max(ActiveSheet.Columns("L"))
For i = 2 To SummaryTableLastRow
    If GreatestPercentIncrease = Cells(i, 11).Value Then
        Range("O2").Value = Cells(i, 9).Value
        ElseIf GreatestPercentDecrease = Cells(i, 11).Value Then
        Range("O3").Value = Cells(i, 9).Value
        ElseIf GreatestTotalVolume = Cells(i, 12).Value Then
        Range("O4").Value = Cells(i, 9).Value
    End If
Next i
Range("P2").Value = GreatestPercentIncrease
Range("P2").NumberFormat = "0.00%"
Range("P3").Value = GreatestPercentDecrease
Range("P3").NumberFormat = "0.00%"
Range("P4").Value = GreatestTotalVolume
    
End Sub
