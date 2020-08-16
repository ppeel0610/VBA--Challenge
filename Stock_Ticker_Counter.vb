Sub Wallstreet()

    ' Loop / Iterate Through All Worksheets
    For Each ws In Worksheets
        ' Column Headers / Data Field Labels
        Range("H1").Value = "Ticker"
        Range("I1").Value = "Yearly Change"
        Range("J1").Value = "Percent Change"
        Range("K1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        ' Set/Declare Initial Variables And Set Default/Baseline Variables
        Dim TickerName As String
        Dim LastRow As Long
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
        
        ' Determine the Last Row
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
            ' Add To Ticker Total Volume
            TotalTickerVolume = TotalTickerVolume + Cells(i, 7).Value
            ' Check If We Are Still Within The Same Ticker Name If It Is Not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
'This is for my ticker name
                ' Set Ticker Name
                TickerName = Cells(i, 1).Value
                ' Print The Ticker Name In The Summary Table
                Range("H" & SummaryTableRow).Value = TickerName
                ' Print The Ticker Total Amount To The Summary Table
                Range("K" & SummaryTableRow).Value = TotalTickerVolume
                ' Reset Ticker Total
                TotalTickerVolume = 0
' Adding in yearly change information
                ' Set Yearly Open, Yearly Close and Yearly Change Name
                YearlyOpen = Range("C" & PreviousAmount)
                YearlyClose = Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                Range("I" & SummaryTableRow).Value = YearlyChange
                    End
                ' Determine Percent Change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                ' Format Double To Include % Symbol And Two Decimal Places
                Range("J" & SummaryTableRow).NumberFormat = "0.00%"
                Range("J" & SummaryTableRow).Value = PercentChange
                ' Conditional Formatting Highlight Positive (Green) / Negative (Red)
                If Range("I" & SummaryTableRow).Value >= 0 Then
                    Range("I" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    Range("I" & SummaryTableRow).Interior.ColorIndex = 3
                End If
                ' Add One To The Summary Table Row
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            Next i
            ' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            LastRow = Cells(Rows.Count, 11).End(xlUp).Row
            ' Start Loop For Final Results
            For i = 2 To LastRow
                If Range("J" & i).Value > Range("O4").Value Then
                    Range("O2").Value = Range("J" & i).Value
                    Range("O3").Value = Range("J" & i).Value
                End If
                If Range("K" & i).Value < Range("Q3").Value Then
                    Range("Q3").Value = Range("K" & i).Value
                    Range("P3").Value = Range("I" & i).Value
                End If
                If Range("L" & i).Value > ws.Range("Q4").Value Then
                    Range("Q4").Value = Range("L" & i).Value
                    Range("P4").Value = Range("I" & i).Value
                End If
            Next i
        ' Format Double To Include % Symbol And Two Decimal Places
            Range("Q2").NumberFormat = "0.00%"
            Range("Q3").NumberFormat = "0.00%"
        ' Format Table Columns To Auto Fit
        Columns("I:Q").AutoFit
    Next ws
End Sub


