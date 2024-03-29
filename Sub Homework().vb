Sub Homework()
    ' Define variables
    Dim I As Long
    Dim LastRow As Long
    Dim Ticker1 As String
    Ticker1 = 0
    Dim Ticker2 As String
    Dim Opening As Double
    Dim Closing As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double
    Opening = 2
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim Summary2 As Integer
    Summary2 = 2
    GreatestIncrease = 0
    GreatestDecrease = 1
    GreatestVolume = 100


    ' Name Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"

'loop through all worksheets
'unable to execute successfully


    ' Defining last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row


' Loop through data to summarize info by ticker symbol - 2.3.6
    For I = 2 To LastRow
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            Ticker1 = Cells(I, 1).Value

            YearlyChange = Cells(I, 6).Value - Cells(Opening, 3).Value
            PercentageChange = (YearlyChange / Cells(Opening, 3).Value) * 1



            Range("I" & Summary_Table_Row).Value = Ticker1
            Range("J" & Summary_Table_Row).Value = YearlyChange
            Range("K" & Summary_Table_Row).Value = PercentageChange
            Range("K:K").NumberFormat = "0.00%"
            Range("P2", "P3").NumberFormat = "0.00%"
            Range("L" & Summary_Table_Row).Value = TotalVolume
            Summary_Table_Row = Summary_Table_Row + 1
            Opening = I + 1
            TotalVolume = 0
        Else
        
            TotalVolume = TotalVolume + Cells(I, 7).Value
        End If
'Color the cells to make positive and negative changes obvious - 2.2.4.
'4 is green and 3 is red
        If Cells(I, 10).Value > 0 Then
        Cells(I, 10).Interior.ColorIndex = 4
        Else
        Cells(I, 10).Interior.ColorIndex = 3
        End If
        
'Output to side table

   If Cells(I, 11).Value > GreatestIncrease Then
        GreatestIncrease = Cells(I, 11).Value
        Ticker2 = Cells(I, 9).Value
        
        Range("P2") = GreatestIncrease
        Range("O2") = Ticker2
    
    End If
    If Cells(I, 11).Value < GreatestDecrease Then
        GreatestDecrease = Cells(I, 11).Value
        Ticker2 = Cells(I, 9).Value
    
        Range("P3") = GreatestDecrease
        Range("O3") = Ticker2
        
    End If
    If Cells(I, 12).Value > GreatestVolume Then
        GreatestVolume = Cells(I, 12).Value
        Ticker2 = Cells(I, 9).Value
    
        Range("P4") = GreatestVolume
        Range("O4") = Ticker2
        
    End If

    Next I

End Sub