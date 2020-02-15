Sub SMAnalysis()

For Each ws In Worksheets
    
    'Sorts B column and then A column to ensure data is in the correct order.
'    ws.Columns("A:G").Sort key1:=ws.Range("B2"), _
'          order1:=xlAscending, Header:=xlYes
'    ws.Columns("A:G").Sort key1:=ws.Range("A2"), _
'          order1:=xlAscending, Header:=xlYes
'
    'Creates Summary table Labels
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Creates Greatest Table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Deacrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("Q2").Style = "Percent"
    ws.Range("Q3").Style = "Percent"
    
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim GreatestPercentTicker As String
    Dim WorstPercentTicker As String
    Dim GreatestVolumeTicker As String
    
    
    
    GreatestPercentIncrease = 0
    GreatestPercentDecrease = 0
    GreatestTotalVolume = 0
    
    
        
    Dim SummaryTable As Double
        SummaryTable = 2
        
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    
    TotalStockVolume = 0
    OpeningPrice = ws.Range("c2").Value
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For Row = 2 To lastrow
        TotalStockVolume = TotalStockVolume + ws.Cells(Row, 7).Value

        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then

            ClosingPrice = ws.Cells(Row, 6).Value
            YearlyChange = ClosingPrice - OpeningPrice

            If OpeningPrice = 0 Then
                PercentageChange = 0
            Else
                PercentageChange = (ClosingPrice / OpeningPrice) - 1
            End If
            
            ws.Range("I" & SummaryTable).Value = ws.Cells(Row, 1).Value
            ws.Range("J" & SummaryTable).Value = YearlyChange

            If YearlyChange >= 0 Then
                ws.Range("j" & SummaryTable).Interior.ColorIndex = 4
            Else
                ws.Range("j" & SummaryTable).Interior.ColorIndex = 3
            End If

            ws.Range("K" & SummaryTable).Value = PercentageChange
            ws.Range("K" & SummaryTable).Style = "Percent"
            ws.Range("L" & SummaryTable).Value = TotalStockVolume
            SummaryTable = SummaryTable + 1
            OpeningPrice = ws.Cells(Row + 1, 3).Value
            
            If PercentageChange > GreatestPercentIncrease Then
                GreatestPercentIncrease = PercentageChange
                GreatestPercentTicker = ws.Cells(Row, 1).Value
            End If

            If PercentageChange < GreatestPercentDecrease Then
                GreatestPercentDecrease = PercentageChange
                WorstPercentTicker = ws.Cells(Row, 1).Value

            End If

            If TotalStockVolume > GreatestTotalVolume Then
                GreatestTotalVolume = TotalStockVolume
                GreatestVolumeTicker = ws.Cells(Row, 1).Value
            End If
            
            ws.Range("P2").Value = GreatestPercentTicker
            ws.Range("Q2").Value = GreatestPercentIncrease
            ws.Range("P3").Value = WorstPercentTicker
            ws.Range("Q3").Value = GreatestPercentDecrease
            ws.Range("P4").Value = GreatestVolumeTicker
            ws.Range("Q4").Value = GreatestTotalVolume
            
            TotalStockVolume = 0
            
        End If
    Next Row
Next ws

End Sub
    


