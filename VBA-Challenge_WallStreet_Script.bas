Attribute VB_Name = "Module1"
'This Function is running a for loop to complete the Summary Table in Columns I to L.
Sub Breakdown_Stock_Ticker():

Dim ws As Worksheet
    
'Variable for Column A
Dim StockTicker As String
    
'Loop to execute on each Worksheet
For Each ws In Worksheets

     'Variable for Total Stock Volume
        Dim TotalStock As Variant
            TotalStock = 0
    
        'Variable establishing the beginning location of the summary table
        Dim SummaryTable As Integer
            SummaryTable = 2

        'Establishes Constant for OpeningPrice
        openingPrice = ws.Cells(2, 3).Value

        'Table's for SummaryTable (Columns I to L)
        ws.Range("I1").Value = "<Ticker>"
        ws.Range("J1").Value = "<Yearly Change>"
        ws.Range("K1").Value = "<Precent Change>"
        ws.Range("L1").Value = "<Total Stock Volume>"

        'Variable to calculate the Last row in each column
        Dim LastRow As Variant
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
        'For loop to sort Stock Calculations
        For Row = 2 To LastRow
            'If Statement to run Stock Ticker, YearlyChange, Conditional Formatting, Percentage Change, and Total Stock Volume.
            If (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value) Then
            'Finding Stock Ticker
            StockTicker = ws.Cells(Row, 1).Value
            ws.Range("I" & SummaryTable).Value = StockTicker
            'Finding Yearly Change
            closingPrice = ws.Cells(Row, 6).Value
            YearlyChange = closingPrice - openingPrice
            ws.Range("J" & SummaryTable).Value = YearlyChange
            'Conditional Formatting Yearly Change Column
                If ws.Cells(SummaryTable, 10) >= 0 Then
                    ws.Cells(SummaryTable, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(SummaryTable, 10).Interior.ColorIndex = 3
                End If
            'Finding Percentage Change
                If openingPrice <> 0 Then
                    PercentChange = ((closingPrice - openingPrice) / openingPrice * 100)
                    'Formats PercentChange variable to 2 decimals and sign %
                    PercentChange = Format(PercentChange, "%0.00")
                    ws.Cells(SummaryTable, 11).Value = PercentChange
                Else
                    ws.Cells(SummaryTable, 11).Value = 0
                    openingPrice = ws.Cells(Row + 1, 3).Value
                End If
            'Finding Total Stock Volume
                TotalStock = TotalStock + ws.Cells(Row, 7).Value
                ws.Range("L" & SummaryTable).Value = TotalStock
                SummaryTable = SummaryTable + 1
                TotalStock = 0
            Else
                TotalStock = TotalStock + ws.Cells(Row, 7).Value
        
    End If
        Next Row
        
        'Label's for Challenge Summary Table Cells
        ws.Range("N2").Value = "<Greatest % Increase>"
        ws.Range("N3").Value = "<Greatest % Decrease>"
        ws.Range("N4").Value = "<Highest Total Volume>"
        ws.Range("O1").Value = "<Stock Ticker>"
        ws.Range("P1").Value = "<Highest Total Volume>"
        
        Dim GreatestTotalVolume As Variant
            GreatestTotalVolume = ws.Cells(2, 12).Value
        Dim GreatestPercentIncrease As Variant
            GreatestPercentIncrease = ws.Cells(2, 11).Value
        Dim GreatestPercentDecrease As Variant
            GreatestPercentDecrease = ws.Cells(2, 11).Value
            
        RowCount = ws.Cells(Rows.Count, 12).End(xlUp).Row
        
        For i = 2 To RowCount
   
            If ws.Cells(i + 1, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = ws.Cells(i + 1, 12).Value
                ws.Range("Q4").Value = GreatestTotalVolume
                ws.Range("P4").Value = ws.Cells(i + 1, 9).Value
            End If
            If ws.Cells(i + 1, 11).Value > GreatestPercentIncrease Then
                GreatestPercentIncrease = ws.Cells(i + 1, 11).Value
                ws.Range("Q2").Value = GreatestPercentIncrease
                'GreatestPercentIncrease = Format(GreatestPercentIncrease, "%0.00")
                ws.Range("P2").Value = ws.Cells(i + 1, 9).Value
            End If
            If ws.Cells(i + 1, 11).Value < GreatestPercentDecrease Then
                GreatestPercentDecrease = ws.Cells(i + 1, 11).Value
                ws.Range("Q3").Value = GreatestPercentDecrease
                'GreatestPercentDecrease = Format(GreatestPercentDecrease, "%0.00")
                ws.Range("P3").Value = ws.Cells(i + 1, 9).Value
            End If
                Next i
                
        Next ws
    
End Sub
