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
                    PercentChange = (closingPrice - openingPrice) / openingPrice
                    ws.Cells(SummaryTable, 11).Value = PercentChange
                Else
                    ws.Cells(SummaryTable, 11).Value = dash
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
        Next ws
    
End Sub
