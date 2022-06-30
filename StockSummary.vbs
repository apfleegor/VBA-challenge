Attribute VB_Name = "Module1"
Sub SummaryStocks()

For Each ws In Worksheets

    ' declare / create variables
    Dim totalVolume As LongLong
    Dim summaryTableRow As Integer
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    ' bonus
    Dim greatestPerInc As Double
    Dim PerIncTicker As String
    
    Dim greatestPerDec As Double
    Dim PerDecTicker As String
    
    Dim greatestStockVolume As LongLong
    Dim StockVolumeTicker As String
    
    ' create summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' create bonus table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ' beginning values
    summaryTableRow = 2
    openingPrice = ws.Range("C2").Value
    greatestPerInc = 0
    greatestPerDec = 0
    greatestStockVolume = 0
    
    ' formatting
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    For Row = 2 To lastrow
    
        totalVolume = totalVolume + ws.Cells(Row, 7).Value
        
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
        
            ' fill in summary table
            
            ' get closing price
            closingPrice = ws.Cells(Row, 6).Value
            
            ' ticker
            ws.Cells(summaryTableRow, 9).Value = ws.Cells(Row, 1).Value
            
            ' yearly change from opening price to closing price
            yearlyChange = closingPrice - openingPrice
            
            ' formatting for yearly change
            If yearlyChange < 0 Then
                ws.Cells(summaryTableRow, 10).Font.ColorIndex = 3
                ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(summaryTableRow, 10).Font.ColorIndex = 4
                ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
            End If
            
            ws.Cells(summaryTableRow, 10).Font.ColorIndex = 1
            ws.Cells(summaryTableRow, 10).Value = yearlyChange
            
            ' percent change from opening price to closing price
            percentChange = (closingPrice - openingPrice) / openingPrice
            ws.Cells(summaryTableRow, 11).Value = percentChange
            
            ' total stock volume of the stock
            ws.Cells(summaryTableRow, 12).Value = totalVolume
            
            ' bonus section
            If percentChange > greatestPerInc Then
                greatestPerInc = percentChange
                PerIncString = ws.Cells(Row, 1).Value
                
            ElseIf percentChange < greatestPerDec Then
                greatestPerDec = percentChange
                PerDecString = ws.Cells(Row, 1).Value
            End If
            
            If totalVolume > greatestStockVolume Then
                greatestStockVolume = totalVolume
                StockVolumeString = ws.Cells(Row, 1).Value
            End If
            
            ' reset for next round
            totalVolume = 0
            openingPrice = ws.Cells(Row + 1, 3).Value
            summaryTableRow = summaryTableRow + 1
            
        End If
    
    Next Row
    
    ' fill out bonus section
    ws.Range("P2").Value = PerIncString
    ws.Range("Q2").Value = greatestPerInc
    
    ws.Range("P3").Value = PerDecString
    ws.Range("Q3").Value = greatestPerDec
    
    ws.Range("P4").Value = StockVolumeString
    ws.Range("Q4").Value = greatestStockVolume

Next ws

End Sub

