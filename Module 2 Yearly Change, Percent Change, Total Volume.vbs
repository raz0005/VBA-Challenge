Sub StockAnalysis()

    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    
    ' Initialize total volume variable
    totalVolume = 0
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set initial values for summary
        summaryRow = 2
        ws.Cells(summaryRow, 9).Value = "Ticker"
        ws.Cells(summaryRow, 10).Value = "Yearly Change"
        ws.Cells(summaryRow, 11).Value = "Percent Change"
        ws.Cells(summaryRow, 12).Value = "Total Volume"
        
        ' Loop through rows to analyze data
        For i = 2 To lastRow
            
            ' Check if the ticker symbol changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Get ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Get opening and closing prices
                openingPrice = ws.Cells(i, 3).Value
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change
                yearlyChange = closingPrice - openingPrice
                
                ' Calculate percent change
                If openingPrice <> 0 Then
                    percentChange = (yearlyChange / openingPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Get total stock volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Output results
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Move to the next row for summary
                summaryRow = summaryRow + 1
                
                ' Reset total volume for the next ticker
                totalVolume = 0
                
            Else
                ' Accumulate total volume for the same ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
    Next ws

End Sub
