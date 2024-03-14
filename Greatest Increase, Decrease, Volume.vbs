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
    Dim greatestIncreaseTicker As String
    Dim greatestIncreasePercent As Double
    Dim greatestDecreaseTicker As String
    Dim greatestDecreasePercent As Double
    Dim greatestVolumeTicker As String
    Dim greatestVolume As Double
    Dim isFirstStock As Boolean
    
    ' Initialize variables for greatest increase, decrease, and volume
    greatestIncreasePercent = -1
    greatestDecreasePercent = 1
    greatestVolume = 0
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summaryRow = 2 ' Set initial row for summary
        
        ' Clear previous summary data
        ws.Cells(1, 9).Resize(ws.Rows.Count, 4).Clear
        
        isFirstStock = True
        
        ' Loop through rows to analyze data
        Do While summaryRow <= lastRow
            ' Get start and end row for current stock block
            startRow = summaryRow
            ticker = ws.Cells(startRow, 1).Value
            Do While ws.Cells(summaryRow + 1, 1).Value = ticker
                summaryRow = summaryRow + 1
            Loop
            endRow = summaryRow
            
            ' Get opening and closing prices
            openingPrice = ws.Cells(i, 3).Value 
            closingPrice = ws.Cells(i, 6).Value 
            
            ' Calculate yearly change
            yearlyChange = closingPrice - openingPrice
            
            ' Calculate percent change
            If openingPrice <> 0 Then
                percentChange = (yearlyChange / openingPrice) * 100
            Else
                ' Avoiding division by zero
                percentChange = 0
            End If
            
            ' Calculate total volume
            totalVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7))) 
            
            ' Output results
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = yearlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = totalVolume
            
            ' Update greatest % increase, decrease, and volume
            If isFirstStock Then
                greatestIncreaseTicker = ticker
                greatestDecreaseTicker = ticker
                greatestVolumeTicker = ticker
                isFirstStock = False
            Else
                If percentChange > greatestIncreasePercent Then
                    greatestIncreasePercent = percentChange
                    greatestIncreaseTicker = ticker
                End If
                
                If percentChange < greatestDecreasePercent Then
                    greatestDecreasePercent = percentChange
                    greatestDecreaseTicker = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
            End If
            
            ' Move to the next row for summary
            summaryRow = summaryRow + 1
        Loop
        
    Next ws
    
    ' Output greatest % increase, decrease, and volume
    Dim wsSummary As Worksheet
    Set wsSummary = ThisWorkbook.Worksheets.Add
    wsSummary.Name = "Summary"
    wsSummary.Cells(1, 1).Value = "Greatest % Increase"
    wsSummary.Cells(2, 1).Value = "Ticker"
    wsSummary.Cells(2, 2).Value = "Value"
    wsSummary.Cells(3, 1).Value = greatestIncreaseTicker
    wsSummary.Cells(3, 2).Value = greatestIncreasePercent & "%"
    
    wsSummary.Cells(5, 1).Value = "Greatest % Decrease"
    wsSummary.Cells(6, 1).Value = "Ticker"
    wsSummary.Cells(6, 2).Value = "Value"
    wsSummary.Cells(7, 1).Value = greatestDecreaseTicker
    wsSummary.Cells(7, 2).Value = greatestDecreasePercent & "%"
    
    wsSummary.Cells(9, 1).Value = "Greatest Total Volume"
    wsSummary.Cells(10, 1).Value = "Ticker"
    wsSummary.Cells(10, 2).Value = "Value"
    wsSummary.Cells(11, 1).Value = greatestVolumeTicker
    wsSummary.Cells(11, 2).Value = greatestVolume

End Sub
