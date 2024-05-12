Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim summaryRow As Integer
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        summaryRow = 2
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        
        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Assume data starts at row 2
        openingPrice = ws.Cells(2, 3).Value
        totalVolume = 0
        
        ' Initialize summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "% Change"
        ws.Range("L1").Value = "Total Volume"
        
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if it's still the same stock
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closingPrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                quarterlyChange = closingPrice - openingPrice
                percentChange = (quarterlyChange / openingPrice) * 100
                
                ' Print the Ticker, Changes and Volume in the Summary Table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Conditional Formatting
                With ws.Cells(summaryRow, 10)
                    .NumberFormat = "0.00"
                    If quarterlyChange > 0 Then
                        .Interior.Color = vbGreen
                    Else
                        .Interior.Color = vbRed
                    End If
                End With
                
                ' Update max and min changes
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If
                
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
                
                ' Reset for the next stock
                openingPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
                summaryRow = summaryRow + 1
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Print the results for the Greatest Increase, Decrease, and Volume
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(3, 15).Value = maxDecreaseTicker
        ws.Cells(4, 15).Value = maxVolumeTicker
        ws.Cells(2, 16).Value = maxIncrease
        ws.Cells(3, 16).Value = maxDecrease
        ws.Cells(4, 16).Value = maxVolume
    Next ws
End Sub
