Sub TickerSummary()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    
        Dim ticker As String
        
        Dim tickerCount As Long
        tickerCount = 0
        
        Dim openPrice As Double
        Dim closePrice As Double
        Dim yearlyChange As Double
        
        Dim percentChange As Double
        Dim totalStockVol As Double
        totalStockVol = 0
                
        Dim summaryTableRow As Double
        summaryTableRow = 2
        
        ' Create Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
              
        ' Use in yearlyChange calculation
        openPrice = Cells(2, 3).Value
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all stock tickers
            For i = 2 To LastRow
            ' finds row ticker value changes
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                ' Add to the totalStockVol
                totalStockVol = totalStockVol + CLng(ws.Cells(i, 7).Value)
                                
            Else
                ' Set the ticker name
                ticker = ws.Cells(i, 1).Value
                ' Print the ticker in the Summary Table
                ws.Range("I" & summaryTableRow).Value = ticker
                
                ' ticker yearlyChange calc, openingPrice defined above
                closePrice = ws.Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                ' Print the Yearly Change in the Summary Table
                ws.Range("J" & summaryTableRow).Value = yearlyChange
                If yearlyChange < 0 Then
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                    End If
                
                'Calculate Percent Change
                If openPrice = 0 Then
                    percentChage = 0
                Else
                    percentChange = (closePrice - openPrice) / openPrice
                    ws.Range("K" & summaryTableRow).Value = percentChange
                End If
                ' Print the Total Volume in the Summary Table
                ws.Range("K" & summaryTableRow).Value = percentChange
                ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
                
                ' Calc Total Stock Volume
                totalStockVol = totalStockVol + CLng(ws.Cells(i, 7).Value)
                ' Print the Total Volume in the Summary Table
                ws.Range("L" & summaryTableRow).Value = totalStockVol
                
                ' Reset the Stock Volume Total & go to the next summary row
                totalStockVol = 0
                summaryTableRow = summaryTableRow + 1
                
                'Update Opening Price or next ticker symbol
                openPrice = ws.Cells(i + 1, 3)
                                    
            End If
            
            Next i
                              
        Next ws
            
        End Sub