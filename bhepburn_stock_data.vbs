Sub stock_data():
    For Each ws In Worksheets
    
        'Initialize variables
        Dim ticker, giTicker, gdTicker, gvTicker, strDate As String
        Dim yearOpen, yearClose, vol, change, pChange, GPI, GPD, GTV As Double
        Dim TotalDataRow As Integer
        
        vol = 0
        TotalDataRow = 2
        GPI = 0
        GPD = 0
        GTV = 0
        
        'Determine Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Generate new column titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Columns("A:Q").AutoFit
        ws.Range("K2:K" & LastRow).NumberFormat = "#.00%"
        ws.Range("Q2:Q3").NumberFormat = "#.00%"
        
        For i = 2 To LastRow
            strDate = ws.Cells(i, 2).Value
            If InStr(5, strDate, "0102") <> 0 Then
                yearOpen = ws.Cells(i, 3).Value
            End If
            If InStr(strDate, "1231") <> 0 Then
                yearClose = ws.Cells(i, 6).Value
            End If
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                change = yearClose - yearOpen
                pChange = change / yearOpen
                ws.Cells(TotalDataRow, 9).Value = ticker
                ws.Cells(TotalDataRow, 10).Value = change
                
                If change < 0 Then
                    ws.Cells(TotalDataRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(TotalDataRow, 10).Interior.ColorIndex = 4
                End If
                
                ws.Cells(TotalDataRow, 11).Value = pChange
                ws.Cells(TotalDataRow, 12).Value = vol
                
                If pChange > GPI Then
                    GPI = pChange
                    giTicker = ticker
                End If
                
                If pChange < GPD Then
                    GPD = pChange
                    gdTicker = ticker
                End If
                
                If vol > GTV Then
                    GTV = vol
                    gvTicker = ticker
                End If
                
                TotalDataRow = TotalDataRow + 1
                yearOpen = 0
                yearClose = 0
                change = 0
                pChange = 0
                vol = 0
            Else
                vol = vol + ws.Cells(i, 7).Value
                
            End If
        Next i
        
        ws.Cells(2, 16).Value = giTicker
        ws.Cells(2, 17).Value = GPI
        ws.Cells(3, 16).Value = gdTicker
        ws.Cells(3, 17).Value = GPD
        ws.Cells(4, 16).Value = gvTicker
        ws.Cells(4, 17).Value = GTV
        
        
    Next ws
End Sub
