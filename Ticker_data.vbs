Sub StocksLoop()

Dim ws As Worksheet

For Each ws In Worksheets
    
    Dim ticker As String
    Dim vol As Double
    vol = 0
    
    Dim openYearly As Double
    Dim closeYearly As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    Dim summaryRow As Long
    summaryRow = 2
    
    Dim lastRow As Long
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    openYearly = ws.Cells(2, 3).Value
    
    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            closeYearly = ws.Cells(i, 6).Value
            yearlyChange = closeYearly - openYearly
             
            If openYearly <> 0 Then
                percentChange = (yearlyChange / openYearly) * 100
            End If
            
            vol = vol + ws.Cells(i, 7).Value
            
            ws.Range("I" & summaryRow).Value = ticker
            ws.Range("J" & summaryRow).Value = yearlyChange
            
            If (yearlyChange > 0) Then
                ws.Range("J" & summaryRow).Interior.ColorIndex = 4
            ElseIf (yearlyChange < 0) Then
                ws.Range("J" & summaryRow).Interior.ColorIndex = 3
            End If
            
            ws.Range("K" & summaryRow).Value = (CStr(percentChange) & "%")
            ws.Range("L" & summaryRow).Value = vol
            
            summaryRow = summaryRow + 1
            openYearly = ws.Cells(i + 1, 3).Value
            vol = 0
        
        Else
            vol = vol + ws.Cells(i, 7).Value
        End If
        
    Next i
    
Next ws

End Sub
