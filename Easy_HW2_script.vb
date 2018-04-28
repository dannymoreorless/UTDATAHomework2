Sub Easy()
Dim lastrow As Long
Dim Total As Double
Dim ii As Integer

For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
    ii = 2
    Total = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
       
            Total = Total + ws.Cells(i, 7).Value
            ws.Cells(ii, 10).Value = Total
            ws.Cells(ii, 9).Value = ws.Cells(i, 1).Value
            Total = 0
            ii = ii + 1
        Else
            Total = Total + ws.Cells(i, 7).Value
    
        End If
    
    Next i
Next ws

End Sub
