Sub VBAHomework()
    Dim Ticker_Name As String
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
    
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            MsgBox (ws.Cells(i, 1).Value)
        End If
    Next i
    
End Sub


Sub VBAHomework2()
    
    For Each ws In Worksheets
    
    Dim Ticker_Name As String
    Dim Total_Volume As Double
    Total_Volume = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total"
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
    
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            Ticker_Name = ws.Cells(i, 1).Value
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            ws.Range("J" & Summary_Table_Row).Value = Total_Volume
            Summary_Table_Row = Summary_Table_Row + 1
            Total_Volume = 0
        Else
            
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        End If
    Next i
    
    Next ws
        
End Sub


