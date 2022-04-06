Sub stocks()

For Each ws In Worksheets

ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"

Dim ticker As String
Dim summary_table As Integer
Dim name As String

summary_table = 2
volume = 0
name = ws.name

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        volume = volume + ws.Cells(i, 7).Value
        ws.Range("J" & summary_table).Value = ticker
        ws.Range("M" & summary_table).Value = volume
        summary_table = summary_table + 1
        volume = 0

    Else
        volume = volume + ws.Cells(i, 7).Value

    End If

    If Right(ws.Cells(i, 2).Value, 4) = "0102" Then
        start_value = ws.Cells(i, 3).Value

    ElseIf Right(ws.Cells(i, 2).Value, 4) = "1231" Then
        end_value = ws.Cells(i, 6).Value
        ws.Range("K" & summary_table - 1).Value = end_value - start_value
        ws.Range("L" & summary_table - 1).Value = (end_value - start_value) / start_value
        ws.Range("L" & summary_table - 1).NumberFormat = "0.00%"

    End If
    
    If ws.Range("K" & summary_table).Value < 0 Then
        ws.Range("K" & summary_table).Interior.ColorIndex = 3
        
    Else
        ws.Range("K" & summary_table).Interior.ColorIndex = 4
        
    End If

Next i

Next ws

End Sub