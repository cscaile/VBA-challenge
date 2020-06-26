Attribute VB_Name = "Module1"
Sub Stock_ticker()
For Each ws In Worksheets

    Dim ticker As String
    Dim volume As Double
    volume = 0
    Dim table As Integer
    table = 1
    Dim year_start As Double
    Dim year_end As Double
    Dim change As Double
    Dim percent As Double
    Dim last_row As Integer
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume As Double
    Dim rng1 As Range
    Dim rng2 As Range
    Dim inc_ticker As String
    Dim dec_ticker As String
    Dim vol_ticker As String
    
ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change $"
ws.Range("k1").Value = "% Change"
ws.Range("l1").Value = "Total Volume"
ws.Range("o2").Value = "Greatest % Increase"
ws.Range("o3").Value = "Greatest % Decrease"
ws.Range("o4").Value = "Greates Total Volume"
ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Value"



    For i = 2 To ws.Range("a1").CurrentRegion.End(xlDown).Row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        volume = volume + ws.Cells(i, 7)
        ws.Range("I" & table).Value = ticker
        ws.Range("L" & table).Value = volume
        volume = 0
        Else
        volume = volume + ws.Cells(i + 1, 7).Value
        End If
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        year_start = ws.Cells(i, 3).Value
        table = table + 1
         Else
        year_end = ws.Cells(i, 6).Value
         End If
         If year_start = 0 Then
         percent = "0"
         change = "0"
         Else
        change = year_end - year_start
        ws.Range("J" & table).Value = change
        percent = (year_end / year_start - 1)
         ws.Range("K" & table).Value = percent
        End If
        Next i
    last_row = ws.Cells(Rows.Count, 12).End(xlUp).Row
    ws.Range("k2:k" & last_row).NumberFormat = "0.00%"
    ws.Range("j2:j" & last_row).NumberFormat = "$0.00"
    ws.Range("l2:l" & last_row).NumberFormat = "#,##0"
    For j = 2 To last_row
        If ws.Cells(j, 11).Value < 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 3
        Else
        ws.Cells(j, 11).Interior.ColorIndex = "4"
        End If
          
    Next j
    Set rng1 = ws.Range("k2:k" & last_row)
    Set rng2 = ws.Range("l2:l" & last_row)
     max_percent = ws.Application.WorksheetFunction.Max(rng1)
    min_percent = ws.Application.WorksheetFunction.Min(rng1)
    max_volume = ws.Application.WorksheetFunction.Max(rng2)
'    inc_ticker = ws.Application.WorksheetFunction.VLookup(max_percent, rng1, -2, False)
    
    ws.Range("q2").Value = max_percent
    ws.Range("q3").Value = min_percent
    ws.Range("q4").Value = max_volume
    ws.Range("q2:q3").NumberFormat = "0.00%"
     ws.Range("q4").NumberFormat = "#,##0"
     ws.Columns("A:q").AutoFit
 Next ws
 
 
End Sub



