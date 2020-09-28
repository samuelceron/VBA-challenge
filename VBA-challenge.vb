Sub forEachWs()
    ' Set variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim ticker_open As Double
    Dim ticker_close As Double
    Dim ticker_change As Double
    Dim ticker_porcentualChange As Double
    Dim ticker_name_increase As String
    Dim ticker_name_decrease As String
    Dim ticker_name_stock As String
    Dim ticker_value_increase As Double
    Dim ticker_value_decrease As Double
    Dim ticker_value_stock As Double
    Dim Summary_Table_Row As Integer
    
    ' Begin the worksheets loop.
    For Each ws In ThisWorkbook.Worksheets
        ticker_value_increase = 0
        ticker_value_decrease = 0
        ticker_value_stock = 0
        ticker_total = 0
        ticker_open = ws.Cells(2, 3).Value
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row     
        Summary_Table_Row = 2
        ' Set headers
        ws.Cells(1, 9).Value = ("Ticker")
        ws.Cells(1, 10).Value = ("Yearly Change")
        ws.Cells(1, 11).Value = ("Percent Char")
        ws.Cells(1, 12).Value = ("Total Stock Value")
        ws.Cells(1, 16).Value = ("Ticker")
        ws.Cells(1, 17).Value = ("Value")
        ws.Cells(2, 15).Value = ("Greatest % Increase")
        ws.Cells(3, 15).Value = ("Greatest % Decrease")
        ws.Cells(4, 15).Value = ("Greatest Total Volume")
        
        ' Loop through all tickers
        For i = 2 To LastRow
            ticker_total = ticker_total + ws.Cells(i, 7)
            ' Check if we are still within the same ticker if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Calculate ticker change
                ticker_cierre = ws.Cells(i, 6).Value
                ticker_change = ticker_cierre - ticker_open
                'Calculate ticker porcentual change
                If (ticker_open = 0) Then
                    ticker_porcentualChange = 0
                Else
                    ticker_porcentualChange = ((ticker_cierre - ticker_open) / ticker_open)
                End If
                ' Set new ticker open
                ticker_open = ws.Cells(i + 1, 3).Value
                ' Set ticker_name and ticker_close
                ticker_name = ws.Cells(i, 1).Value
                ticker_close = ws.Cells(i, 3).Value

                ' Print the ticker in the Summary table
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
                ws.Range("J" & Summary_Table_Row).Value = ticker_change
                ws.Range("L" & Summary_Table_Row).Value = ticker_total
                ws.Range("K" & Summary_Table_Row).Value = ticker_porcentualChange
                ws.Range("K" & Summary_Table_Row).Style = "Percent"
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                ' Formatting conditional change cells
                If ((ws.Range("J" & Summary_Table_Row).Value) <= 0) Then
                    With ws.Range("J" & Summary_Table_Row).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                ElseIf ((ws.Range("J" & Summary_Table_Row).Value) > 0) Then
                    With ws.Range("J" & Summary_Table_Row).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 5287936
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End If
                Summary_Table_Row = Summary_Table_Row + 1
                ticker_total = 0                                              
            End If
        Next i
        LastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        ' Loop through summary table
        For j = 2 To LastSummaryRow
            'Increase
            If ws.Cells(j, 11).Value > 0 Then
                If ws.Cells(j, 11).Value > ticker_value_increase Then
                    ticker_value_increase = ws.Cells(j, 11).Value
                    ticker_name_increase = ws.Cells(j, 9).Value
                End If
            'Decrease
            ElseIf (ws.Cells(j, 11).Value < 0) Then
                If ws.Cells(j, 11).Value < ticker_value_decrease Then
                    ticker_value_decrease = ws.Cells(j, 11).Value
                    ticker_name_decrease = ws.Cells(j, 9).Value
                End If
            End If
            'Greatest Total stock
            If ws.Cells(j, 12).Value > ticker_value_stock Then
                    ticker_value_stock = ws.Cells(j, 12).Value
                    ticker_name_stock = ws.Cells(j, 9).Value
            End If
        Next j
        
        ' Print Challenge results
        ws.Cells(2, 16).Value = ticker_name_increase
        ws.Cells(2, 17).Value = ticker_value_increase
        ws.Cells(3, 16).Value = ticker_name_decrease
        ws.Cells(3, 17).Value = ticker_value_decrease
        ws.Cells(4, 16).Value = ticker_name_stock
        ws.Cells(4, 17).Value = ticker_value_stock
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Columns("J:Q").EntireColumn.AutoFit
    Next
    
End Sub