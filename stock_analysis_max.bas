Attribute VB_Name = "Module4"
Sub stock_analysis_max():

'Loop through every Worksheet
Dim ws As Worksheet

For Each ws In Worksheets

    'Column headers
    Dim ticker As String
    Dim opening, closing, price_change, percent_change, total_volume As Double

    'Start with the opening row of the first ticker
    Dim first_t_row As Long
    first_t_row = 2
    opening = ws.Cells(2, 3).Value
    total_volume = 0

    'Set up results table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Dim results_row As Long
    results_row = 2
    
    'Loop through every row
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To last_row
    
        'Calculations start when the last row of a ticker symbol is found
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
    
            ticker = ws.Cells(i, 1).Value
            closing = ws.Cells(i, 6).Value
            price_change = closing - opening
        
            'If opening starts at 0, prevent division by 0
            If opening = 0 Then
                opening = 1
            End If
        
            percent_change = price_change / opening
        
            'Accumulate total stock volume just for this ticker symbol
            For j = first_t_row To i
                total_volume = total_volume + ws.Cells(j, 7).Value
            Next j
        
            'List results
            ws.Cells(results_row, 9).Value = ticker
            ws.Cells(results_row, 10).Value = price_change
            ws.Cells(results_row, 11).Value = percent_change
            ws.Cells(results_row, 12).Value = total_volume
        
            'Formatting in results table
            If price_change > 0 Then
                ws.Cells(results_row, 10).Interior.ColorIndex = 4
            ElseIf price_change < 0 Then
                ws.Cells(results_row, 10).Interior.ColorIndex = 3
            End If
            
            ws.Cells(results_row, 11).NumberFormat = "0.00%"
        
            'Reset for next ticker symbol
            first_t_row = i + 1
            opening = ws.Cells(i + 1, 3).Value
            total_volume = 0
            results_row = results_row + 1
    
        End If
                
    Next i
    
    
    'Set up second results table
    Dim results_last_row As Long
    results_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Dim percent_increase_v, percent_decrease_v, max_volume_v As Double
    Dim percent_increase_r, percent_decrease_r, max_volume_r As Double
    Dim percent_increase_t, percent_decrease_t, max_volume_t As String
        
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
            
    'Calculate the values first
    percent_increase_v = WorksheetFunction.max(ws.Range("K2:K" & results_last_row))
    percent_decrease_v = WorksheetFunction.Min(ws.Range("K2:K" & results_last_row))
    max_volume_v = WorksheetFunction.max(ws.Range("L2:L" & results_last_row))
        
    'Use values to find row with matching ticker symbol
    percent_increase_r = WorksheetFunction.Match(percent_increase_v, ws.Range("K2:K" & results_last_row), 0)
    percent_decrease_r = WorksheetFunction.Match(percent_decrease_v, ws.Range("K2:K" & results_last_row), 0)
    max_volume_r = WorksheetFunction.Match(max_volume_v, ws.Range("L2:L" & results_last_row), 0)
        
    'Use row to find matching ticker symbol
    percent_increase_t = ws.Range("I" & percent_increase_r + 1).Value
    percent_decrease_t = ws.Range("I" & percent_decrease_r + 1).Value
    max_volume_t = ws.Range("I" & max_volume_r + 1).Value
        
    'Format to %
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
        
    'Fill in results
    ws.Range("P2").Value = percent_increase_t
    ws.Range("Q2").Value = percent_increase_v
    ws.Range("P3").Value = percent_decrease_t
    ws.Range("Q3").Value = percent_decrease_v
    ws.Range("P4").Value = max_volume_t
    ws.Range("Q4").Value = max_volume_v

        
Next ws


End Sub


