Attribute VB_Name = "VBAc1"
Sub VBAstocks()

'Connect all sheets
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    'Find last row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Column Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'Establish Variables
        Dim stock_ticker As String
        Dim stock_volume As Double
        stock_volume = 0
        Dim Opener As Double
        Dim Closer As Double
        Dim YearlyChange As Double
        Dim PercentageChange As Double
        Dim summary_table_row As Double
        summary_table_row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long

 'set openerprice
     Opener = Cells(2, Column + 2).Value
        
        
    For i = 2 To LastRow
        'Find if we are still in the stock ticker year, if not then...
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        
            'set stock ticker name
                stock_ticker = Cells(i, Column).Value
                Cells(summary_table_row, Column + 8).Value = stock_ticker
            'set Closer
                Closer = Cells(i, Column + 5).Value
            'set yearly change
                YearlyChange = Closer - Opener
                Cells(summary_table_row, Column + 9).Value = YearlyChange
            'Set percentage change
            If (Opener = 0 And Closer = 0) Then
                PercentageChange = 0
            ElseIf (Opener = 0 And Closer <> 0) Then
                PercentageChange = 1
            Else
                PercentageChange = YearlyChange / Opener
                Cells(summary_table_row, Column + 10).Value = PercentageChange
                Cells(summary_table_row, Column + 10).NumberFormat = "0.00%"
            End If
            
            'Add Total Stock Volume
            stock_volume = stock_volume + Cells(i, Column + 6).Value
            Cells(summary_table_row, Column + 11).Value = stock_volume
           'Add a summary table row
           summary_table_row = summary_table_row + 1
           ' reset Opener
           Opener = Cells(i + 1, Column + 2)
           'reset volumn
           stock_volume = 0
           Else
            stock_volume = stock_volume + Cells(i, Column + 6).Value
            End If
        Next i
    
            
        'Identify Last row for YearlyChange
            LRYC = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
                
        'make colorful
        For k = 2 To LRYC
                    If (Cells(k, Column + 9).Value > 0 Or Cells(k, 10).Value = 0) Then
                        Cells(k, Column + 9).Interior.ColorIndex = 10
                    ElseIf Cells(k, Column + 10).Value < 0 Then
                    Cells(k, Column + 9).Interior.ColorIndex = 3
                    End If
        Next k

        'Set Metrics Greatest %increase, decrease, and Total Volume
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"

        'Find and plot data points
        For N = 2 To LRYC
            If Cells(N, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LRYC)) Then
                Cells(2, Column + 15).Value = Cells(N, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(N, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(N, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LRYC)) Then
                Cells(3, Column + 15).Value = Cells(N, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(N, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(N, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LRYC)) Then
                Cells(4, Column + 15).Value = Cells(N, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(N, Column + 11).Value
            End If

        Next N

    Next ws

End Sub




