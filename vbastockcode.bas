Attribute VB_Name = "revised"
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
    ' Needs to be equivilent to zero, so it is a placeholder like with stock volume
    Dim Closer As Double
    ' Needs to be equivilent to zero, so it is a placeholder like with stock volume
    Dim YearlyChange As Double
    ' Needs to be equivilent to zero, so it is a placeholder like with stock volume
    Dim PercentageChange As Double
    ' Needs to be equivilent to zero, so it is a placeholder like with stock volume
    Dim summary_table_row As Integer
    ' Needs to be equivilent to zero, so it is a placeholder like with stock volume
    Dim Column As Integer
    Column = 1
    summary_table_row = summary_table_row + 1
    Dim i As Long
        For i = 2 To LastRow
        'You do need to find the open price, use an if statement
        'The volume needs to be found
        'Find end of stock_ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'set stock ticker name
                stock_ticker = Cells(i, 1).Value
            'set opener
                Opener = Cells(2, 3).Value
                'This value appears to be static, change 2 to i
            'set Closer
                Closer = Cells(i, 6).Value
            'set yearly change
                YearlyChange = Closer - Opener
                Cells(summary_table_row, Column + 1).Value = YearlyChange
            'Set percentage change
            If (Opener = 0 And Closer = 0) Then
                PercentageChange = 0
            ElseIf (Opener = 0 And Closer <> 0) Then
                PercentageChange = 1
            Else
                PercentageChange = YearlyChange / Opener
                Cells(summary_table_row, 11).Value = PercentageChange
                Cells(summary_table_row, 11).NumberFormat = "0.00%"
            End If
            'Add Stock Volume
            stock_volume = stock_volume + Cells(i, 7).Value
            Cells(summary_table_row, 12).Value = stock_volume
           'Add a summary table row
          summary_table_row = summary_table_row + 1
           ' reset Opener
           Opener = Cells(i + 1, 3)
           'reset volumn
           stock_volume = 0
           Else
            stock_volume = stock_volume + Cells(i, 7).Value
            End If
        Next i
        'Identify Last row for YearlyChange
     LRYC = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
        'make colorful
        'Keep the if statement below and put it in the "i" Loop
        For k = 2 To LRYC
                    If (Cells(k, 10).Value > 0 Or Cells(k, 10).Value = 0) Then
                        Cells(k, 10).Interior.ColorIndex = 10
                    ElseIf Cells(k, 10).Value < 0 Then
                        Cells(k, 10).Interior.ColorIndex = 3
                    End If
        Next k
Next ws

End Sub
