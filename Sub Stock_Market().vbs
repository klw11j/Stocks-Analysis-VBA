SSub Stock_Market()
    
Application.ScreenUpdating = False
    
    'Loop through each worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets

    'Declare Variables
    Dim vol_total As Double
    Dim summary_table_row As Integer
    Dim open_date As Double
    Dim close_date As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim lastrow As Long
        
        'Set values
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        total_vol = 0
        summary_table_row = 2
        open_date = 0
        close_date = 0
        yearly_change = 0
        percent_change = 0
        
        'Create labels
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        For i = 2 To lastrow
            
            'Determine opening value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                open_date = ws.Cells(i, 3).Value
            End If

            'Add to the "Total Stock Value" amount
            total_vol = total_vol + ws.Cells(i, 7)

            'Determine when the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(summary_table_row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(summary_table_row, 12).Value = total_vol

                'Calculate yearly change
                close_date = ws.Cells(i, 6).Value
                yearly_change = close_date - open_date
                ws.Cells(summary_table_row, 10).Value = yearly_change

               'Format coloring
                If yearly_change >= 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                End If
                
        'Find percent change
        If open_date = 0 And close_date = 0 Then
            percent_change = 0
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        ElseIf open_date = 0 Then
        
            'Cannot divide by 0
            Dim new_stock As String
            new_stock = "New Stock"
            ws.Cells(summary_table_row, 11).Value = new_stock
        Else
            percent_change = yearly_change / open_date
            ws.Cells(summary_table_row, 11).Value = percent_change
            ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                summary_table_row = summary_table_row + 1
                vol_total = 0
            End If
            
                date_open = 0
                date_close = 0
                yearly_change = 0
                percent_change = 0
                
            End If
        
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    'Determine the maximum and minimum values
    
  For i = 2 To lastrow
                Dim GPI As Double, GPD As Double, GTV As Double
                
    'Find the maximum percent change
            If ws.Cells(i, 11).Value > GPI Then
                GPI = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                
                End If
                
    'Find the minimum percent change
            If ws.Cells(i, 11).Value < GPD Then
                GPD = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                
                End If
                
    'Find the maximum total volume
            If ws.Cells(i, 12).Value > GTV Then
                GTV = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                
                End If
            
            Next i

    Next ws

End Sub