Sub ticker()

' Define variables
Dim ticker As String
Dim open_price, close_price As Double
Dim percent_change As Double

Dim volume, result_row, last_row As Integer
Dim yearly_change As Double

' Define variables for "greatest" totals
Dim greatest_increase As Double
Dim greatest_increase_ticker As String
Dim greatest_decrease As Double
Dim greatest_decrease_ticker As String
Dim greatest_total As Double
Dim greatest_total_ticker As String



For Each ws In Worksheets

    ' Initialize variables
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    open_price = ws.Cells(2, 3)
    result_row = 2
    greatest_increase = 0
    greatest_decrease = 999
    greatest_total = 0

    ' Loop through the rows
    For i = 2 To LastRow

    
        volume = volume + ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
    
    
            ' Check whether the next row is a new ticker
            If (ticker <> ws.Cells(i + 1, 1).Value) Then
                ' We're on a new ticker.  Print the simple results
                ws.Cells(result_row, 9).Value = ticker
                ws.Cells(result_row, 12).Value = volume
                
                ' Update the close_price for use in calculating the yearly change
                close_price = ws.Cells(i, 6)
                yearly_change = close_price - open_price
                        
                        
                    ' Check whether close or open price are zero before dividing
                    If (close_price = 0 Or open_price = 0) Then
                        
                        ' We have a zero.  Set percent_change to 0
                        percent_change = 0
                        ws.Cells(result_row, 10).Value = yearly_change
                        ws.Cells(result_row, 11).Value = FormatPercent(percent_change, 12)
                        
                        
                        
                        ' Neither of the values are 0.  Divide as normal
                        Else
                
                           
                            ' No zeros.  Calculate and print the close price
                            percent_change = Round((close_price / open_price), 2)
                            ws.Cells(result_row, 10).Value = yearly_change
                            ws.Cells(result_row, 11).Value = FormatPercent(percent_change, 12)
                            
                            
                            
                    End If
        
                ' Update values for next ticker cycle
                result_row = result_row + 1
                open_price = ws.Cells(i + 1, 3)
                
                ' -----------------------------------------------------------------------------
                ' Check for the "greatest" totals

                ' Greatest % Increase
                If (percent_change > greatest_increase) Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = ticker
            
                End If
            
                ' Greatest % Decrease
                If (percent_change < greatest_decrease) Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ticker
            
                End If
            
                ' Greatest Total Volume
                If (volume > greatest_total) Then
                    greatest_total = volume
                    greatest_total_ticker = ticker
            
                End If
                
            ' Reset volume for next ticker cycle
            volume = 0
            
            End If
            

            

    Next i
    
    ' Set the "Greatest" totals
    ws.Cells(2, 16) = FormatPercent(greatest_increase, 12)
    ws.Cells(3, 16) = FormatPercent(greatest_decrease, 12)
    ws.Cells(4, 16) = greatest_total
    ws.Cells(2, 15) = greatest_increase_ticker
    ws.Cells(3, 15) = greatest_decrease_ticker
    ws.Cells(4, 15) = greatest_total_ticker

Next ws


End Sub


