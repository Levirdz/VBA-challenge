Sub stock_analysis()

    'Script Replication
    Dim wb As Workbook
    Dim ws As Worksheet
    For Each wb In Workbooks
        For Each ws In wb.Worksheets
            ws.Activate
    
    
            'Variables
            Dim ticker_symbol As String
            Dim ticker_point As Long
            Dim yearly_change As Double
            Dim percent_change As Double
            Dim stock_volume As Double
            Dim open_price As Double
            Dim close_price As Double
            
            'Summary Variables
            Dim greatest_increase_ticker As String
            Dim greatest_decrease_ticker As String
            Dim greatest_volume_ticker As String
            Dim greatest_increase_value As Double
            Dim greatest_decrease_value As Double
            Dim greatest_volume_value As Double

            'Looping Variable
            Dim i As Long
    
            'Variable Values
            i = 2
            ticker_point = 2
            open_price = Cells(2, 3).Value
            greatest_increase_value = 0
            greatest_decrease_value = 0
            greatest_volume_value = 0
            
            'Headers
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
            
            ws.Range("I1:Q1").Font.Bold = True
            ws.Range("O1:O4").Font.Bold = True
            
            'Looping
            Do While Cells(i, 1).Value <> ""
                
                If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                    stock_volume = stock_volume + Cells(i, 7).Value
                    i = i + 1
                Else
                    Cells(ticker_point, 9).Value = Cells(i, 1).Value
                    close_price = Cells(i, 6).Value
                    yearly_change = close_price - open_price
                    Cells(ticker_point, 10).Value = yearly_change
                    
                    'Conditional Formatting
                    If yearly_change < 0 Then
                        Cells(ticker_point, 10).Interior.ColorIndex = 46
                    Else
                        Cells(ticker_point, 10).Interior.ColorIndex = 43
                    End If
                    
                    If Not open_price = 0 Then
                        percent_change = yearly_change / open_price
                    End If
                    
                    Cells(ticker_point, 11).Value = Format(percent_change, "Percent")

                    stock_volume = stock_volume + Cells(i, 7).Value
                    
                    'Summary Table
                    If percent_change > greatest_increase_value Then
                        greatest_increase_value = percent_change
                        greatest_increase_ticker = Cells(i, 1).Value
                        Cells(2, 17).Value = Format(greatest_increase_value, "Percent")
                    End If
                    
                    If percent_change < greatest_decrease_value Then
                        greatest_decrease_value = percent_change
                        greatest_decrease_ticker = Cells(i, 1).Value
                        Cells(3, 17).Value = Format(greatest_decrease_value, "Percent")
                    End If
                    
                    If stock_volume > greatest_volume_value Then
                        greatest_volume_value = stock_volume
                        greatest_volume_ticker = Cells(i, 1).Value
                    End If
                    
                    Cells(ticker_point, 12).Value = stock_volume
                    stock_volume = 0
                        
                    open_price = Cells(i + 1, 3).Value
                    ticker_point = ticker_point + 1
                    i = i + 1
                        
                End If
            Loop
            
                'Summary Results
                Cells(2, 16).Value = greatest_increase_ticker
                Cells(3, 16).Value = greatest_decrease_ticker
                Cells(4, 16).Value = greatest_volume_ticker
                Cells(2, 17).Value = greatest_increase_value
                Cells(3, 17).Value = greatest_decrease_value
                Cells(4, 17).Value = greatest_volume_value
                
                'Fitting New Columns
                ws.Columns("I:Q").AutoFit
            
        Next ws
    Next wb

End Sub