Sub Calculate_Stock():

' variable for ticker
Dim ticker As String

' variable for number of tickers
Dim number_tickers As Integer

' variable for the last row in each worksheet.
Dim Last_Row As Long

' variable for opening price for specific year
Dim opening_price As Double

' variable for closing price for specific year
Dim closing_price As Double

' variable for yearly change
Dim yearly_change As Double

' variable for percent change
Dim percent_change As Double

' variable for total stock volume
Dim total_stock_volume As Double

' Loop through workbook
For Each ws In Worksheets

    ' Activate Worksheet
    ws.Activate

    ' Find the last row of each worksheet
    Last_Row = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add header columns 
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Zero out Variables
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    ' Loop for Tickers
    For i = 2 To Last_Row

        ' Value of Ticker
        ticker = Cells(i, 1).Value
        
        ' Get values for opening prices
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        ' Getting Total Stock Volume
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        ' Run for next ticker
        If Cells(i + 1, 1).Value <> ticker Then
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            ' Value for Closing Prices
            closing_price = Cells(i, 6)
            
            ' Formula for yearly Change
            yearly_change = closing_price - opening_price
            
            ' Place Yearly Change
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' If yearly change value is greater than 0, shade cell green.
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ' If yearly change value is less than 0, shade cell red.
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            ' If yearly change value is 0, shade cell yellow.
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Calculate percent change value for ticker.
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            ' Format the percent_change value as a percent.
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            ' Set opening price back to 0 when we get to a different ticker in the list.
            opening_price = 0
            
            ' Add total stock volume value to the appropriate cell in each worksheet.
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Set total stock volume back to 0 when we get to a different ticker in the list.
            total_stock_volume = 0
        End If
        
    Next i

Next ws

End Sub
