Sub StockData()

' ------  Defining variables ------
Dim ws As Worksheet
Dim ticker As String
Dim number_tickers As Integer
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim lastRowState As Long

Dim greatest_increase As Double
Dim greatest_increase_ticker As String
Dim greatest_decrease As Double
Dim greatest_decrease_ticker As String
Dim greatest_stockvolume As Double
Dim greatest_stockvolume_ticker As String


' ------ Loop through Sheets
For Each ws In Worksheets
    lastRowState = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

'-------- Create headers for new columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'-------- Initialize variables
    number_tickers = 0
    opening_price = 0
    total_stock_volume = 0

    For i = 2 To lastRowState
        ticker = ws.Cells(i, 1).Value

        If opening_price = 0 Then
           opening_price = ws.Cells(i, 3).Value
        End If

        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
 
        If ws.Cells(i + 1, 1).Value <> ticker Then
            number_tickers = number_tickers + 1
            ws.Cells(number_tickers + 1, 9).Value = ticker

            closing_price = ws.Cells(i, 6).Value
            yearly_change = closing_price - opening_price
            ws.Cells(number_tickers + 1, 10).Value = yearly_change

'-------- Color Formatting
            If yearly_change > 0 Then
                 ws.Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                    ws.Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
        End If

            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / opening_price
            End If

 '-------- Formatting for percent
    ws.Cells(number_tickers + 1, 11).Value = percent_change
    ws.Cells(number_tickers + 1, 11).NumberFormat = "0.00%"

'-------- Store total stock volume
    ws.Cells(number_tickers + 1, 12).Value = total_stock_volume

 '-------- Reset variables for next ticker
                opening_price = 0
                total_stock_volume = 0
        End If
    Next i

'-------- Bonus
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        lastRowState = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

        ' Initialize bonus portion variables
        greatest_increase = -1
        greatest_decrease = 1
        greatest_stockvolume = 0

'-------- Loop to find greatest increase, decrease and volume
        For i = 2 To lastRowState
            If ws.Cells(i, 11).Value > greatest_increase Then
                greatest_increase = ws.Cells(i, 11).Value
                greatest_increase_ticker = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 11).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(i, 11).Value
                greatest_decrease_ticker = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 12).Value > greatest_stockvolume Then
                greatest_stockvolume = ws.Cells(i, 12).Value
                greatest_stockvolume_ticker = ws.Cells(i, 9).Value
            End If
        Next i

'-------- Results
        ws.Range("P2").Value = greatest_increase_ticker
        ws.Range("Q2").Value = greatest_increase
        ws.Range("Q2").NumberFormat = "0.00%"

        ws.Range("P3").Value = greatest_decrease_ticker
        ws.Range("Q3").Value = greatest_decrease
        ws.Range("Q3").NumberFormat = "0.00%"

        ws.Range("P4").Value = greatest_stockvolume_ticker
        ws.Range("Q4").Value = greatest_stockvolume
    Next ws

End Sub
