Attribute VB_Name = "Module1"
Sub Stock_market()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets


'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Create variable to hold stock volume
'Dim stock_volume As Double
'stock_volume = 0

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long


'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0
Dim sum_stock As Double
sum_stock = 0
Dim sum_ticker As Double
sum_ticker = 2
Dim FindRow As Double


'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow

'Calculate stock volume
sum_stock = sum_stock + ws.Cells(i, 7).Value

'Calculate open price
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
open_price = ws.Cells(i, 3).Value
End If


'Ticker symbol output
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
ws.Cells(sum_ticker, 9).Value = Ticker


'Calculate change in Price
close_price = ws.Cells(i, 6).Value
price_change_percent = close_price - open_price
ws.Cells(sum_ticker, 10).Value = price_change_percent


'Colors yearly change
If ws.Cells(sum_ticker, 10).Value > 0 Then
ws.Cells(sum_ticker, 10).Interior.ColorIndex = 4
Else: ws.Cells(sum_ticker, 10).Interior.ColorIndex = 3
End If



'Fixing the open price equal zero problem
If open_price <> 0 Then
price_change_percent = (price_change_percent / open_price)

End If


'outputs
ws.Cells(sum_ticker, 11).Value = price_change_percent
If ws.Cells(sum_ticker, 11).Value > 0 Then
ws.Cells(sum_ticker, 11).Interior.ColorIndex = 4
Else: ws.Cells(sum_ticker, 11).Interior.ColorIndex = 3
End If
ws.Cells(sum_ticker, 11).NumberFormat = "0.00%"
ws.Cells(sum_ticker, 12).Value = sum_stock
sum_stock = 0

sum_ticker = sum_ticker + 1


End If


Next i




'Calculate Greatest % increase
ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("k:k"))
ws.Cells(2, 17).NumberFormat = "0.00%"
'Ticker increase output
FindRow = Application.WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("k:k"), 0)
ws.Cells(2, 16).Value = ws.Cells(FindRow, 9).Value


'Calculate Greatest % decrease
ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("k:k"))
ws.Cells(3, 17).NumberFormat = "0.00%"
'Ticker decrease output
FindRow = Application.WorksheetFunction.Match(ws.Cells(3, 17).Value, ws.Range("k:k"), 0)
ws.Cells(3, 16).Value = ws.Cells(FindRow, 9).Value


'Calculate Greatest total volume
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("l:l"))
'Ticker Greatest total output
FindRow = Application.WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L:L"), 0)
ws.Cells(4, 16).Value = ws.Cells(FindRow, 9).Value



Next ws

End Sub


