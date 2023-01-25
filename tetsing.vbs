Sub VBA_challenge():
Dim sheet As Worksheet
'For loop to make sure the code work for multisheets at one run
For Each sheet In Worksheets
'MsgBox (sheet.Name)

Dim ticker As String
Dim total_vol As Double
Dim row As Integer
Dim open_price As Double
Dim close_price As Double

'Row Titles
sheet.Range("I1").Value = "Ticker"
sheet.Range("J1").Value = "Yearly Change"
sheet.Range("K1").Value = "Percent Change"
sheet.Range("L1").Value = "Total Stock Volume"
total_vol = 0

'Last_row Counts
Dim last_row As Double
last_row = sheet.Cells(Rows.Count, 1).End(xlUp).row
'MsgBox (last_row)

'Loop through rows in the column
For i = 2 To last_row
If open_price = 0 Then
open_price = sheet.Cells(i, 3).Value
End If

'Searches for when the value of the next cell is different than that of the current cell
If sheet.Cells(i - 1, 1) = sheet.Cells(i, 1) And sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
close_price = sheet.Cells(i, 6).Value
price_change = close_price - open_price
percent_change = price_change / open_price
ticker = sheet.Cells(i, 1).Value
total_vol = total_vol + sheet.Cells(i, 7).Value
sheet.Cells(i, 9).Value = ticker
sheet.Cells(i, 10).Value = Format(price_change, "0.00")
sheet.Cells(i, 11).Value = Format(percent_change, "0.00%")
sheet.Cells(i, 12).Value = total_vol

'Change color of the cells based on the value of the price_change
'Green if price_change is positive
If sheet.Cells(i, 10).Value > 0 Then
sheet.Cells(i, 10).Interior.ColorIndex = 4

'Red if the price_change is not positive
Else
sheet.Cells(i, 10).Interior.ColorIndex = 3
End If

total_vol = 0
open_price = 0

Else
total_vol = total_vol + sheet.Cells(i, 7).Value

End If

Next i


Next sheet




End Sub




