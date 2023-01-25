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
sheet.Range("I1").Value = "Ticker Symbol"
sheet.Range("J1").Value = "Yearly Change ($)"
sheet.Range("K1").Value = "Percent Change"
sheet.Range("L1").Value = "Total Stock Volume"
total_vol = 0
sheet.Range("O2").Value = "Greatest % Increase"
sheet.Range("O3").Value = "Greatest % Decrease"
sheet.Range("O4").Value = "Greatest Total Volume"
sheet.Range("P1").Value = "Ticker"
sheet.Range("Q1").Value = "Value"


'Last_row Counts
Dim last_row As Double
last_row = sheet.Cells(Rows.Count, 1).End(xlUp).row
'MsgBox (last_row)

'Loop through rows in the column
row = 2
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
sheet.Cells(row, 9).Value = ticker
sheet.Cells(row, 10).Value = Format(price_change, "0.00")
sheet.Cells(row, 11).Value = Format(percent_change, "0.00%")
sheet.Cells(row, 12).Value = total_vol

'Change color of the cells based on the value of the price_change
'Green if price_change is positive
If sheet.Cells(row, 10).Value > 0 Then
sheet.Cells(row, 10).Interior.ColorIndex = 4

'Red if the price_change is not positive
Else
sheet.Cells(row, 10).Interior.ColorIndex = 3
End If

total_vol = 0
open_price = 0
row = row + 1

Else
total_vol = total_vol + sheet.Cells(i, 7).Value

End If

Next i
Dim last_ticker As Double
last_ticker = sheet.Cells(Rows.Count, 9).End(xlUp).row
'MsgBox (last_ticker)
greatest_increase = 0
greatest_decrease = 0
For j = 2 To last_ticker
If sheet.Cells(j, 11).Value > greatest_increase Then
greatest_increase = sheet.Cells(j, 11).Value
sheet.Range("P2").Value = sheet.Cells(j, 9).Value
sheet.Range("Q2").Value = Format(greatest_increase, "0.00%")
End If
Next j
For k = 2 To last_ticker
If sheet.Cells(k, 11) < greatest_decrease Then
greatest_decrease = sheet.Cells(k, 11).Value
sheet.Range("P3").Value = sheet.Cells(k, 9).Value
sheet.Range("Q3").Value = Format(greatest_decrease, "0.00%")
End If
Next k
'MsgBox (greatest_decrease)
greatest_vol = 0
For m = 2 To last_ticker

If sheet.Cells(m, 12) > greatest_vol Then
greatest_vol = sheet.Cells(m, 12).Value
sheet.Range("P4").Value = sheet.Cells(m, 9).Value
sheet.Range("Q4").Value = greatest_vol
End If
Next m
'MsgBox (greatest_vol)
Next sheet
End Sub
