Attribute VB_Name = "Module1"
Option Explicit

Sub stock():

'worksheet
Dim ws As Worksheet

'dim everything
Dim i As Double
Dim Lastrow As Double
Dim a As Integer
Dim open_price As Double
Dim Ticker_name As String
Dim total_stock_volume As Double
Dim yearly_change As Double
Dim close_price As Double
Dim percent_change As Double

'dim bonus
Dim max_percent As Double
max_percent = 0
Dim min_percent As Double
min_percent = 0
Dim ticker_max As String
Dim ticker_min As String
Dim total_stock_volume_max As Double
total_stock_volume_max = 0
Dim ticker_gtv As String


'each worksheets
For Each ws In Worksheets

'labels for variables
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'labels for variables bonus
Sheet1.Range("Q1").Value = "Ticker"
Sheet1.Range("R1").Value = "Value"
Sheet1.Range("O2").Value = "Greatest % Increase"
Sheet1.Range("O3").Value = "Greatest % Decrease"
Sheet1.Range("O4").Value = "Greatest Total Volume"

'--

'count all row
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'dim a for loop
a = 1

'open price first ticker
open_price = ws.Cells(2, 3).Value


'dim i
For i = 2 To Lastrow

    
'total stock volume
total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    
'loop
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
'ticker name
        
Ticker_name = ws.Cells(i, 1).Value
a = a + 1
ws.Cells(a, 9).Value = Ticker_name
        
'yearly change
        
close_price = ws.Cells(i, 6).Value
yearly_change = close_price - open_price
ws.Cells(a, 10).Value = yearly_change
        
         
'percent change
If open_price <> 0 Then
percent_change = yearly_change / open_price
ws.Cells(a, 11).Value = percent_change
ws.Cells(a, 11).NumberFormat = "0.00%"
open_price = ws.Cells(i + 1, 3).Value

 
'total_stock_volume
ws.Cells(a, 12).Value = total_stock_volume
total_stock_volume = 0

End If
End If
        
'change color
If ws.Cells(i, 10).Value > 0 Then
ws.Range("J" & i).Interior.ColorIndex = 4
ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Range("J" & i).Interior.ColorIndex = 3

End If

'min max percent change (bonus)
If ws.Cells(i, 11).Value > max_percent Then
max_percent = ws.Cells(i, 11).Value
ticker_max = ws.Cells(i, 9).Value
ElseIf ws.Cells(i, 11).Value < min_percent Then
min_percent = ws.Cells(i, 11).Value
ticker_min = ws.Cells(i, 9).Value

' locate min max to the table
Sheet1.Cells(2, 18).Value = max_percent
Sheet1.Cells(2, 18).NumberFormat = "0.00%"
Sheet1.Cells(2, 17).Value = ticker_max
Sheet1.Cells(3, 18).Value = min_percent
Sheet1.Cells(3, 18).NumberFormat = "0.00%"
Sheet1.Cells(3, 17).Value = ticker_max

End If

'total stock volume max (bonus)
If ws.Cells(i, 12).Value > total_stock_volume_max Then
total_stock_volume_max = ws.Cells(i, 12).Value
ticker_gtv = ws.Cells(i, 9).Value

'locate  max to the table
Sheet1.Cells(4, 18).Value = total_stock_volume_max
Sheet1.Cells(4, 17).Value = ticker_gtv


End If

Next i
a = 0
Next ws
End Sub


