# VBA-Challenge
Sub stockmarket()

Dim ws As Worksheet
'For Each ws In Workbooks
'ws.Active

'fill out headers
 Range("I1").EntireColumn.Insert
 Range("I1").Value = "Ticker"
 Range("J1").EntireColumn.Insert
 Range("J1").Value = "yearly_change"
 Range("K1").EntireColumn.Insert
 Range("K1").Value = "percent_change"
 Range("L1").EntireColumn.Insert
 Range("L1").Value = "total _stock_volume"
 Range("Q2").Value = "Greastest % Increase"
 Range("Q3").Value = "Greastest % Decrease"
 Range("Q4").Value = "Greastest Total Volume"
 Range("R1").Value = "Ticker"
 Range("S1").Value = "Value"


 'define variables

 Dim ticker As String
 Dim Summary_Table_Row As Integer
 Dim open_price As Double
 Dim close_price As Double
 Dim number_tickers As Double
 Dim yearly_change As Double
 Dim percnt_change As Double
 Dim total_stock_volume As Double
 Dim i As Double
 Dim Greatest_Percentage_Increase As Double
 Dim Greastest_Percentage_Decrease As Double
 Dim Greatest_Total_Volume As Double
 Dim last_row As Long
 Dim Percentage As Double


 'set initial values
   ticker = 0
   yearly_change = 0
   open_price = 0
   close_price = 0
   Percent_Change = 0
   total_stock_volume = 0
   number_tickers = 0

  last_row = Cells(Rows.Count, 1).End(xlUp).Row
   Summary_Table_Row = 2

   'Loop through all Tickers
   For i = 2 To last_row


  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

  ticker = Cells(i, 1).Value

  Range("I" & Summary_Table_Row).Value = ticker

 total_stock_volume = total_stock_volume + Cells(i, 7).Value

 Range("L" & Summary_Table_Row).Value = total_stock_volume
 total_stock_volume = 0


 Summary_Table_Row = Summary_Table_Row + 1
 Else
 total_stock_volume = total_stock_volume + Cells(i, 7).Value
 End If
 
 If open_price = 0 Then

  open_price = Cells(i, 3).Value

  End If

close_price = Cells(i, 6).Value


'Calculate yearly change
    yearly_change = close_price - open_price
    Range("J" & Summary_Table_Row).Value = yearly_change

    'Calculate Percent Change column
    Percent_Change = FormatPercent((close_price / open_price) - 1)
    Range("K" & Summary_Table_Row).Value = Percent_Change

If yearly_change >= 0 Then
Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
Else
Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
End If


 Next i



 End Sub

