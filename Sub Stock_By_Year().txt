Sub Stock_By_Year()

'Loop through all sheets
    For Each ws In Worksheets
    
'Set inicial variable

Dim Ticker_Symbol As String
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim j As Integer
Dim Open_price As Double
Dim Close_price As Double
Dim Total_Volume As Double


j = 2
Total_Volume = 0

' Initial Open Price
Open_price = ws.Cells(2, 3).Value

'Find last row
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all tickers

        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 
'Output ticker symbols
 
Ticker_Symbol = ws.Cells(i, 1).Value
ws.Cells(j, 10).Value = Ticker_Symbol
  
' Output yearly change

Close_price = ws.Cells(i, 6).Value
Yearly_Change = Close_price - Open_price
ws.Cells(j, 11).Value = Yearly_Change

' Output Percentage change

Percentage_Change = Yearly_Change / Open_price
ws.Cells(j, 12).Value = Percentage_Change
ws.Cells(j, 12).NumberFormat = "0.00%"

' Output total stock volume

' Find last row Yearly Change
Dim YChangelastrow As Long
YChangelastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

'' Define colors
For j = 2 To YChangelastrow
 
 ' Set positive change in green

 If ws.Cells(j, 11).Value >= 0 Then
  ws.Cells(j, 11).Interior.ColorIndex = 10
  
' Set negative change in red

 ElseIf ws.Cells(j, 11).Value < 0 Then
  ws.Cells(j, 11).Interior.ColorIndex = 3

End If
Next j

' Setting Greatest % (increase, decrease and total volume)

' Add titles to the Column Header

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker Symbol"
ws.Cells(1, 17).Value = "Value"

'updating opening price

Open_price = ws.Cells(i + 1, 3).Value

 ' Add titles to the Column Header
 
        ws.Cells(1, 10).Value = "Ticker Symbol"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"

' Autofit display data

        ws.Columns("J:Z").AutoFit
        
End If

Next i

    Next ws
        
End Sub
