Sub Stock_Analysis()

'set variables/dim
        Dim ticker As String
        Dim startRow As Long
        Dim endRow As Long
        Dim yearlyChange As Double
        Dim percentageChange As Double
        Dim totalVolume As Double
       
 'define last row
        For Each ws In Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
 'starting points
totalVolume = 0
startRow = 2
openPrice = ws.Cells(2, "C").Value

'create table headers/customer view
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percentage Change"
    ws.Cells(1, "L").Value = "Total Volume"

'run this loop over data in
For i = 2 To lastRow

'define where
totalVolume = totalVolume + ws.Cells(i, "G").Value
ticker = ws.Cells(i, "A").Value

If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
ws.Cells(startRow, "I").Value = ticker

closePrice = ws.Cells(i, "F").Value
ws.Cells(startRow, "J").Value = closePrice - openPrice

'add colour
If ws.Cells(startRow, "J").Value > 0 Then
ws.Cells(startRow, "J").Interior.ColorIndex = 4

Else
ws.Cells(startRow, "J").Interior.ColorIndex = 3
End If

'<> means not equal to
If openPrice <> 0 Then

' ,2 to two decimal places
ws.Cells(startRow, "K").Value = FormatPercent((closePrice - openPrice) / openPrice, 2)
Else
ws.Cells(startRow, "K").Value = 0
End If

ws.Cells(startRow, "L").Value = totalVolume

startRow = startRow + 1
openPrice = ws.Cells(i + 1, "C").Value
totalVolume = 0
End If

Next i

Next ws
End Sub


