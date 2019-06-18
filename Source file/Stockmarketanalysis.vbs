Sub StockMarketanalyst():

Dim Ws As Worksheet

For Each Ws In Worksheets

Dim i As Long
Dim j As Long
Dim Sum_vol As Double
'Declare Dest_row as the row indicator for the summary table
Dim Dest_row As Long

'Get the last row of the stock data
Dim lRow As Long
lRow = Cells(Rows.Count, 1).End(xlUp).Row


Dim O_price As Double 'O_price is the open price of stock
Dim C_price As Double 'C_price is the closed price of stock

Dim Price_change As Double
O_price = Ws.Cells(2, 3).Value 'At initial, get the open price for the first stock in the stock data table

Sum_vol = 0
Dest_row = 2
'Assign heads of columns in the summary table
Ws.Cells(1, 9).Value = "Ticker"
Ws.Cells(1, 10).Value = "Yearly Change"
Ws.Cells(1, 11).Value = "Percent Change"
Ws.Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To lRow
    If Ws.Cells(i, 1).Value <> Ws.Cells(i + 1, 1).Value Then

    'Get the closed price for the sticker
    C_price = Ws.Cells(i, 6).Value
    
    'Calculate the Yearly Change and update the format and color
    Price_change = CDec(C_price - O_price) 'Get the better precise floating point with Decimal data type
    Ws.Cells(Dest_row, 10).Value = Price_change
    Ws.Cells(Dest_row, 10).NumberFormat = "0.00000000"
    
    'Conditional Formatting the color based on the value of cell
        If Ws.Cells(Dest_row, 10).Value > 0 Then
        Ws.Cells(Dest_row, 10).Interior.ColorIndex = 4
        Else
        Ws.Cells(Dest_row, 10).Interior.ColorIndex = 3
        End If
    
    'Avoid the O_price equal to 0, we will run the test and assign change value to be equal to 100% if open price is 0
        If O_price <> 0 Then
        Ws.Cells(Dest_row, 11).Value = Price_change / O_price
        Else
        Ws.Cells(Dest_row, 11).Value = 1
        End If
    'Update the format for Percent Change
    Ws.Cells(Dest_row, 11).NumberFormat = "0.00%"

    'Fill the sticker name and total volume to summary table
    Ws.Cells(Dest_row, 9).Value = Ws.Cells(i, 1).Value
    Ws.Cells(Dest_row, 12).Value = Ws.Cells(i, 7).Value + Sum_vol

    'After update the result, increase the summary table row 1 unit and reset the volume counter to 0
    Dest_row = Dest_row + 1
    Sum_vol = 0
    'Get the open price for the next sticker
    O_price = Ws.Cells(i + 1, 3).Value

    Else
    'Add the volume to the counter
    Sum_vol = Sum_vol + Ws.Cells(i, 7).Value
    End If
Next i
'Autofit the colunmwidth
Ws.Columns("I:L").AutoFit

'Dealing with the next filtered table
'Get the row number of summary table
Dim lRow_ST As Long
lRow_ST = Cells(Rows.Count, 9).End(xlUp).Row

Dim Great_value As Single
Dim Great_ticker As String

Dim Small_value As Single
Dim Small_ticker As String

Great_value = 0
Small_value = 0

Dim Max_vol As Double
Dim Max_ticker As String

Max_vol = 0

'Assign the labels for the filtered table
Ws.Cells(1, 16).Value = "Ticker"
Ws.Cells(1, 17).Value = "Value"
Ws.Cells(2, 15).Value = "Greatest % Increase"
Ws.Cells(3, 15).Value = "Greatest % Decrease"
Ws.Cells(4, 15).Value = "Greatest Total Volume"

'Find the max value and min value
For j = 2 To lRow_ST
    If Ws.Cells(j, 11).Value < Small_value Then
    Small_value = Ws.Cells(j, 11).Value
    Small_ticker = Ws.Cells(j, 9).Value
    
    ElseIf Ws.Cells(j, 11).Value > Great_value Then
    Great_value = Ws.Cells(j, 11).Value
    Great_ticker = Ws.Cells(j, 9).Value
    End If
    'Find the max volumn value
    If Ws.Cells(j, 12).Value > Max_vol Then
    Max_vol = Ws.Cells(j, 12).Value
    Max_ticker = Ws.Cells(j, 9).Value
    End If
Next j

'Fill the results to the according cells
Ws.Cells(2, 16).Value = Great_ticker
Ws.Cells(2, 17).Value = Great_value
Ws.Cells(2, 17).NumberFormat = "0.00%"
Ws.Cells(3, 16).Value = Small_ticker
Ws.Cells(3, 17).Value = Small_value
Ws.Cells(3, 17).NumberFormat = "0.00%"
Ws.Cells(4, 16).Value = Max_ticker
Ws.Cells(4, 17).Value = Max_vol

'Autofit the columnwidth
Ws.Columns("O:Q").AutoFit
    
Next Ws

End Sub



