Attribute VB_Name = "Module1"
Sub stocks()

'Set Worksheets
Dim ws As Worksheet
For Each ws In Worksheets

'Set initial variable for holding Ticker
Dim TickerName As String

Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Keep track of the location for each ticker name in the Ticker Counter
Dim TickerTable As Integer
TickerTable = 2

'Set initial variable for opening & closing price
Dim open_price As Double

Dim close_price As Double

Dim Yearly_Change As Double

    'Name the columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Create a script that loops through all the stocks for one year and outputs the following information:
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        
        'First step of the loop
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TickerName = ws.Cells(i, 1).Value
        open_price = ws.Cells(i, 3).Value
        close_price = ws.Cells(i, 6).Value
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        Yearly_Change = close_price - open_price
        percent_change = (Yearly_Change / open_price) * 100
        
        'Print Ticker name in the TickerTable
        ws.Range("I" & TickerTable).Value = TickerName
        'Print Total stock volume in the TickerTable
        ws.Range("L" & TickerTable).Value = Total_Stock_Volume
        'Print Yearly Change in the TickerTable
        ws.Range("J" & TickerTable).Value = Yearly_Change
        'Print Percent Change in the TickerTable
        ws.Range("K" & TickerTable).Value = percent_change
        
        'Add 1 to the TickerTable
        TickerTable = TickerTable + 1
        
        'Reset the Total Stock Volume
        Total_Stock_Volume = 0
        
        End If
    
    Next i
    
    ws.Columns("K").NumberFormat = "0.00%"
    
    'Formatting
    For i = 2 To 91
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i



Next ws
End Sub
