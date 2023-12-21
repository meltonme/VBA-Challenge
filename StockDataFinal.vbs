Sub StockData():

'Loop through entire workbook
For Each ws In Worksheets

    'Declare Variables
    Dim WorksheetName As String
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim LowPrice As Double
    Dim HighPrice As Double
    Dim Volume As Integer
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim GreatVol As Double
    Dim i As Double
    Dim LastRow As Double
    Dim FirstRow As Double


    'Create Column Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Drcrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"


    'Start the Loop
    'Define the Variables
    FirstRow = 2
    Volume = 0
    OpenPrice = ws.Cells(2, 3)
    ClosePrice = 0
    GreatIncr = 0
    LastRow = ws.Cells(rows.Count, 1).End(xlUp).Row
    'Set up the loop to pull the tickers
    For i = 2 To LastRow
        Total = Total + ws.Cells(i, 7)
        If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            ws.Cells(FirstRow, 12) = Total
            ws.Cells(FirstRow, 9) = ws.Cells(i, 1)
            ClosePrice = ws.Cells(i, 6)
            'Calculate the difference between the Close Price and the Open Price to determine the yearly change
            ws.Cells(FirstRow, 10) = ClosePrice - OpenPrice
            'Create conditional statements to color code the yearly change
            If ws.Cells(FirstRow, 10) > 0 Then
                ws.Cells(FirstRow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(FirstRow, 10).Interior.ColorIndex = 3
            End If

            'Calculate the percent change between the Open Price and the Close Price
            If OpenPrice <> 0 Then
                ws.Cells(FirstRow, 11) = FormatPercent((ClosePrice - OpenPrice) / OpenPrice, 2)
            Else
                ws.Cells(FirstRow, 11) = 0
            End If

            FirstRow = FirstRow + 1
            OpenPrice = ws.Cells(i + 1, 3)
            Total = 0
            
            'Loop to find the values for the Greatest Increase, Greatest Decrease, and Greatest Volume
            'Including the corresponding ticker
            If ws.Cells(FirstRow, 11) > GreatIncr Then
                GreatIncr = ws.Cells(FirstRow, 11)
                GreatIncrTicker = ws.Cells(FirstRow, 9)
            End If
            
            If ws.Cells(FirstRow, 11) < GreatDecr Then
                GreatDecr = ws.Cells(FirstRow, 11)
                GreatDecrTicker = ws.Cells(FirstRow, 9)
            End If
            If ws.Cells(FirstRow, 12) > GreatVol Then
                GreatVol = ws.Cells(FirstRow, 12)
                GreatVolTicker = ws.Cells(FirstRow, 9)
            End If
            
        End If
    Next i
'Set up the location for where the below values are going to land on the spreadsheet
ws.Cells(2, 16) = GreatIncrTicker
ws.Cells(2, 17) = FormatPercent(GreatIncr, 2)

ws.Cells(3, 16) = GreatDecrTicker
ws.Cells(3, 17) = FormatPercent(GreatDecr, 2)

ws.Cells(4, 16) = GreatVolTicker
ws.Cells(4, 17) = GreatVol
Next ws

'Used this message box to make sure that the ran all the way through.
MsgBox ("Finished!")

End Sub


