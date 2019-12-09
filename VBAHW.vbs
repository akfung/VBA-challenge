Sub VBA_HW():
    'loop over each worksheet
    Dim current As Worksheet
    For Each current In Worksheets
        'variables n shit
        Dim lastRow As Long
        Dim printRow As Double
        Dim currenitVol As Double
        Dim yearOpenPrice As Double
        Dim yearEndPrice As Double
        Dim yearChange As Double
        'find last occupied row in active sheet, set stock vol to 0 and printrow to row 2
        lastRow = current.Cells(Rows.Count, 2).End(xlUp).Row
        currentVol = 0 
        printRow = 2
        yearOpenPrice = current.Cells(i, 3).Value
        'Print header row
        current.Cells(1, 9).Value = "Ticker"
        current.Cells(1, 10).Value = "Yearly Change"
        current.Cells(1, 11).Value = "Percent Change"
        current.Cells(1, 12).Value = "Total Stock Volume"

        
        

        'For loop to add up the stock volume for each day entry
        For i = 2 to lastRow
        'Loop to add the trading volume for that row to the running total in currentVol 
            currentVol = currentVol + current.Cells(i, 7)
            
            'enter stuff if the ticker changes
            If current.Cells(i, 1).Value <> current.Cells(i + 1, 1).Value Then
                'print ticker
                current.Cells(printRow, 9).Value = current.Cells(i, 1).Value
                'print yearly change
                yearEndPrice = current.Cells(i, 3)
                yearChange = yearOpenPrice - yearEndPrice
                current.Cells(printRow, 10).Value = yearChange
                'print percent change
                current.Cells(printRow, 11).Value = yearChange/yearOpenPrice
                'print total year volume
                current.Cells(printRow, 12).Value = currentVol
                'reset total volume, set open price for next stock, iterate printrow
                currentVol = 0
                yearOpenPrice = current.Cells(i + 1, 3)
                printRow = printRow + 1
            End If
        Next i
    Next current
End Sub