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
        'find last occupied row in active sheet, set stock vol to 0 and printrow to row 2
        lastRow = Cells(Rows.Count, 2).End(xlUp).Row
        currentVol = 0 
        printRow = 2

        'Print header row
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        yearOpenPrice = Cells(2, 3).Value
        

        'For loop to add up the stock volume for each day entry
        For i = 2 to lastRow
            currentVol = currentVol + Cells(i, 7)

            'enter stuff if the ticker changes
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                'print ticker
                Cells(printRow, 9).Value = Cells(i, 1).Value
                'print yearly change
                yearEndPrice = Cells(i, 3)
                Cells(printRow, 10).Value = yearOpenPrice - yearEndPrice
                'print percent change
                Cells(printRow, 11).Value = Cells(printRow, 10).Value/yearOpenPrice
                'print total year volume
                Cells(printRow, 12).Value = currentVol
                'reset total volume, set open price for next stock, iterate printrow
                currentVol = 0
                yearOpenPrice = Cells(i + 1, 3)
                printRow = printRow + 1
            End If
        Next i
    Next 
End Sub