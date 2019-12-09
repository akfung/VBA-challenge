Sub VBA_HW():
    'loop over each worksheet
    Dim current As Worksheet
    For Each current In Worksheets
        'variables n shit
        Dim lastRow As Long
        Dim printRow As Double
        Dim currenitVol As Double
        Dim yearOpenPrice As Long
        Dim yearEndPrice As Long
        Dim yearChange As Long
        'find last occupied row in active sheet, set stock vol to 0 and printrow to row 2
        lastRow = current.Cells(Rows.Count, 2).End(xlUp).Row
        currentVol = 0 
        printRow = 2
        yearOpenPrice = current.Cells(2, 3).Value
        'Print header row
        current.Cells(1, 9).Value = "Ticker"
        current.Cells(1, 10).Value = "Yearly Change"
        current.Cells(1, 11).Value = "Percent Change"
        current.Cells(1, 12).Value = "Total Stock Volume"

        
        

        'For loop to add up the stock volume for each entry
        For i = 2 to lastRow
        'Loop to add the trading volume for that row to the running total in currentVol 
            currentVol = currentVol + current.Cells(i, 7)
            
            'enter stuff if the ticker in the currently read row is different than the ticker in the next row
            If current.Cells(i, 1).Value <> current.Cells(i + 1, 1).Value Then
                'print ticker
                current.Cells(printRow, 9).Value = current.Cells(i, 1).Value

                'print yearly change
                yearEndPrice = current.Cells(i, 6)
                yearChange = yearEndPrice - yearOpenPrice
                current.Cells(printRow, 10).Value = yearChange

                'set cell format to percentage and print percent change 
                'Also include If statment to handle divided by 0 nonsense
                current.Cells(printrow, 11).NumberFormat = "0.00%"
                If yearOpenPrice <> 0 Then
                    current.Cells(printRow, 11).Value = yearChange/yearOpenPrice
                Else 
                    current.Cells(printRow, 11).Value = 0
                End If

                'Conditional formatting depending on if the perent change is positive or negative
                If current.Cells(printrow, 11).Value >= 0 Then
                    current.Cells(printrow, 11).Interior.ColorIndex = 4
                Else
                    current.Cells(printrow, 11).Interior.ColorIndex = 3
                End If
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