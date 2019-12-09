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
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim mostVolume As Double

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

        'Final rundown stats for greatest % increase/decrease and greatest volume
        'print labels for final rundown and set the ticker value to the first entry in the year end stats table. Set value formats
        current.Cells(2, 14).Value = "Greatest % Increase"
        current.Cells(3, 14).Value = "Greatest % Decrease"
        current.Cells(4, 14).Value = "Greatest Total Volume"
        current.Cells(1, 15).Value = "Stock"
        current.Cells(1, 16).Value = "Value"
        current.Range("P2:P3").NumberFormat = "0.00%"
        greatestDecrease = current.Cells(2,11).Value
        greatestIncrease = current.Cells(2,11).Value
        mostVolume = current.Cells(2,12).Value

        'Loop through each row to check if it has the highest increase/decrease/volume so far
        For ticker = 2 To printRow-1
            'Check if this stock has the greatest % increase
            If current.Cells(ticker, 11) > greatestIncrease Then
                current.Cells(2, 15).Value = current.Cells(ticker, 9)
                greatestIncrease = current.Cells(ticker, 11)
                current.Cells(2, 16).Value = greatestIncrease
            End If

            'Check if this stock has the greatest % decrease
            If current.Cells(ticker, 11) < greatestDecrease Then
                current.Cells(3, 15).Value = current.Cells(ticker, 9)
                greatestDecrease = current.Cells(ticker, 11)
                current.Cells(3, 16).Value = greatestDecrease
            End If

            'Check if this stock has the most volume
            If current.Cells(ticker, 12) > mostVolume Then
                current.Cells(4, 15).Value = current.Cells(ticker, 9)
                mostVolume = current.Cells(ticker, 12)
                current.Cells(4, 16).Value = mostVolume
            End If

        Next ticker

    Next current
End Sub