Sub Stocks()

    'Define Variables
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim Volume As Double
    Dim chart As Integer
    Dim finalrow As Long
    Dim percentageChange As Double


    finalrow = Cells(Rows.Count, 1).End(xlUp).Row
    openPrice = 0
    closePrice = 0
    Volume = 0
    yearChange = 0
    chart = 2

    'Print Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"

    'Retrieve value of first ticker as opening values are taken from first day of each ticker
    For i = 1 To 2
            openPrice = Cells(2, 3).Value
    Next i

    'Process rest of data
    For i = 2 To finalRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Create summary table
            ' Retrieve Ticker and print to table
            ticker = Cells(i, 1).Value
            Range("I" & chart).Value = Cells(i, 1).Value 

            'Retrieve Volume data and print to chart
            Volume = Volume + (Cells(i, 7).Value)
            Range("L" & chart).Value = Volume

            'Retrieve closing price for ticker
            closePrice = Cells(i, 6).Value   

            'Calculate Yearly change and format accordingly
            yearChange = closePrice - openPrice
            Range("J" & chart).Value = yearChange
                If yearChange > 0 Then
                    Range("J" & chart).Interior.ColorIndex = 4
                Else
                    Range("J" & chart).Interior.ColorIndex = 3
                End If

            'Calculate percentage and take into account the possibility of 0 Value ticker
                If (openPrice = 0) And (closePrice = 0) Then
                    Range("K" & chart).Value = 0
                Else
                    percentageChange = (closePrice / openPrice) - 1
                    Range("K" & chart).Value = percentageChange
                End if
            Range("K" & chart).NumberFormat = "0.00%"

            'Reset Variables to 0 and Set chart to next row
            chart = chart + 1
            openPrice = 0
            closePrice = 0
            yearChange = 0
            Volume = 0

            'Copy opening Price from from first day for next ticker
            openPrice =Cells(i + 1, 3).Value
    
        Else
            'Add volume data to Volume variable
            Volume = Volume + (Cells(i, 7).Value)
        End If
    Next i

End Sub