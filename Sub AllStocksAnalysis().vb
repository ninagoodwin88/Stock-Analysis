Sub AllStocksAnalysis()
    Worksheets("AllStocksAnalysis").Activate

    Range("A1").Value = "Stock Analysis"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim tickers(11) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'loop over all the rows
    For i = 2 To RowCount
        'loop over the columns
        For j = 0 To 11
        'Do stuff with tickers
    
        
        If Cells(i, 1).Value = tickers(11) Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> tickers(11) And Cells(i, 1).Value = tickers(11) Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> tickers(11) And Cells(i, 1).Value = tickers(11) Then

            endingPrice = Cells(i, 6).Value

        End If
        Next

    Next i

    Worksheets("AllStocksAnalysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1


End Sub

