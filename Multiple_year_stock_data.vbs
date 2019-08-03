Sub CompileData()

    Dim TestCount, LastRow, LastRowNew, TickerCount As Integer
    Dim VolumeTotal, PercentChange, Gvolume, Ginc, Gdec, YearOpen, YearClose, PriceDif, NewTicker As Double
    Dim Ginctick, Gdectick, Gvoltick As String

For Each ws In Worksheets    
    'Set Variables'
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastRowNew = ws.Cells(Rows.Count, 11).End(xlUp).Row
    VolumeTotal = 0
    TickerCount = 2
    NewTicker = 2
    TestCount = 0
    Ginc = 0
    Gdec = 0
    Gvolume = 0
        
    'Create Headers'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Cycle through all rows'
    For i = 2 To LastRow

        'Calculating Volume'
        VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value

        
        'Check to see if we get to a new ticker'
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Save the name of the ticker before change'
            Ticker = ws.Cells(i, 1).Value

            'Print the ticker and the volume'
            ws.Range("I" & TickerCount).Value = Ticker
            ws.Range("L" & TickerCount).Value = VolumeTotal

            'reset volume after a new ticker'
            VolumeTotal = 0

            'Establish year open'
            YearOpen = ws.Cells(NewTicker, 3).Value

            'Establish year close'
            YearClose = ws.Cells(i, 6).Value
            
            'finds the row of the changed ticker'
            NewTicker = i + 1

            'Calculate Difference'
            PriceDif = YearClose - YearOpen

            'Print the price difference'
            ws.Range("J" & TickerCount).Value = PriceDif

            'Conditional Formatting for price difference'
            If PriceDif >= 0 Then
                ws.Range("J" & TickerCount).Interior.ColorIndex = 4
            Else
                ws.Range("J" & TickerCount).Interior.ColorIndex = 3
            End If

            'find the percentage change'
            If YearOpen <> 0 Then
                PercentChange = Round(PriceDif / YearOpen, 4)
            Else   
                PercentChange = 0
            End If

            'format and print percentage change'
            ws.Range("K" & TickerCount).Value = Format(PercentChange, "0.00%")

            'count tickers to print on the correct line'
            TickerCount = TickerCount + 1
        
        End If

    Next i
    
        'set headers for hard solution'
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'loop through the change data'
        For j = 2 To LastRowNew
            
            'find if the next percent change is larger and save ticker'
            If ws.Cells(j + 1, 11).Value > Ginc Then
                Ginc = ws.Cells(j + 1, 11).Value
                Ginctick = ws.Cells(j + 1, 9).Value
            End If
            
            'find if the next percent change is lower and save ticker'
            If ws.Cells(j + 1, 11).Value < Gdec Then
                Gdec = ws.Cells(j + 1, 11).Value
                Gdectick = ws.Cells(j + 1, 9).Value
            End If
        
            'find if the next volume is larger and save ticker'
            If ws.Cells(j + 1, 12).Value > Gvolume Then
                Gvolume = ws.Cells(j + 1, 12).Value
                Gvoltick = ws.Cells(j + 1, 9).Value
            End If


        Next j
        
        'print the values'
        ws.Range("P2").Value = Ginctick
        ws.Range("P3").Value = Gdectick
        ws.Range("P4").Value = Gvoltick
        ws.Range("Q2").Value = Format(Ginc, "0.00%")
        ws.Range("Q3").Value = Format(Gdec, "0.00%")
        ws.Range("Q4").Value = Gvolume

next ws

End Sub


