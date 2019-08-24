Sub stockscriptpt2()
'TAKE TICKER, YEARLY NUMERICAL CHANGE FROM OPENING TO CLOSING, %CHANGE, TOTAL STOCK VOLUME
'CONDITIONAL FORMATING FOR POSITIVE IN GREEN, NEGATIVE IN RED
    'Loop through each ws
    For Each ws In Worksheets
            
        'Number of rows in sheet
        NumOfRows = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Make Ticker total table
        SummaryTableRow = 2

        'Set up first row of table
        ws.Range("I1").Value = "Ticker Name"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Ticker Volume Total"

        'Loop through rows, either adding to volume total or sending previous total to a summary table
        'and moving onto the next unique ticker
        For i = 2 To NumOfRows

            If i = 2 Then
            
                'Make new ticker to compare
                ticker_name = ws.Cells(i, 1).Value

                'Reset Ticker volume total
                tickervolumetotal = ws.Range("G2").Value
                
                openingvalue = ws.Cells(i, 3).Value
            
            'Determine if current cell ticker value is the same as the previous
            ElseIf ws.Cells(i, 1).Value <> ticker_name Then

                Dim percentchange As Double

                closingvalue = ws.Cells(i - 1, 6).Value
                yearlychange = closingvalue - openingvalue
                percentchange = Round(openingvalue / yearlychange, 2)

                ws.Range("I" & SummaryTableRow).Value = ticker_name
                ws.Range("L" & SummaryTableRow).Value = tickervolumetotal
                ws.Range("J" & SummaryTableRow).Value = yearlychange
                ws.Range("K" & SummaryTableRow).Value = percentchange


                'Make new ticker to compare
                ticker_name = ws.Cells(i, 1).Value
                
                'Increment Summary Table Row
                SummaryTableRow = SummaryTableRow + 1

                'Reset Ticker volume total
                tickervolumetotal = 0

                openingvalue = ws.Cells(i, 3).Value

            Else
                tickervolumetotal = tickervolumetotal + ws.Cells(i, 7).Value
            End If
               
        Next i

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"


        NumOfRows = ws.Cells(Rows.Count, "I").End(xlUp).Row
        j = 2
        
        highestpercent = ws.Cells(j, 11).Value
        
        highestpercent.NumberFormat = "00.0%"

        lowestpercent = ws.Cells(j, 11).Value
        
        lowestpercent.NumberFormat = "00.0%"
        
        highestvolume = ws.Cells(j, 12).Value

        Rows = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For j = 3 To Rows
            
            If ws.Cells(j, 11).Value > highestpercent Then
                highestpercent = ws.Cells(j, 11).Value
                highesttickerper = ws.Cells(j, 8).Value
            End If

            If ws.Cells(j, 11).Value < lowestpercent Then
                lowestpercent = ws.Cells(j, 11).Value
                lowesttickerper = ws.Cells(j, 8).Value
            End If

            If ws.Cells(j, 12).Value > highestvolume Then
                highestvolume = ws.Cells(j, 12).Value
                highesttickervol = ws.Cells(j, 8).Value
            End If

        Next j

        ws.Range("P2").Value = highesttickerper
        ws.Range("P3").Value = lowesttickerper
        ws.Range("P4").Value = highesttickervol
        ws.Range("Q2").Value = highestpercent
        ws.Range("Q3").Value = lowestpercent
        ws.Range("Q4").Value = highestvolume

    Next ws
    MsgBox ("finished")
End Sub