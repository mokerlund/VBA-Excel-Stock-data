sub stockscript()

    'Loop through each worksheet
    For each ws in Worksheets

        ' Assigned last ticker to compare to current (ws.Range used to simplify)
        ticker_name = ws.Range("A2").value
        
        'Number of rows in sheet
        NumOfRows = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Make Ticker total table
        SummaryTableRow = 2

        'Set up first row of table
        ws.Range("I1").value = "Ticker Name"
        ws.Range("J1").value = "Ticker Volume Total"

        'Ticker volume total to add and keep as a record for each (ws.Range used to simplify)
        tickervolumetotal = ws.Range("G2").value

        'Loop through rows, either adding to volume total or sending previous total to a summary table
        'and moving onto the next unique ticker
        for i = 3 to NumOfRows

            'Determine if current cell ticker value is the same as the previous 
            if ws.cells(i,1).value <> ticker_name then

            'Make new ticker to compare
            ticker_name = ws.cells(i, 1).value

            'Print to Summary Table
            ws.Range("I" & SummaryTableRow).value = ticker_name
            ws.Range("J" & SummaryTableRow).value = tickervolumetotal
            
            'Increment Summary Table Row
            SummaryTableRow = SummaryTableRow + 1

            'Reset Ticker volume total
            tickervolumetotal = 0

            else
        
                tickervolumetotal = tickervolumetotal + ws.cells(i, 7).value
            End if
        Next i
    Next ws

    Msgbox("Finished!")

end sub