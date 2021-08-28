Sub Stock_Analysis_Bonus()

'Loop through all sheets to apply code for each year (2014,2015,2016)
For Each ws In Worksheets

    'Create Summary Table by defining Column Headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

    'Set an initial variable for holding the ticker string
    Dim Ticker As String

    'Set an initial variable for holding the Total Stock Volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

    'Set initial variables for the Yearly Change and Percent Change Calculations
    Dim Open_Price As Double
    Open_Price = ws.Cells(2, 3).Value

    Dim Close_Price As Double
    Close_Price = 0

    Dim Yearly_Change As Double
    Yearly_Change = 0

    Dim Percent_Change As Double
    Percent_Change = 0

    'Location of each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Finding last Row to identify End of Loop value
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all rows by the ticker string
    For i = 2 To LastRow

        'Check if we are still within the same ticker range
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

            'Set ticker value
            Ticker = ws.Cells(i, 1).Value

            'Print the ticker value in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker

            'Add Stock Volumes to come up to with total for that ticker
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

            'Print the Stock Volume Total in the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

            'Calculate Yearly Change
            Close_Price = ws.Cells(i, 6).Value
            Yearly_Change = Close_Price - Open_Price

            'Print Yearly Change with Condition
                If (Yearly_Change > 0) Then

                    'Colour it in Green if value above 0
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                ElseIf (Yearly_Change <= 0) Then

                    'Colour it in Red if value below 0
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                End If

             'Calculate Percent Change with Condition
                If Open_Price > 0 Then
                Percent_Change = (Yearly_Change / Open_Price)

                Else
                Percent_Change = 0

                End If

            'Print Percent Change in Summary Table with percentage format
            ws.Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "Percent")

            'Reset the row counter by adding 1 to the Summary_Table_Row variable
            Summary_Table_Row = Summary_Table_Row + 1

            'Reset Total Stock Volume
            Total_Stock_Volume = 0

            'Reset Close Price
            Close_Price = 0

            'Reset Open Price
            Open_Price = ws.Cells(i + 1, 3).Value

        Else
            'Continue adding to the Stock Volume Total if the ticker string remains the same
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

        End If

    Next i

'Create Bonus Summary Table by defining Column Headers
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Total Volume"
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"

'Finding last Row of the Summary Table (ST) to identify End of Loop value
LastRow_ST = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To LastRow_ST

        'Find the Max Value in the Percent Change Row
        If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow_ST)) Then
            ws.Range("O2").Value = ws.Cells(i, 9).Value
            ws.Range("P2").Value = Format(ws.Cells(i, 11).Value, "Percent")

        'Find the Min Value in the Percent Change Row
        ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow_ST)) Then
            ws.Range("O3").Value = ws.Cells(i, 9).Value
            ws.Range("P3").Value = Format(ws.Cells(i, 11).Value, "Percent")

        'Find the Largest Total Stock Volume
        ElseIf ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L2:K" & LastRow_ST)) Then
            ws.Range("O4").Value = ws.Cells(i, 9).Value
            ws.Range("P4").Value = ws.Cells(i, 12).Value

        End If

    Next i

Next ws

End Sub
