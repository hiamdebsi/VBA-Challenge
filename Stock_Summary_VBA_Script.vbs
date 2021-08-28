Sub Stock_Summary()

'Create Summary Table by defining Column Headers
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"

' Set an initial variable for holding the ticker code
Dim Ticker As String

' Set an initial variable for holding the Total Stock Volume
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Set initial variables for the Yearly Change and Percent Change Calculations
Dim Open_Price As Double
Open_Price = Cells(2, 3).Value

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
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all rows by the ticker name
For i = 2 To LastRow

    'Check if we are still within the same Ticker Range
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

        'Set Ticker Value
        Ticker = Cells(i, 1).Value

        'Print the Ticker Value in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker

        'Add Stock Volumes to come up to with total for that ticker
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

        'Print the Stock Volume Total in the Summary Table
        Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

        'Calculate Yearly Change
        Close_Price = Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price

        'Print Yearly Change with Condition
            If (Yearly_Change > 0) Then

                'Colour it in Green, because value above 0
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Range("J" & Summary_Table_Row).Value = Yearly_Change

            ElseIf (Yearly_Change <= 0) Then

                'Colour it in Red, because value below 0
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Range("J" & Summary_Table_Row).Value = Yearly_Change

            End If

         'Calculate Percent Change with Condition
            If Open_Price > 0 Then
            Percent_Change = (Yearly_Change / Open_Price)

                Else
                Percent_Change = 0

            End If

        'Print Percent Change in Summary Table as a Percentage Format
        Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "Percent")

        'Reset the row counter by adding 1 to the Summary_Table_Row
        Summary_Table_Row = Summary_Table_Row + 1

        'Reset Total Stock Volume
        Total_Stock_Volume = 0

        'Reset Close Price
        Close_Price = 0

        'Reset Open Price
        Open_Price = Cells(i + 1, 3).Value

    Else
        'Continue adding to the Stock Volume Total if the ticker name remains the same
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

Next i

End Sub
