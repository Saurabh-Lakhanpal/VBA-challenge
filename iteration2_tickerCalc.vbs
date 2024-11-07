Attribute VB_Name = "Module2"
'This is the Second iteration that Calculates the following information only on worksheet "Q1" in addition to First iteration.
'   1. Column "I" - Distinct ticker symbol.
'   2. Column "J" - Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
'   3. Adds conditional formatting that will highlight positive change in Light green and negative change in Pink.
'   4. Column "k" - The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
'   5. Column "L" - The total stock volume of the stock.
'   6. Auto fits the cell value to width so that it is legible.

Sub tickerCalc()

    
        Dim ws As Worksheet
        
        'Declare a variable for holding the ticker name
        Dim tickername As String
    
        'Declare a varable for holding a total sum of sold volume of the ticker trade
        Dim tickervolume As Double
        tickervolume = 0

        'Keep track of the location for each ticker name in summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Quarterly Change = Close Price at the end of a trading Quarter - Open Price at the beginning of the trading Quarte
        'Percent change =((Close - Open)/Open)*100
        Dim open_price As Double
        'Set starting open_price. Other opening prices will be determined in the conditional loop.
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim Quarterly_change As Double
        Dim percent_change As Double

        'Label the summary Table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows by the ticker names
        'Make sure that the ticker names are sorted and are alpha-numeric/string variables.

        For i = 2 To lastrow

            'Searches for when the value of the next cell is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Set the ticker name
                tickername = Cells(i, 1).Value

                'Add the volume of trade
                tickervolume = tickervolume + Cells(i, 7).Value

                'Print the ticker name in the summary table
                Range("I" & summary_ticker_row).Value = tickername

                'Print the ticker name in the summary table
                Range("L" & summary_ticker_row).Value = tickervolume

                'Now collect information about closing price
                close_price = Cells(i, 6).Value

                'Calculate Quarterly change
                Quarterly_change = (close_price - open_price)
              
                'Print the Quarterly change for each ticker in the summary table
                Range("J" & summary_ticker_row).Value = Quarterly_change

                'Check for the non-divisibilty condition when calculating the percent change
                If (open_price = 0) Then

                    percent_change = 0

                Else
                    
                    percent_change = Quarterly_change / open_price
                
                End If

                'Print the Quarterly change for each ticker in the summary table
                Range("K" & summary_ticker_row).Value = percent_change
                Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
                'Reset the row counter. Add one to the summary_ticker_row
                summary_ticker_row = summary_ticker_row + 1

                'Reset volume of trade to zero
                tickervolume = 0

                'Reset the opening price
                open_price = Cells(i + 1, 3)
            
                Else
              
                'Add the volume of trade
                tickervolume = tickervolume + Cells(i, 7).Value

            
            End If
        
        Next i

    'Conditional formatting that will highlight positive change in light green and negative change in pink
    'First find the last row of the summary table

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code Quarterly change
    
    For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 35
            Else
                Cells(i, 10).Interior.ColorIndex = 38
            End If
    Next i
    
        'Autofit to keep the population legible
        Cells.Rows.AutoFit
        Cells.Columns.AutoFit

End Sub
