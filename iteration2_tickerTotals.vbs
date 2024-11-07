Attribute VB_Name = "Module1"
'This is the First iteration that propogates the following information only on worksheet "Q1".
'   1. Column "I" - Distinct Ticker Symbols.
'   2. Column "J" - Aggregates the total stock volume sold for each of the stock.
    
Sub tickerTotals()

    
        Dim ws As Worksheet

        'Declare a variable for holding the ticker name
        Dim tickername As String
        
    
        'Declare a varable for holding a total sum of sold volume of the ticker trade
        Dim tickervolume As Double
        tickervolume = 0

        'Keep track of the location for each ticker name in summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2

        'Label the summary Table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Volume"

        'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows by the ticker names
        'Make sure that all the ticker names are sorted and are alpha-numeric/string variables.
            
        For i = 2 To lastrow

            'Searches for when the value of the next cell is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Set the ticker name
                tickername = Cells(i, 1).Value

                'Add the volume of trade
                tickervolume = tickervolume + Cells(i, 7).Value

                'Print the ticker name in the summary table
                Range("I" & summary_ticker_row).Value = tickername

                'Print the trade volume for each ticker in the summary table
                Range("J" & summary_ticker_row).Value = tickervolume

                'Add one to the summary_ticker_row
                summary_ticker_row = summary_ticker_row + 1

                'Reset tickervolume to zero
                tickervolume = 0

            Else
              
                'Add the volume of trade
                tickervolume = tickervolume + Cells(i, 7).Value

            End If
        
        Next i
          
            'Autofit to keep the population legible
            Cells.Rows.AutoFit
            Cells.Columns.AutoFit
  

End Sub
