Attribute VB_Name = "Module5"
'This is an additional functionality, where I added a button for the reviewer to click and see the iteration by iteration results with each click.

Sub RunNextMacro()
    Dim HiddenSheet As Worksheet
    Dim counter As Integer

    ' Set the worksheet and get the counter value
    Set HiddenSheet = ThisWorkbook.Sheets("HiddenSheet")
    counter = HiddenSheet.Range("A1").Value

    ' Run the corresponding macro
    Select Case counter
        Case 1
            Call tickerTotals: MsgBox "Sheet Q1 Shows Each Ticker Stock's Totals Volume, TRIGGER MACRO AGAIN"
        Case 2
            Call tickerCalc: MsgBox "Sheet Q1 Shows Each Ticker Stock's Price Changes, TRIGGER MACRO AGAIN"
        Case 3
            Call tickerAnalysis: MsgBox "Sheet Q1 Shows Each Ticker Stock's Sales Analysis, TRIGGER MACRO AGAIN"
        Case 4
            Call tickerloop: Sheets("Q1").Activate: MsgBox "All Quaterl'y Sheets show analysis, NEXT TRIGGER RESETS WORKBOOK"
        Case 5
            For Each ws In ThisWorkbook.Worksheets
            ws.Columns("H:R").ClearContents
            ws.Columns("H:R").Interior.ColorIndex = xlNone
            ws.Activate
            ActiveWindow.DisplayGridlines = True
            Cells.Columns.ColumnWidth = 13
            Next ws
            Sheets("Q1").Activate
        Case Else
            Exit Sub
    End Select
    
    ' Update the counter
    counter = counter + 1
    If counter > 5 Then counter = 1:
    ' Save the updated counter value
    HiddenSheet.Range("A1").Value = counter

End Sub



