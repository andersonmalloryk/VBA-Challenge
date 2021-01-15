Attribute VB_Name = "Module1"
Sub marketAnalysis()

'-------------------------------------
'Loop to run through all the worksheets
'-------------------------------------
For Each ws In ActiveWorkbook.Worksheets
ws.Activate
'-------------------------------------
'Loop through all of the sheets
'-------------------------------------

    Dim lastRow As Long
    Dim start As Double
    Dim Final As Double

    Dim Ticker As String

    Dim yearlyChange As Double
    Dim PercentChange As String
    Dim volumeTotal As LongLong

    Dim tickerSummary As Integer
    tickerSummary = 2

    start = Cells(2, 3).Value
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set color
    loss = 3
    profit = 4

    'Add the column labels
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly_Change"
    Cells(1, 11) = "Percent_Change"
    Cells(1, 12) = "Total_Stock_Volume"

    'Loop through all stocks for one year
    For i = 2 To lastRow
              
        'Check ticker to see if it has changed
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the Ticker name
            Ticker = Cells(i, 1).Value
        
            'Reset the Final
            Final = Cells(i, 6).Value
        
            'Set the yearlyChange to the last closing number Cells (i,6) first open number Cells(i,3)
            yearlyChange = (Final - start)
                
                If start <> 0 Then
                PercentChange = ((Final - start) / start)
                Else
                PercentChange = 0
                End If
                
            'Add to the volumeTotal
            volumeTotal = volumeTotal + Cells(i, 7).Value
        
            'Print the ticker name to the tickerSummary table
            Range("I" & tickerSummary).Value = Ticker
        
            'Print the yearlyChange to the tickerSummary table
            Range("J" & tickerSummary).Value = yearlyChange
        
                'color cells
                If yearlyChange > 0 Then
                Range("J" & tickerSummary).Interior.ColorIndex = profit
                ElseIf yearlyChange <= 0 Then
                Range("J" & tickerSummary).Interior.ColorIndex = loss
                End If
             
            'Print the percentChange to the tickerSummary table
            Range("K" & tickerSummary).Value = PercentChange
            Range("K" & tickerSummary).NumberFormat = "0.00%"
        
            'Print the volumeTotal to the tickerSummary table
            Range("L" & tickerSummary).Value = volumeTotal
        
            'Add one to the tickerSummary table
            tickerSummary = tickerSummary + 1
        
            'Reset the volumeTotal to zero
            volumeTotal = 0
        
            'Rest Start
            start = Cells(i + 1, 3).Value
        
            'if the cell immediately following a row is the same ticker
            Else
             
            'add to the volumeTotal
            volumeTotal = volumeTotal + Cells(i, 7).Value
    
        End If

    Next i

Next ws

End Sub
