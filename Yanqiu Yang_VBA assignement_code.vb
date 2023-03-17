Sub stockloop()

' Set the initial value for holding the ticker name
Dim ticker_name As String

' Set an initial variable for toal volume per ticker
Dim total_volume As Double
total_volume = 0
Dim Percent_Max As Double
Dim Percent_Min As Double
Dim Max_Volume As Double

' Keep track of the location for each ticker in the summary table
Dim Summary_table_row As Integer
Summary_table_row = 2

'need to set up the open prcie of the beginning of the year
Dim open_price As Double
'set up the initial value of open price
open_price = Cells(2, 3).Value

Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'count rows
last_row = Cells(Rows.Count, 1).End(xlUp).Row

' Loop for all tickers total volume
For i = 2 To last_row

    ' check if we are still within the same ticker, if it's not
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'set the ticker name
        ticker_name = Cells(i, 1).Value
        yearly_change = Cells(i, 10).Value
        percent_change = Cells(i, 11).Value
        
        'add to ticker total volume, ticker open & close price
        total_volume = total_volume + Cells(i, 7).Value
        close_price = Cells(i, 6).Value
        
       'calculate yealy change
       yearly_change = (close_price - open_price)
       
       'calculate percent change
       If (open_price) = 0 Then
       Cells(i, 11).Value = 0
       Else
       Cells(i, 11).Value = yearly_change / open_price
       End If
    
        'print the ticker name, yearly change, percent change in the summary table
        Cells(1, 9).Value = "ticker name"
        Range("I" & Summary_table_row).Value = ticker_name
        Cells(1, 10).Value = "Yearly Change"
        Range("J" & Summary_table_row).Value = yearly_change
        Cells(1, 11).Value = "Percent Change"
        Range("K" & Summary_table_row).Value = percent_change
        Range("K" & Summary_table_row).NumberFormat = "0.00%"
        Cells(1, 12).Value = "Total Volume"
        Range("L" & Summary_table_row).Value = total_volume
        
        'Add one to the summary table row
        Summary_table_row = Summary_table_row + 1
        
        'Reset the ticker total volume, open price
        total_volume = 0
        open_price = Cells(i + 1, 3).Value
        
    'if the cell immediately following a row is the same ticker. Add to the ticker total volume
        Else
        
        total_volume = total_volume + Cells(i, 7).Value
    
    End If
    
Next i

    'Greater decrease & increase
    Percent_Max = WorksheetFunction.Max(ActiveSheet.Columns("k"))
    Percent_Min = WorksheetFunction.Min(ActiveSheet.Columns("k"))
    Max_Volume = WorksheetFunction.Max(ActiveSheet.Columns("l"))
    Range("Q2").Value = FormatPercent(Percent_Max)
    Range("Q3").Value = FormatPercent(Percent_Min)
    Range("Q4").Value = Max_Volume
    
    'apply the ticker_name to the results
    For i = 2 To WorksheetFunction.CountA(ActiveSheet.Columns(9))
    If Percent_Max = Cells(i, 11).Value Then
        Range("P2").Value = Cells(i, 9).Value
    ElseIf Percent_Min = Cells(i, 11).Value Then
        Range("P3").Value = Cells(i, 9).Value
    Else
        Max_Volue = Cells(i, 12).Value
        Range("P4") = Cells(i, 9).Value
    End If
    
    'Color the positive change in green, negative change in red in summary table
        If Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        ElseIf Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        End If
    
    'print the name of the summary table
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(4, 15).Value = "Greatest total volume"
    
 Next i
 
End Sub



