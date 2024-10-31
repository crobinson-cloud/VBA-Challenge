Attribute VB_Name = "Module1"
Sub stockAnalysis():

    Dim total As Double                           ' total stock value
    Dim row As Long                               ' loop control variable that will go through in a sheet
    Dim rowCount As Double                 ' variable that holds the number of rows in a sheet
    Dim quarterlyChange As Double       ' variable that holds the quarterly change for each stock in a sheet
    Dim percentChange As Double         ' variable that holds the percentchange for each stock in a sheet
    Dim summaryTablerow As Long      ' variable that holds the rows of the summary table row
    Dim stockStartRow As Long            ' variable that holds the start of a stock's rows in the sheet
    Dim StartValue As Long                 ' start row for a stock (location of first open)
    Dim lastTicker As String                ' finds the last tickerin the sheet
    
    ' Set the Title Row of the Summary section
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Value"
    
    ' Setup the title row of the Aggregate Section
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'initialize values
    summaryTablerow = 0                  ' summary table row starts at 0 in the sheet (add 2) in relation to the header
    total = 0                                       ' total stock volume for a stock starts at 0
    quarterlyChange = 0                     ' quarterly change starts at 0
    stockStartRow = 2                       ' first stock on the sheet starts on row 2
    StartValue = 2                             ' first open of the first stock is on row 2
    
    '   get the value of the last row in the current sheet
    rowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    '   find the last ticker so that we can break out of the last loop
        lastTicker = Cells(rowCount, 1).Value
        
    ' loop until we get to the end of the sheet
    For row = 2 To rowCount
    
        ' check to see if the ticker changed
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
        
            'if there is a change in Column A (Column 1)
            
            ' add to the total stock volume one last time
            total = total + Cells(row, 7).Value   ' Gets value from the 7th column( G)
            
            ' check to see if the value of the total stock volume is 0
            If total = 0 Then
                ' print the results in the summary tabke section (Columns I - L)
                Range("I" & 2 + summaryTablerow).Value = Cells(row, 1).Value   'prints the ticker value from column A
                Range("J" & 2 + summaryTablerow).Value = 0                             ' prints a 0 in column J (Quarterly Change)
                Range("K" & 2 + summaryTablerow).Value = 0                           'prints a 0 in column K (% change)
                Range("L" & 2 + summaryTablerow).Value = 0                            'prints a 0 in column L (total stock volume)
            Else
                ' find the first non-zero first open value for the stock
                If Cells(StartValue, 3).Value = 0 Then
                    ' if the first open is 0, search for the first non-zero stock open value by moving to the next rows
                    For findValue = StartValue To row
                    
                        ' check to see if the next (or rows afterwards) open value does not equal 0
                        If Cells(findValue, 3).Value <> 0 Then
                            'once we have a non-zero first open value, that value becomes the row where we track our first open
                            StartValue = findValue
                            ' break out of the loop
                            Exit For
                        End If
                    
                    Next findValue
                    
                    End If
                        
                        ' calcutlate the quarterly change (difference in the last close -first open)
                        quarterlyChange = Cells(row, 8).Value - Cells(StartValue, 3).Value
                        
                        ' calculate the percent change (quarterly change / first open)
                        percentChange = quarterlyChange / Cells(StartValue, 3).Value
                        
                        
                       ' print the results in the summary table section (Columns I - L)
                    Range("I" & 2 + summaryTablerow).Value = Cells(row, 1).Value   'prints the ticker value from column A
                    Range("J" & 2 + summaryTablerow).Value = quarterlyChange   ' prints a value in column J (Quarterly Change)
                    Range("K" & 2 + summaryTablerow).Value = percentChange    'prints a value in column K (% change)
                    Range("L" & 2 + summaryTablerow).Value = total                     'prints a value in column L (total stock volume)
                        
                      'color the Quarterly Change column in the summary section based on the value of the quarterly change
                      If quarterlyChange > 0 Then
                        ' color the cell green
                        Range("J" & 2 + summaryTablerow).Interior.ColorIndex = 4
                      ElseIf quarterlyChange < 0 Then
                        ' color the cell red
                        Range("J" & 2 + summaryTablerow).Interior.ColorIndex = 3
                    Else
                        ' color the cell clear or no change
                        Range("J" & 2 + summaryTablerow).Interior.ColorIndex = 0
                  End If
                  
                  ' reset / update the values for the next ticker
                  total = 0                             ' resets total stock volume for the next ticker
                  averageChange = 0            ' resets the average change for the next ticker
                  quarterlyChange = 0          ' resets the quarterly change for the next ticker
                  ' move to the next row in the summary table
                  summaryTablerow = summaryTablerow + 1
                
                End If
            
            Else
            
                '   If we are in the same ticker, keep adding to the total stock volume
                total = total + Cells(row, 7).Value   ' Gets value from the 7th column( G)
                
        End If
        
    Next row
    
      ' clean up (if needed) to avoid extra data being placed in the summary section
      ' find the last row of data in the summary table by finding the last ticker in the summary section
        
      ' update the summary table row
        summaryTablerow = Cells(Rows.Count, "I").End(xlUp).row
        
    ' find the last data in the extra rows from columns J-L
    Dim lastExtraRow As Long
    lastExtraRow = Cells(Rows.Count, "J").End(xlUp).row
        
    ' loop that clears the extra data from columns I-L
    For e = summaryTablerow To lastExtraRow
            ' for loop that goes through columns I-L (9-12)
            For Column = 9 To 12
            Cells(e, Column).Value = ""
            Cells(e, Column).Interior.ColorIndex = 0
        Next Column
      Next e
    
      ' print the summary aggregates
      ' after generating info in the summary section, find the greatest % increase and decrease, then find the greatest total stock volume
      Range("Q2").Value = WorksheetFunction.Max(Range("K2:K" & summaryTablerow + 2))
      Range("Q3").Value = WorksheetFunction.Min(Range("K2:K" & summaryTablerow + 2))
      Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & summaryTablerow + 2))
    
    ' use Match () to find the row numbers of the ticker names associated with the greates % increase and decrease, then find the greatest total stock volume
    Dim greatestIncreaseRow As Double
    Dim greatestDecreaseRow As Double
    Dim greatestTotVolRow As Double
    greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & summaryTablerow + 2)), Range("K2:K" & summaryTablerow + 2), 0)
    greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & summaryTablerow + 2)), Range("K2:K" & summaryTablerow + 2), 0)
    greatestTotVolRow = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & summaryTablerow + 2)), Range("L2:L" & summaryTablerow + 2), 0)
        
    'display the ticker symbol for the greatest % increase, greatest % decrease, greatest total stock volume
     Range("P2").Value = Cells(greatestIncreaseRow + 1, 9).Value
     Range("P3").Value = Cells(greatestDecreaseRow + 1, 9).Value
     Range("P4").Value = Cells(greatestTotVolRow + 1, 9).Value
        
   ' format the summary table columns
    For s = 0 To summaryTablerow
        Range("J" & 2 + SummaryRow).NumberFormat = "0.00"               ' formats the Quarterly Change
        Range("K" & 2 + SummaryRow).NumberFormat = "0.00%"            ' formats the Percent Change
        Range("L" & 2 + SummaryRow).NumberFormat = "#,###"             ' formats the total stock volume
   Next s
       
    ' format the summary aggregates
    Range("Q2").NumberFormat = "0.00%"                          ' formats the greatest % increase
    Range("Q3").NumberFormat = "0.00%"                          ' formats the greatest % decrease
    Range("Q4").NumberFormat = "#,###"                          ' formats the greatest total stock volume
    
    'Autofit infor across all columns
    Columns("A:Q").AutoFit
    
    
    
    
    
    
End Sub

