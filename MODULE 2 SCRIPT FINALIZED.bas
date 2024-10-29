Attribute VB_Name = "Module1"
Sub stockanalysis():
'variables
    Dim total As Double
    Dim row As Long
    Dim rowCount As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim summaryTableRow As Long
    Dim stockStartRow As Long
    Dim startValue As Long
    Dim lastTicker As String
    
    'worksheet loop
    For Each ws In Worksheets
   
     'titles for table
     ws.Range("I1").Value = "ticker"
     ws.Range("J1").Value = "quarterly change"
     ws.Range("K1").Value = "percent change"
     ws.Range("L1").Value = "total stock volume"
     ws.Range("p1").Value = "ticker"
     ws.Range("q1").Value = "value"
     ws.Range("o2").Value = "greatest % increase"
     ws.Range("o3").Value = "greatest % decrease"
     ws.Range("o4").Value = "greatest total volume"
    
     'initizalize values
     summaryTableRow = 0
     total = 0
     quarterlyChange = 0
     stockStartRow = 2
     startValue = 2
    
     'last row count
     rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
    
    
     lastTicker = ws.Cells(rowCount, 1).Value
    
     'first loop to end
     For row = 2 To rowCount
    
     'change in ticker
         If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
        
         'change in column a
        
         'add total stock volume
         total = total + ws.Cells(row, 7).Value
        
        
         'value of stock volume equaling to zero
         If total = 0 Then
             ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value 'ticker value added to column a
         'prints 0s if needed be
             ws.Range("J" & 2 + summaryTableRow).Value = 0
             ws.Range("K" & 2 + summaryTableRow).Value = 0
             ws.Range("L" & 2 + summaryTableRow).Value = 0
         Else
             'find first non zero open for stock otherwise, search for first nonzero stock open
             If ws.Cells(startValue, 3).Value = 0 Then
                 For findValue = startValue To row
                 'check if the following open value does not equal zero
                     If ws.Cells(findValue, 3).Value <> 0 Then
                     'update star value to find to reloop
                         startValue = findValue
                         Exit For ' break loop
                     End If
             
                
                 Next findValue
             End If
            
             'quarterly change (last close minus first open)
             quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
            
             'percent change (quarterly,first open)
             percentChange = quarterlyChange / ws.Cells(startValue, 3).Value
            
             'print results
             ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value 'ticker value column A
             ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange 'prints value in column j
             ws.Range("K" & 2 + summaryTableRow).Value = percentChange
             ws.Range("L" & 2 + summaryTableRow).Value = total
             
         'change color depending on value for the quarterly change
             If quarterlyChange > 0 Then
                  ' green
                  ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
             ElseIf quarterlyChange < 0 Then
                 'red
                 ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
             Else
                 'no change/white
                 ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
            
             End If
             
            'reset values for next ticker depending on stock volume, average, quarterly and percentage
             total = 0
             averageChange = 0
             quarterlyChange = 0
             startValue = row + 1
             summaryTableRow = summaryTableRow + 1
             
            
             End If
         
         Else
                 'if in same ticker, keep adding to total stock volume
                 total = total + ws.Cells(row, 7).Value ' from 7th column
         End If
         
     Next row
     
     'recommendation to clean up data to avoid confussion. find last row by finding last ticker
        
     'update summary table row
     summaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
        
     'find the last data in extra rows from columns
     Dim lastExtraRow As Long
     lastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).row
        
        
     'loop clearing extra data from I-L
     For e = summaryTableRow To lastExtraRow
         'columns from 9-12
         For Column = 9 To 12
             ws.Cells(e, Column).Value = ""
             ws.Cells(e, Column).Interior.ColorIndex = 0
         Next Column
     Next e
        
    'print summary of aggregates
    
     ws.Range("q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2))
     ws.Range("q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2))
     ws.Range("q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2))
    
    'match function to find row numbers of tickers associated with the results above
     Dim greatestIncreaseRow As Double
     Dim greatestDecreaseRow As Double
     Dim greatestTVRow As Double
     greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("k2:k" & summaryTableRow + 2)), ws.Range("k2:k" & summaryTableRow + 2), 0)
     greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("k2:k" & summaryTableRow + 2)), ws.Range("k2:k" & summaryTableRow + 2), 0)
     greatestTVRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2)), ws.Range("L2:L" & summaryTableRow + 2), 0)
    
    
    'print ticker symbols for greatest increase, decrease and total
    ws.Range("p2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
    ws.Range("p3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
    ws.Range("p4").Value = ws.Cells(greatestTVRow + 1, 9).Value
    
    'summary table number format decimal percentage
    For s = 0 To summaryTableRow
         ws.Range("j" & 2 + s).NumberFormat = "0.00"
         ws.Range("k" & 2 + s).NumberFormat = "0.00%"
         ws.Range("L" & 2 + s).NumberFormat = "#,###"
        
    Next s
    
     ' format numerical values for increase, decrease, total aggregates
     ws.Range("q2").NumberFormat = "0.00%"  'inc
     ws.Range("q3").NumberFormat = "0.00%"  'dec
     ws.Range("q4").NumberFormat = "#,###"  'total
    
    
     'autofit for columns
     ws.Columns("a:q").AutoFit
    
    Next ws
    
    

End Sub
