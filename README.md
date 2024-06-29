# VBA-challenge

'Open VBA_challenge folder in this repo for VBA script file and screenshots of results

'VBA code

Sub quarterly_stocks():
    
    'define dimensions
    Dim total As Double
    Dim i As Long 'row
    Dim j As Integer 'column
    Dim change As Double
    Dim start As Long
    Dim lastRow As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Single
    Dim averageChange As Double
    Dim ws As Worksheet
    
   'define variables and starting values
   'Enable script to run on every worksheet (that is, every quarter) at once.
   
    For Each ws In Worksheets
        i = 0
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        
        'title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        'define data scope
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Create a script that loops through all the stocks for each quarter:
        For i = 2 To lastRow
        
            ticker = ws.Cells(i, 1).Value
    
            If ws.Cells(i + 1, 1).Value <> ticker Then
            
                'store results in variable
                total = total + ws.Cells(i, 7).Value
                
                If total = 0 Then
                    
                    ws.Range("I" & 2 + j).Value = ticker
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                
                Else
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = change / ws.Cells(start, 3)
                
                    start = i + 1
                    'Output each new "Ticker" in column "I"
                    ws.Range("I" & 2 + j) = ticker
                     'output in column J "Quarterly Change" from the opening price at the beginning of a given quarter to the closing price at the end of that quarter
                    ws.Range("J" & 2 + j) = change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    
                    'apply conditional formatting to highlight positive change in green and negative change in red.
                        If change > 0 Then
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        ElseIf change < 0 Then
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                        End If
                    
                    'output in column K "Percent Change"  from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
                    ws.Range("K" & 2 + j).Value = percentChange
                    'apply formatting appropriately to the percent change column
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    'output in column L "Total Stock Volume" of the stock
                    ws.Range("L" & 2 + j).Value = total
                    
                End If
        
                total = 0
                change = 0
                j = j + 1
                days = 0
                dailyChange = 0
                
            Else
                'If next ticker matches, add to previous ticker
                total = total + ws.Cells(i, 7).Value
            
            End If
        
        Next i
     
        'Define new column ranges, new variables for locating max/min numbers and associated tickers
        percent_change_column = ws.Range("K2:K" & lastRow)
        total_stock_column = ws.Range("L2:L" & lastRow)
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(percent_change_column), percent_change_column, 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(percent_change_column), percent_change_column, 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(total_stock_column), total_stock_column, 0)
    
        'Return in "P2" Greatest % Increase output maximum number in Percent Change column and associated ticker in "O2"
        ws.Range("P2") = "%" & WorksheetFunction.Max(percent_change_column) * 100
        ws.Range("O2") = ws.Cells(increase_number + 1, 9)
    
        'Return in "P3" Greatest % Decrease output minimum number in Percent Change column and associated ticker in "O3"
        ws.Range("P3") = "%" & WorksheetFunction.Min(percent_change_column) * 100
        ws.Range("O3") = ws.Cells(decrease_number + 1, 9)
    
        'Return in "P4" Greatest Total Volume ouput maximum number in Total Stock Volume column and associated ticker in "O4"
        ws.Range("P4") = WorksheetFunction.Max(total_stock_column)
        ws.Range("O4") = ws.Cells(volume_number + 1, 9)
    
    Next ws
    
End Sub
