Option Explicit

'Calculate and display the yearly change, precentage of change and total stock value for all the tickers
Sub Moderate_Solution(wkSheet As Worksheet)
        Dim ticker As String
        Dim sum_vol As LongLong
        Dim year_open, year_close, year_change, perc_change As Double
        Dim irow, jcol As Integer
        Dim get_lastrow, rownum As Long
        Dim is_firstrow As Boolean
        
        'Get the last row number
        get_lastrow = wkSheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        'display the column headings
        wkSheet.Range("J1").Value = "Ticker"
        wkSheet.Range("K1").Value = "Yearly Change"
        wkSheet.Range("L1").Value = "Percent Change"
        wkSheet.Range("L1:L" & get_lastrow).NumberFormat = "0.00%"              'format the percentage column
        wkSheet.Range("M1").Value = "Total Stock Volume"
        
        
        'initialize the current row variable
        rownum = 2
        is_firstrow = True
        For irow = 2 To get_lastrow
                ticker = wkSheet.Cells(irow, 1).Value          'assign the ticker value of the current row

                If is_firstrow = True Then                                       'check if the current row is the first row for the current ticker
                        year_open = wkSheet.Cells(irow, 3).Value        'set the opening value for the year for the current ticker
                        sum_vol = 0                                                     'initialise sum of stock as zero
                End If

                 If ticker = wkSheet.Cells(irow + 1, 1).Value Then          'compare ticker value with next row
                        is_firstrow = False
                        sum_vol = sum_vol + wkSheet.Range("G" & irow).Value             'add the stock value to the total sum
                 Else
                        year_close = wkSheet.Range("F" & irow).Value                            'set the closing value of the tickers last row as year closing value
                        year_change = year_close - year_open                                        'calculate the yearly change.  difference between closing value & opening value for the year
                        sum_vol = sum_vol + wkSheet.Range("G" & irow).Value             'add the stock value to the total sum
                        If year_open <> 0 Then
                                perc_change = year_change / year_open                           'calculate the yearly change percentage if opening value is not zero
                        Else:
                                perc_change = 0
                        End If
                        
                        wkSheet.Range("J" & rownum).Value = ticker                            'display the ticker value
                        wkSheet.Range("K" & rownum).Value = year_change                 'display the yearly change value
                        If year_change < 0 Then                                                             'check if the yearly change is negative
                                wkSheet.Range("K" & rownum).Interior.Color = vbRed              'color the cell red if yearly change is negative
                        Else
                                wkSheet.Range("K" & rownum).Interior.Color = vbGreen            'color the cell green if yearly change is positive
                        End If
                        wkSheet.Range("L" & rownum).Value = perc_change                         'display the percentage of change
                        wkSheet.Range("M" & rownum).Value = sum_vol                             'display the total stock volume

                        rownum = rownum + 1                                                                   'increment the rownumber so as to display the next ticker in the next row
                        is_firstrow = True                                                                           'set the first row as true for the next ticker
                End If
        Next irow
End Sub

'Calculate and display the greatest increase, greatest decrease, and max stock volume
Sub Hard_Solution(wkSheet As Worksheet)
        Dim max_perc, min_perc As Double
        Dim max_vol As LongLong
        Dim irow, get_lastrow As Integer

        'display labels
        wkSheet.Range("P4").Value = "Greatest % Increase"
        wkSheet.Range("P5").Value = "Greatest % Decrease"
        wkSheet.Range("P6").Value = "Greatest Stock Volume"
        wkSheet.Range("Q3").Value = "Ticker"
        wkSheet.Range("R3").Value = "Value"
        
        max_perc = wkSheet.Range("L2").Value                    'set the maximum percentage to be the first ticker percentage change
        min_perc = wkSheet.Range("L2").Value                     'set the minimum percentage to be the first ticker percentage change
        max_vol = wkSheet.Range("M2").Value                     ' set the maximum total stock value to be the first ticker total stock
        get_lastrow = wkSheet.Cells(Rows.Count, "L").End(xlUp).Row              'get the last row number for the displayed ticker values
        Dim ticker1, ticker2, ticker3 As String
        
        For irow = 2 To get_lastrow
                
                If wkSheet.Range("L" & irow).Value > max_perc Then              'check if the current percentage is greater than the max_perc
                        max_perc = wkSheet.Range("L" & irow).Value                 'assign max_perc as the current value
                        ticker1 = wkSheet.Range("J" & irow).Value                      'assign the corresponding ticker id to the ticker1 variable
                'End If
                ElseIf wkSheet.Range("L" & irow).Value < min_perc Then          'check if the current percentage is less than the min_perc
                        min_perc = wkSheet.Range("L" & irow).Value                   'assign min_perc value as the current value
                        ticker2 = wkSheet.Range("J" & irow).Value                       'assign the corresponding ticker id to the ticker2 variable
                End If
                If wkSheet.Range("M" & irow).Value > max_vol Then               'cheack if current total is greater than the saved max_vol
                        max_vol = wkSheet.Range("M" & irow).Value                   'assign max_stock value as the current value
                        ticker3 = wkSheet.Range("J" & irow).Value                       'assign the corresponding ticker id to the ticker3 variable
                End If
        Next irow
        
        'display the ticker values and the results in the worksheet
        wkSheet.Range("Q4").Value = ticker1
        wkSheet.Range("Q5").Value = ticker2
        wkSheet.Range("Q6").Value = ticker3
        wkSheet.Range("R4").Value = max_perc
        wkSheet.Range("R4").NumberFormat = "0.00%"
        wkSheet.Range("R5").Value = min_perc
        wkSheet.Range("R5").NumberFormat = "0.00%"
        wkSheet.Range("R6").Value = max_vol
End Sub

Sub Main()
        Dim ws As Worksheet
        
        For Each ws In Worksheets
                Moderate_Solution ws            'call the sub routine Moderate_Solution
                Hard_Solution ws                   'call the sub routine Hard_Solution
        Next ws
End Sub


