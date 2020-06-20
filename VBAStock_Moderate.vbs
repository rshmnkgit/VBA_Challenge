Option Explicit

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


Sub Main()
        Dim ws As Worksheet
        For Each ws In Worksheets
                Moderate_Solution ws
        Next ws
End Sub

