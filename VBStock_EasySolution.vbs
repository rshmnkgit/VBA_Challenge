Option Explicit

Sub Easy_Solution(wkSheet As Worksheet)
    Dim ticker As String
    Dim sum_vol As LongLong
    Dim irow, jcol As Integer
    Dim get_lastrow, rownum As Long
    Dim is_firstrow As Boolean

    'Get the last row number
    get_lastrow = wkSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    'MsgBox wkSheet.Name
    'write the column headings
    wkSheet.Range("J1").Value = "Ticker"
    wkSheet.Range("K1").Value = "Total Stock Volume"
    
    'initialize the current row variable
    rownum = 2
    
    is_firstrow = True                  'initialize the firstrow as true for the first ticker
    For irow = 2 To get_lastrow
        ticker = wkSheet.Cells(irow, 1).Value          'assign the ticker value of the current row
        
        If is_firstrow = True Then     'check whether the current row is the first row for this ticker symbol
                sum_vol = 0
        End If
        
        If ticker = wkSheet.Cells(irow + 1, 1).Value Then          'compare ticker value with next row
                is_firstrow = False                                 'if the next row ticker is the same, is_firstrow is reset to false
                sum_vol = sum_vol + wkSheet.Range("G" & irow).Value         'add the current volume to the sum
        Else
                sum_vol = sum_vol + wkSheet.Range("G" & irow).Value         'add the closing volume to the sum
                
                wkSheet.Range("J" & rownum).Value = ticker                              'display the ticker value in the cell
                wkSheet.Range("K" & rownum).Value = sum_vol                       'display the total volume in the cell
                
                rownum = rownum + 1             'increment the current row number
                is_firstrow = True                      'next row is the firstrow for the new ticker
        End If
        
    Next irow
    
End Sub


Sub Main()
        Dim ws As Worksheet
              
        For Each ws In Worksheets
                Easy_Solution ws
        Next ws
End Sub

