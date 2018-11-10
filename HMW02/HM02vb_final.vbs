
Function getChangeValue(n_open, n_close)
    ' gets difference between closing and opening values
    getChangeValue = n_close - n_open
End Function

Function getPerChange(n_endValue, n_startValue)
' gets the perccentage in chnage, to be formated as % in excel.
    If (n_startValue > 0 And n_endValue > 0) Then
        getPerChange = (((n_endValue - n_startValue) * 100) / n_startValue) / 100
    Else
        getPerChange = n_endValue / 100
    End If
 End Function

Public Sub printGreatestTitles()
'Prints hard option titles
    Cells(2, "O") = "Greatest % Increase"
    Cells(3, "O") = "Greatest % Decrease"
    Cells(4, "O") = "Greatest Total Volume"
    Cells(1, "P") = "Ticker"
    Cells(1, "Q") = "Value"
End Sub



Public Sub runAllSheets()
' https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
' https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet
        Dim WS_Count As Integer
        Dim i As Integer
        WS_Count = ActiveWorkbook.Worksheets.Count

         For i = 1 To WS_Count
           ActiveWorkbook.Worksheets(i).Activate
             Call getVolumeByTicker
            MsgBox "Working on  :" + ActiveWorkbook.Worksheets(i).Name

         Next i

End Sub



' Print Volume by Ticker
Public Sub getVolumeByTicker()
    Dim usedRows As Double
    Dim tickerName As String
    Dim volume As Double
    Dim tickerRow As Integer
    Dim startValue As Double
    Dim endValue As Double
    Dim change As Double
    Dim pchange As Double
   
    Dim greatestArr(2, 1) As String
    ' (0,0): Greatest Incr Ticker   (0,1) Greatest Incr Val
    ' (1,0): Greatest Dec Ticker  (1,1) Greatest Dec Val
    ' (2,0): Greatest Vol Ticker (2,1) Greatest vol Val
    'Initilialize greatest ones
        For i = 1 To 3
            For j = 1 To 2
                greatestArr(i - 1, j - 1) = "0"
            Next j
        Next i
    ' Initialize number or data Rows, first Ticker name and Volume, first opening value
    ' and first print row.
    usedRows = ActiveSheet.UsedRange.Rows.Count
    tickerName = Cells(2, 1).Value
    volume = Cells(2, "G").Value
    startValue = Cells(2, "C").Value
    tickerRow = 2

    'Go thu all rows, check if ticker changed, add up volume if no change
    For i = 2 To usedRows + 1
        If (Cells(i, "A").Value = tickerName) Then
            volume = volume + Cells(i, "G")
        Else
            'started another ticker, get closing value, get change from opening to closing and print it
            endValue = Cells(i - 1, "F").Value
            change = getChangeValue(startValue, endValue)
            Cells(tickerRow, "J") = change
            'get percentage in change  and print
            pchange = getPerChange(endValue, startValue) 'print % of change
            Cells(tickerRow, "K") = pchange
            ' print ticker and volume
            Cells(tickerRow, "I") = tickerName
            Cells(tickerRow, "L") = volume
            
            ' Check for greatestness
            If (pchange > CDbl(greatestArr(0, 1))) Then
                greatestArr(0, 0) = tickerName
                greatestArr(0, 1) = Str(pchange)
                
            End If
            If (pchange < CDbl(greatestArr(1, 1))) Then
                greatestArr(1, 0) = tickerName
                greatestArr(1, 1) = Str(pchange)
            End If
            If (volume > CDbl(greatestArr(2, 1))) Then
                greatestArr(2, 0) = tickerName
                greatestArr(2, 1) = Str(volume)
            End If
            
            ' increment row position
            tickerRow = tickerRow + 1
            ' set initial ticker and volume values for new ticker
            If (Cells(i, "A").Value <> "") Then
                volume = Cells(i, "G").Value
                tickerName = Cells(i, "A").Value
                startValue = Cells(i, "C").Value
            End If
        End If
    Next i
    
    printGreatestTitles
     'print Greatest values
      For i = 1 To 3
            For j = 1 To 2
                Cells(i + 1, j + 15) = greatestArr(i - 1, j - 1)
            Next j
        Next i
    
    
End Sub




