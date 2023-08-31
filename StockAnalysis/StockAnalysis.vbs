Sub StockAnalysis ()

  ' each worksheet in the workbook contains data for one year
  ' run this script on each worksheet
  For Each ws in Worksheets

    ' generate the title row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("Q2").Value = "Greatest % Increase"
    ws.Range("Q3").Value = "Greatest % Decrease"
    ws.Range("Q4").Value = "Greatest Total Volume"

    ws.Columns("I:L").AutoFit
    ws.Columns("P:Q").AutoFit

    Dim averageChange As Double
    Dim change        As Double  ' opening daily stock price - closing daily stock price
    Dim dailyChange   As Double
    Dim days          As Integer
    Dim i             As Long    ' row counter
    Dim j             As Integer
    Dim percentChange As Double
    Dim rowCount      As Long    ' the number of rows in the data set
    Dim start         As Long
    Dim total         As Double  ' total stock volume

    ' Rows.Count drops down to the last row in the worksheet (1048576)
    ' .End(xlUp) goes up to the last nonempty cell of the data set
    ' to avoid selection errors, such as
    ' excluding empty cells in the data set, or
    ' dropping down to the last row in the worksheet from the last nonempty cell of the data set
    rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row

    change = 0
    j      = 0
    start  = 2
    total  = 0

    For i = 2 To rowCount
      ' --------------------
      ' if the ticker in the next row (i + 1) of col A
      ' is DIFFERENT than
      ' the ticker in the current row (i) of col A
      ' --------------------
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        total = total + ws.Cells(i, 7).Value ' add current ticker vol to total vol
                                             ' and do some stuff before the next ticker
        
        If total = 0 Then
          ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
          ws.Range("J" & 2 + j).Value = 0
          ws.Range("K" & 2 + j).Value = "%" & 0
          ws.Range("L" & 2 + j).Value = 0
        Else
          ' find the first non zero starting value
          If ws.Cells(start, 3) = 0 Then
            For find_value = start To i
              If ws.Cells(find_value, 3).Value <> 0 Then
                start = find_value
                Exit For
              End If
            Next find_value
          End If

          change           = ws.Cells(i, 6) - ws.Cells(start, 3) '  close - open
          percentageChange = change / ws.Cells(start, 3)         ' (close - open) / open

          start = i + 1

          ' print the results
          ws.Range("I" & 2 + j).Value        = ws.Cells(i, 1).Value ' ticker          `Ticker`
          ws.Range("J" & 2 + j).Value        = change               ' absolute change `Yearly Change`
          ws.Range("J" & 2 + j).NumberFormat = "0.00"
          ws.Range("K" & 2 + j).Value        = percentChange        ' percent  change `Percent Change`
          ws.Range("K" & 2 + j).NumberFormat = "0.00"
          ws.Range("L" & 2 + j).Value        = total                ' total           `Total Stock Volume`

          ' color code column `Yearly Change`
          Select Case change
            Case Is > 0
              ws.Range("J" & 2 + j).Interior.ColorIndex = 4 ' green
            Case Is < 0
              ws.Range("J" & 2 + j).Interior.ColorIndex = 3 ' red
            Case Else
              ws.Range("J" & 2 + j).Interior.ColorIndex = 0 ' white
          End Select

        End If

        ' reset the variables for a new ticker
        change = 0
        days   = 0
        j      = j + 1
        total  = 0

      ' --------------------
      ' if the ticker in the next row (i + 1) of col A
      ' is NOT DIFFERENT than
      ' the ticker in the current row (i) of col A
      ' --------------------
      Else
        total = total + ws.Cells(i, 7).Value ' add current ticker vol to total vol
                                             ' and move on to the next row
      End If
    Next i

    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100 ' Greatest % Increase   = max(percentChange) x 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100 ' Greatest % Decrease   = min(percentChange) x 100
    ws.Range("Q4") =       WorksheetFunction.Max(ws.Range("L2:L" & rowCount))       ' Greatest Total Volume = max(totalStockVolume)

    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    volume_number   = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number   + 1, 9)
  Next ws
End Sub