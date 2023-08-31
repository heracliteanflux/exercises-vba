Sub StockAnalysis ()

  ' each worksheet in the workbook contains data for one year
  ' run this script on each worksheet
  For Each ws in Worksheets

    ' generate the header row for the output
    ws.Range("I1").Value = "Ticker"                ' header for output row per ticker
    ws.Range("J1").Value = "Yearly Change"         ' header for output row per ticker
    ws.Range("K1").Value = "Percent Change"        ' header for output row per ticker
    ws.Range("L1").Value = "Total Stock Volume"    ' header for output row per ticker

    ws.Range("P1").Value = "Ticker"                ' header row for tickers with greatest values
    ws.Range("Q1").Value = "Value"                 ' header row for tickers with greatest values
    ws.Range("Q2").Value = "Greatest % Increase"
    ws.Range("Q3").Value = "Greatest % Decrease"
    ws.Range("Q4").Value = "Greatest Total Volume"

    ws.Columns("I:L").AutoFit
    ws.Columns("P:Q").AutoFit

    Dim change        As Double  ' dailyChange = opening daily stock price - closing daily stock price
    Dim i             As Long    ' row counter for the entire data set
    Dim j             As Integer ' points to the output row which holds the printed results for a ticker, starting with row 2
    Dim percentChange As Double  ' (opening daily stock price - closing daily stock price) / opening daily stock price
    Dim rowCount      As Long    ' the number of rows in the data set
    Dim start         As Long    ' points to a row of col 3, opening daily stock price
    Dim total         As Double  ' total stock volume

    ' Rows.Count drops down to the last row in the worksheet (1048576)
    ' .End(xlUp) goes up to the last nonempty cell of the data set
    ' to avoid selection errors, such as
    ' excluding empty cells in the data set, or
    ' dropping down to the last row in the worksheet from the last nonempty cell of the data set
    rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' initialization
    change = 0
    j      = 0
    start  = 2 ' skip the header row and point to the first row of data
    total  = 0

    ' for each row in the data set
    For i = 2 To rowCount

      ' --------------------
      ' if the ticker in the next row (i + 1) of col A
      ' is DIFFERENT than
      ' the ticker in the current row (i) of col A
      ' --------------------
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        total = total + ws.Cells(i, 7).Value ' add current ticker vol to total vol
                                             ' as would be done in the case that the ticker is still the same

        ' print zero values in the case in which
        ' the total stock volume for this ticker is equal to zero
        If total = 0 Then
          ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value ' ticker  
          ws.Range("J" & 2 + j).Value = 0                    ' zero absolute change
          ws.Range("K" & 2 + j).Value = "%" & 0              ' zero percent  change
          ws.Range("L" & 2 + j).Value = 0                    ' zero total stock volume

        Else

          ' find the cell containing the first nonzero opening daily stock price
          If ws.Cells(start, 3) = 0 Then                 ' if the opening daily stock price (Col 3) is equal to zero...
            For find_value = start To i                  ' ...iterate over rows to the current row pointer
              If ws.Cells(find_value, 3).Value <> 0 Then ' if any value is nonzero...
                start = find_value                       ' ...then assign the cell that contains the nonzero value to `start`
                Exit For                                 ' ...and break out of the loop
              End If
            Next find_value
          End If

          ' compute the daily change and daily percentage change
          change        = ws.Cells(i, 6) - ws.Cells(start, 3) '  close - open
          percentChange = change / ws.Cells(start, 3) * 100   ' (close - open) / open * 100

          ' increment the cell containing the first nonzero opening daily stock price
          start = i + 1

          ' print the results
          ws.Range("I" & 2 + j).Value        = ws.Cells(i, 1).Value ' ticker
          ws.Range("J" & 2 + j).Value        = change               ' yearly  change
          ws.Range("J" & 2 + j).NumberFormat = "0.00"
          ws.Range("K" & 2 + j).Value        = percentChange        ' percent change
          ws.Range("K" & 2 + j).NumberFormat = "0.00"
          ws.Range("L" & 2 + j).Value        = total                ' total stock volume

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