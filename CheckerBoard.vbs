' draw an 8 x 8 checker board

Sub CheckerBoard ()
  ' a counter to track the cell number
  Dim cell_num As Integer
  Dim i, j     As Integer
  cell_num = 1

  For i = 1 To 8                            ' for each row of the board
    For j = 1 To 8                          ' for each cell of the row
      If cell_num Mod 2 = 0 then            ' if the row number is even...
        Cells(i, j).Interior.ColorIndex = 1 ' ...color the cell black
      Else                                  ' else
        Cells(i, j).Interior.ColorIndex = 3 ' ...color the cell red
      End If
      cell_num = cell_num + 1
    Next j
    cell_num = cell_num + 1
  Next i
End Sub