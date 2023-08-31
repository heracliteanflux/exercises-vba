Sub Colors ()  
  Dim red    As Integer
  Dim green  As Integer
  Dim blue   As Integer
  Dim yellow As Integer

  Cells(1, 1) = "Red"
  Cells(1, 2) = "Green"
  Cells(1, 3) = "Blue"
  Cells(1, 4) = "Yellow"

  red    = 3
  green  = 4
  blue   = 5
  yellow = 6
  
  Cells(1, 1).Font.ColorIndex = red
  Cells(1, 2).Font.ColorIndex = green
  Cells(1, 3).Font.ColorIndex = blue
  Cells(1, 4).Font.ColorIndex = yellow
  
  Range("A2:A5").Interior.ColorIndex = red
  Range("B2:B5").Interior.ColorIndex = green
  Range("C2:C5").Interior.ColorIndex = blue
  Range("D2:D5").Interior.ColorIndex = yellow
End Sub