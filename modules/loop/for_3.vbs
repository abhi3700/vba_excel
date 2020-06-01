Dim c As Integer, i As Integer, j As Integer

For c = 1 To 3
    For i = 1 To 6
        For j = 1 To 2
            Worksheets(c).Cells(i, j).Value = 100
        Next j
    Next i
Next c