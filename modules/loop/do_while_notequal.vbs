Dim i As Integer
i = 1

Do While Cells(i, 1).Value <> ""
    Cells(i, 2).Value = Cells(i, 1).Value + 10
    i = i + 1
Loop