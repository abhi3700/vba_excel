' Get current page no. where the cursor is on.
Sub getCurrentPageNum()
    Dim intCurrentPage As Integer

    intCurrentPage = Selection.Information(wdActiveEndAdjustedPageNumber)
    MsgBox (intCurrentPage)
End Sub