Sub getPageDim()
Dim objPage As Page
 
Set objPage = ActiveDocument.ActiveWindow _
 .Panes(1).Pages.Item(1)
MsgBox ("width: " & objPage.Width & ", height: " & objPage.Height)
End Sub