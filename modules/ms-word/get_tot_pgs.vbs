Sub get_tot_pgs_doc()
tot_pgs = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
MsgBox (tot_pgs)
End Sub