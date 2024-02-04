Sub DelEmptyPara()
Do
With Selection.Find
.Text = "^p^p"
.Replacement.Text = "^p"
.Forward = True
.Wrap = wdFindContinue
End With
Loop Until Selection.Find.Execute(Replace:=wdReplaceAll) = False
End Sub
