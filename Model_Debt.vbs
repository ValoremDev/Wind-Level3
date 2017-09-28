Sub Model_Debt()
' Sub faisant la distinction entre les deux méthodes de sculptage de dette

Type_Amortissement = Range("Type_Amo").Value
Application.StatusBar = "Running Model..."

Call break_circular_reference_CCA

'Cette procédure permet de sculter la dette et d'appeller la macro brisant la référence circulaire de l'IS

If Type_Amortissement = "K+I constant" Then
    Call adjust_debt_constant
Else
    Call adjust_debt_sculpt
End If

Application.CutCopyMode = False
Sheets("ExecSum").Activate

End Sub
