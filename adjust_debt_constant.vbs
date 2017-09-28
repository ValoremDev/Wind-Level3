Sub adjust_debt_constant()
' Calcul de la quantité de dette levable en cas de sculptage en K+I constant avec respect d'un DSCR minimum

'On vide la ligne de cash sweep paste afin d'aviter tout problème de sizing
Application.StatusBar = "Adjusting Constant Senior Debt"
Range("cash_sweep_paste").ClearContents

Do
    Range("DSCR_target").GoalSeek Goal:=Range("DSCR_const").Value, ChangingCell:=Range("Dette_const")
    Call break_circular_reference_IS

    Range("DSRA_KI_Copy").Copy
    Range("DSRA_KI_Paste").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    
    If Range("type_DSCR_junior").Value = "Sculpt" Then
        Call adjust_debt_junior_sculpt
    End If

Loop Until Range("Check_Dette").Value = "OK"

'On vide la ligne de cash sweep paste pour éviter d'afficher un Check "NOT OK"
Range("cash_sweep_paste").ClearContents

End Sub
