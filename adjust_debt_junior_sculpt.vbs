Sub adjust_debt_junior_sculpt()
' Fixe le montant en Sculpt de la Dette Junior
' Cette prodédure calcule le montant maximal de dette en sculptant avec un DSCR fixe, 
' et casse la référence circulaire qui s'applique avec le calcul des intérêts.

Application.StatusBar = "Sculpting Junior Debt"
Application.ScreenUpdating = False

'Initialisation du DSCR de sculptage : le modèle prend d'abord le DSCR rentré par l'utilisateur
Range("DSCR_calc_junior").Value = Range("DSCR_junior_sculpt").Value

Do Until Range("Check_junior_dette").Value = "OK"

    #If Debugging_JD Then
        Stop
    #End If
    
    Count_SJD = Count_SJD + 1
    
    Range("Debt_junior_copy").Copy
    Range("Debt_junior_paste").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    Call break_circular_reference_IS

    
    If Count_SJD < 2 Then
        Range("Debt_copy").Copy
        Range("Debt_paste").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    End If
Loop

End Sub





