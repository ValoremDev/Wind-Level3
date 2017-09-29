Sub adjust_debt_sculpt()

'Cette prodédure calcule le montant maximal de dette en sculptant avec un DSCR fixe, et casse la référence circulaire qui s'applique avec le calcul des intérêts.
'On vide la ligne de cash sweep paste afin d'aviter tout problème de sizing


Application.StatusBar = "Sculpting Senior Debt"
'Set delta_CAPEX = Range("delta_CAPEX").Offset(0, 1 - scenario)

If debugBool = True Then
    Debug.Print "Debug Loop n° " & debugLoopValue
    Valindex = CStr(debugLoopValue)
    Set CrawlerJson("'" & Valindex & "'") = JsonConverter.ParseJson("{""Exit Code"" : "" "" ,""Delta Capex"" : 0, ""DSCR Calc"" : 0, ""Gearing Calc"" : 0}")
    Dim debugSDLoop As Integer: debugSDLoop = 0
    Debug.Print delta_CAPEX.Value
End If

Range("cash_sweep_paste").ClearContents

Application.ScreenUpdating = False

Dim Gearing_sup As Boolean
Gearing_sup = False
Dim Loop_Count As Integer: Loop_Count = 0

'Initialisation du DSCR de sculptage : le modèle prend d'abord le DSCR rentré par l'utilisateur
Range("DSCR_calc").Value = Range("DSCR_sculpt").Value

Do Until Range("Flag_Dette").Value < 0.0000001 And Range("Check_Bilan").Value = "OK" _
    And Range("Check_IS").Value = "OK" And Range("Check_dette").Value = "OK" _
    And Range("Check_gearing") = "OK" And Range("Check_junior_dette") = "OK" _
    And Loop_Count > 0

    
'Si le gearing était trop élevé, on revient à la balise "Gearing_Max" et on resculpte la dette
Gearing_Max:
    Do
        Do
            If debugBool = True Then
                debugSDLoop = debugSDLoop + 1
                
            End If
            
            Application.StatusBar = "Sculpting Senior Debt"
            Loop_Count = Loop_Count + 1
            
            Range("Debt_copy").Copy
            Range("Debt_Paste").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
            
            Application.Wait (Now + MSECONDS)
            
            Range("DSRA_Sculpt_Copy").Copy
            Range("DSRA_Sculpt_Paste").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
            
            If Range("Check_Dette") = "OK" And Range("Flag_Dette").Value < 0.0000001 Then
                Debug.Print "Breaking IS"
                Call break_circular_reference_IS
                If Range("Check_IS").Value = "OK" And Range("Check_Dette").Value = "OK" Then
                    Do
                        Debug.Print "Breaking CCA"
                        Call break_circular_reference_CCA
                        Call break_circular_reference_IS
                    Loop Until Range("Check_IS").Value = "OK"
                End If
            End If
            
            If Range("type_DSCR_junior").Value = "Sculpt" And Range("active_Junior_Debt").Value = "Oui" Then
                Call adjust_debt_junior_sculpt
            End If
            
            If debugBool = False And Loop_Count = 40 Then
                Debug.Print "Asserting infinite loop, Manually raising DSCR by 0.001"
                Range("DSCR_Calc").Value = Range("DSCR_Calc").Value + 0.001
            End If
            
            If debugBool = True And debugSDLoop = 50 Then
                
                Debug.Print "Exiting Sculpting after " & debugSDLoop & " Loops"
                CrawlerJson("'" & Valindex & "'")("Exit Code") = "Loop ERROR"
                CrawlerJson("'" & Valindex & "'")("Delta CAPEX") = delta_CAPEX.Value
                GoTo outputDebug
            End If

        Loop Until Range("Check_Dette") = "OK"
                
    Loop Until Range("Check_Dette") = "OK" And Range("Check_IS") = "OK"

' Incrémentation du DSCR : si le gearing calculé est supérieur au gearing max, on augmente le DSCR et on re-sculpte la dette
' Si le DSCR est inférieur à [gearing max - la précision], on re-sculpte la dette en diminuant progressivement le DSCR  pour atteindre un DSCR
' de sculptage qui permette au gearing  calculé d'être le plus proche possible du gearing max
' Si le DSCR a été rentré en profil manuel, ce module est inutile.

    If Range("Type_profil_DSCR").Value = "Constant" Then
    
        If Range("Gearing_calc").Value > Range("Gearing_Max").Value Then
        
            Range("DSCR_calc").Value = Range("dscr_calc").Value + STEP_DSCR_1
            
            Debug.Print "Augmentation du DSCR : " & Range("DSCR_Calc").Value & " de " & STEP_DSCR_1
            Debug.Print ""
            
            Gearing_sup = True
            GoTo Gearing_Max
            
        End If
        
        If Range("Gearing_calc").Value < Range("Gearing_Max").Value - PRECISION_GEARING_1 And Gearing_sup = True Then
            
            Range("DSCR_calc").Value = Range("dscr_calc").Value - STEP_DSCR_2
            Debug.Print "Diminution du DSCR : " & Range("DSCR_Calc").Value & " de " & STEP_DSCR_2
            Debug.Print ""

            GoTo Gearing_Max

        End If
        
        If Range("Gearing_calc").Value < Range("Gearing_Max").Value - PRECISION_GEARING_2 And Gearing_sup = True Then
            Range("DSCR_calc").Value = Range("dscr_calc").Value - STEP_DSCR_3
            
            Debug.Print "Diminution du DSCR : " & Range("DSCR_Calc").Value & " de " & STEP_DSCR_3
            Debug.Print ""
            
            GoTo Gearing_Max

        End If
        
        If Range("Gearing_calc").Value < Range("Gearing_Max").Value - PRECISION_GEARING_3 And Gearing_sup = True Then
            Range("DSCR_calc").Value = Range("dscr_calc").Value - STEP_DSCR_4
            
            Debug.Print "Diminution du DSCR : " & Range("DSCR_Calc").Value & " de " & STEP_DSCR_4
            Debug.Print ""
            
            GoTo Gearing_Max

        End If
        
    End If

Loop

    If debugBool = True Then
            CrawlerJson("'" & Valindex & "'")("Exit Code") = "OK"
            CrawlerJson("'" & Valindex & "'")("Delta CAPEX") = delta_CAPEX.Value
            CrawlerJson("'" & Valindex & "'")("DSCR Calc") = Range("DSCR_calc").Value
            CrawlerJson("'" & Valindex & "'")("Gearing Calc") = Range("Gearing_calc").Value
    End If
outputDebug:




Gearing_sup = False
'On vide la ligne de cash sweep paste pour éviter d'afficher un Check "NOT OK"
Range("cash_sweep_paste").ClearContents

End Sub

