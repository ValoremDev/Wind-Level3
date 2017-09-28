Public Sub Sortie_BP()
' Sub TOP level regroupant l'ensemble de l'ajustement
' Passage du modèle en P90 pour sculpter la Dette
' PUIS Passage en Productible sélectionné pour ajuster l'IS et obtenir la rentabilité

Debug.Print "-----------------------------------------------"
Debug.Print "Starting Model"
Debug.Print ""

prod = Range("Choix_Nh").Value

Count_SJD = 0

    If prod = 3 Then
        
        Call Model_Debt
        
    Else
    
        Range("Choix_Nh").Value = 3
        
        Call Model_Debt
        
        Range("Choix_Nh").Value = prod
        Call break_circular_reference_cash_sweep
                
    End If

Debug.Print ""
Debug.Print "Ending Model"
Debug.Print "-----------------------------------------------"

ThisWorkbook.Worksheets("ExecSum").Activate
Application.StatusBar = False

End Sub

