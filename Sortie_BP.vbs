Public Sub Sortie_BP()

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

Debug.Print "-----------------------------------------------"
Debug.Print "Ending Model"
Debug.Print ""

ThisWorkbook.Worksheets("ExecSum").Activate
Application.StatusBar = False

End Sub

