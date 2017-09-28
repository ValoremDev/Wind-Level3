Sub Recherche_TRI()
' Sub permettant de rechercher le Tarif correspondant au TRI Actionnaire voulu dans le scénario de production rentré

    Application.ScreenUpdating = False
    Application.StatusBar = "Searching for Project TRI"

    Debug.Print "-----------------------------------------------"
    Debug.Print "Starting IRR Lookup"
    Debug.Print "-----------------------------------------------"

    ' Vérification du Type de tarif utilisé
    Cas_Scenario = Range("Cas_Scenario").Value
    
    prod = Range("Choix_Nh").Value
    
    If Range("Type_Tarif_Actif").Value <> "Appel d'Offre" Then
        MsgBox "Passage en mode Appel d'Offre pour le projet actif"
        Debug.Print "Changing Mode => Appel d'Offre"
        Range("Array_Type_Tarif")(Cas_Scenario).Value = "Appel d'Offre"
    End If

    'Initialisation des variables locales Up et Down
    Up = False
    Down = False
    
    ' Initialisation du Step
    step_Tarif = STEP_TARIF_INITIAL
    
    ' Initialisation du TRI Visé
    TRI_Renta = Range("TRI_Visé").Value
    
    Do
        ' Sculptage DETTE en P90
        Range("Choix_Nh").Value = 3
        Call Model_Debt
        
        ' Calcul du TRI en P Choisi
        Range("Choix_Nh").Value = prod
        Call break_circular_reference_IS
        
        TRI_Actionnaire = Range("TRI_Actionnaire").Value
        
        If TRI_Actionnaire > TRI_Renta + PRECISION_TRI Then
            Up = True
            ' Baisse du Tarif
            Range("Range_Tarif_Recherche_TRI").Offset(0, Cas_Scenario).Value = Range("Range_Tarif_Recherche_TRI").Offset(0, Cas_Scenario).Value - step_Tarif
            Debug.Print TRI_Actionnaire * 100# & "% : Réduction du tarif: " & step_Tarif & " €/MWh"
            Debug.Print "Nouveau Tarif: "; Range("Range_Tarif_Recherche_TRI").Offset(0, Cas_Scenario).Value
            
        ElseIf TRI_Actionnaire < TRI_Renta - PRECISION_TRI Then
            Down = True
            ' Hausse du tarif
            Range("Range_Tarif_Recherche_TRI").Offset(0, Cas_Scenario).Value = Range("Range_Tarif_Recherche_TRI").Offset(0, Cas_Scenario).Value + step_Tarif
            Debug.Print TRI_Actionnaire * 100# & "% : Augmentation du tarif: " & step_Tarif & " €/MWh"
            Debug.Print "Nouveau Tarif: "; Range("Range_Tarif_Recherche_TRI").Offset(0, Cas_Scenario).Value
            
        End If
        
        If Up And Down Then
            ' Réduction du step
            step_Tarif = step_Tarif / 2
            
            Debug.Print "Step tarif: Nouveau Tarif = " & step_Tarif
            Debug.Print ""

            Up = False
            Down = False
        End If
        
    Loop Until Abs(TRI_Actionnaire - TRI_Renta) < PRECISION_TRI

    Debug.Print "IRR FOUND : " & TRI_Actionnaire * 100# & "%"

    Debug.Print "-----------------------------------------------"
    Debug.Print "Ending IRR Lookup"
    Debug.Print "-----------------------------------------------"

    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Sheets("ExecSum").Select
End Sub
