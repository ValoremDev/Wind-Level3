Sub break_circular_reference_IS()

'Cette procédure permet de casser la référence circulaire de l'IS sur l'onglet calc
Application.StatusBar = "Breaking Circular Reference IS"

Do
    #If Debugging_IS Then
        Stop
    #End If
    
    Range("IS_copy").Copy
    Range("IS_paste").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    Application.Wait (Now + MSECONDS)
    

Loop Until Range("Check_IS") = "OK"

Application.CutCopyMode = False

End Sub

    


    

