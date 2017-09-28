Sub break_circular_reference_CCA()

'Cette procédure permet de casser la référence circulaire des CCAs entre les onglets Calc_Op et Calc_C
Application.StatusBar = "Breaking Circular Reference CCA"

Do
    
    #If Debugging_CCA Then
            Stop
    #End If
    
    Range("CCAs_copy").Copy
    Range("CCAs_paste").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    Application.Wait (Now + MSECONDS)
    

Loop Until Range("Check_CCA") = "OK"

Application.CutCopyMode = False

End Sub


