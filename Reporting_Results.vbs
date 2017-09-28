Sub Reporting_Results()
' Reporte les résultats des 3 scénarios de Productible (P90 -> P50) dans l'onglet "Slide COMEC"

Dim loopVal As Integer: loopVal = 3
Application.ScreenUpdating = False

    While loopVal > 0:
        Range("Choix_Nh").Value = loopVal
        If Range("Choix_Nh").Value = 3 Then
        
            Call Model_Debt
            Call break_circular_reference_cash_sweep
            
        Else

            Call break_circular_reference_cash_sweep
            
        End If
        
            Range("TRI_FP_Sortie").Copy
            
            Worksheets("Slide COMEX").Activate
                
            Range("Cell_Out").Offset(RowOffset:=-0, columnOffset:=-loopVal).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

            Range("TRI_Valorem_Sortie").Copy
            Range("Cell_Out").Offset(RowOffset:=1, columnOffset:=-loopVal).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

            Range("Payback_Sortie").Copy
            Range("Cell_Out").Offset(RowOffset:=2, columnOffset:=-loopVal).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    
        loopVal = loopVal - 1
    
    Wend

Application.CutCopyMode = False
Application.StatusBar = False
Application.ScreenUpdating = True

End Sub
