Function Calc_CVAE(ChiffreAffaire As Double)
' :ChiffreAffaire: Le chiffre d'affaire annuel
' Retourne la valeur de la base de CVAE en fonction de l'entrée
'Calcul de la CVAE selon les dispositions fiscales en vigueur en 2017 en France
    
    If ChiffreAffaire < 500 Then
        Calc_CVAE = 0
        
    ElseIf ChiffreAffaire < 3000 Then
        Calc_CVAE = 0.5 / 100 * (ChiffreAffaire - 500) / 2500
        
    ElseIf ChiffreAffaire < 10000 Then
        Calc_CVAE = 0.5 / 100 + (0.9 / 100 * (ChiffreAffaire - 3000) / 7000)
        
    ElseIf ChiffreAffaire < 500000 Then
        Calc_CVAE = 1.4 / 100 + (0.1 / 100 * (ChiffreAffaire - 10000) / 40000)
        
    Else: Calc_CVAE = (ChiffreAffaire >= 50000) * 1.5 / 100
    
    End If
    
End Function
