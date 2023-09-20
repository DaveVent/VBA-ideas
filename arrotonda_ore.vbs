' FUNZIONE PER ARROTONDARE LE ORE LAVORATE A UNA DETERMINATA UNITA' DI QUARTO D'ORA, MEZZ'ORA, TRE QUARTI D'ORA E ORA

Function arrotonda_ore(x As Variant, Optional quarto_ora As Double = 0, Optional mezz_ora As Double = 0, Optional tre_quarti_ora As Double = 0, Optional ora As Double = 0)
Dim risultato As Double, valore As Double

    'Estrae i minuti
    valore = x.Value - Int(x.Value)
    
    
    'Mezz'ora
    If quarto_ora = 0 And mezz_ora <> 0 Then
    
        If valore < mezz_ora Then
        
            risultato = Int(x.Value)
            
        ElseIf valore >= mezz_ora And valore <= tre_quarti_ora Then
        
            risultato = Int(x.Value) + 0.5
        
        ElseIf valore >= tre_quarti_ora And valore <= ora Then
        
            risultato = Int(x.Value) + 0.75
        
        Else:
        
            risultato = Int(x.Value) + 1
    
        End If
    
    'Tre quarti d'ora
    ElseIf quarto_ora = 0 And mezz_ora = 0 And tre_quarti_ora <> 0 Then
    
        If valore < tre_quarti_ora Then
        
            risultato = Int(x.Value)
        
        ElseIf valore >= tre_quarti_ora And valore <= ora Then
        
            risultato = Int(x.Value) + 0.75
        
        Else:
        
            risultato = Int(x.Value) + 1
    
        End If
    
    'Ora
    ElseIf quarto_ora = 0 And mezz_ora = 0 And tre_quarti_ora = 0 Then
    
        If valore < ora Then
        
            risultato = Int(x.Value)
        
        Else:
        
            risultato = Int(x.Value) + 1
    
        End If
    
    'Quarto d'ora
    Else:
    
        If valore < quarto_ora Then
        
            risultato = Int(x.Value)
            
        ElseIf valore >= quarto_ora And valore < mezz_ora Then
        
            risultato = Int(x.Value) + 0.25
            
        ElseIf valore >= mezz_ora And valore <= tre_quarti_ora Then
        
            risultato = Int(x.Value) + 0.5
        
        ElseIf valore >= tre_quarti_ora And valore <= ora Then
        
            risultato = Int(x.Value) + 0.75
        
        Else:
        
            risultato = Int(x.Value) + 1
        
        End If
    
    
    End If


'Output
arrotonda_ore = risultato

End Function
