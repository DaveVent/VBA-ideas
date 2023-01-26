' FUNZIONE PER CONVERTIRE UN NUMERO DECIMALE ARROTONDATO AL 60ESIMO IN UN NUMERO ARROTONDATO AL 100ESIMO (ES. 3,3 -> 3,5) E VICEVERSA (ES. 3,5 -> 3,3)

Function converti_formato_num(x As Double, Optional conversione As Long = 100)

If conversione = 100 Then
    
    converti_formato_num = Int(x) + (x - Int(x)) / 60 * 100

ElseIf conversione = 60 Then

    converti_formato_num = Int(x) + (x - Int(x)) * 60 / 100

Else

converti_formato_num = CVErr(xlErrValue)

End If

End Function
