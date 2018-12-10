Sub EntrPosicionador(item_partnumber as String,posic_partnumber as String)
	Dim var as String
	
	If Right(TrimPosicionador(posic_partnumber),1) = "5" Or posic_partnumber.serie = "FY-303" Then
		var = "PROFIBUS PA"
	ElseIf Right(TrimPosicionador(posic_partnumber),1) = "0" Then
		var = "4-20mA"
	ElseIf Right(TrimPosicionador(posic_partnumber),1) = "1" Or posic_partnumber.serie = "FY-301" Then
		var = "4-20mA+Hart"
	Else
		var = "-"
	End If
	
	EntrPosicionador = var
End Sub