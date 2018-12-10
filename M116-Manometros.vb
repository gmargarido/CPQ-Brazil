Sub Manometros(item_partnumber as String,actuator_partnumber as String)
	Dim var as String
	
	If Right(TrimPosicionador(posic_partnumber),1) = "3" Or Right(TrimPosicionador(posic_partnumber),1) = "4"
		Or posic_partnumber.serie = "FY-301" Or posic_partnumber.serie = "FY-303" Or posic_partnumber.serie = "FY-400" Then
		var = "SIM"
	Else
		var = "-"
	End If
	
	Manometros = var
End Sub