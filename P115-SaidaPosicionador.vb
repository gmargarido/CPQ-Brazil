Sub SaidaPosicionador(item_partnumber as String, posic_partnumber as String)
	Dim var as String
	
	If EntrPosicionador(item_partnumber,posic_partnumber) = "-" Then
		var = "-"
	Else
		var = "0~" + AlimentAtuador(item_partnumber,actuator_partnumber)
	End If
	
	SaidaPosicionador = var
End Sub
