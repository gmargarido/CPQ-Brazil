Sub TipoHaste(item_partnumber as String)
	Dim var as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)
	
	If CheckType(item_partnumber) =  "Knife Valve" Then
		If Atuador <> ""
			var = "ASCENDENTE"
		End If
	Else
		'No se aplica a los otros tipos de valvula.
		var = "-"
	End If
	
	TipoHaste = var
End Sub