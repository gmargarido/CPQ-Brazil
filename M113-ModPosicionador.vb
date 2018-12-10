Sub ModeloPosicionador(item_partnumber as String, posic_partnumber as String)
	Dim var as String

	If TipoPosicionador(item_partnumber, posic_partnumber) = "ELETROPNEUMÁTICO" Then
		var = "S"+Left(posic_partnumber,2)
	Else
		var = "-"	
	End If

	ModeloPosicionador = var
End Sub