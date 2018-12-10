Sub CaracPosicionador(item_partnumber as String, posic_partnumber as String)
	Dim var as String
	
	If TipoPosicionador(item_partnumber, posic_partnumber) = "ELETROPNEUMÁTICO" Then
		var = "LINEAR"
	Else
		var = "-"	
	End If
	
	CaracPosicionador = var
End Sub
