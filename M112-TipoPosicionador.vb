Sub TipoPosicionador(item_partnumber as String, posic_partnumber as String)
	Dim var as String
	
	If posic_partnumber <> "" Then
		If CaracteristicaAbertura(item_partnumber) = "%=" Then
			var = "ELETROPNEUMÁTICO"
		Else
			var = "-"
		End If
	Else
		var = "-"
	End If
	
	TipoPosicionador = var
End Sub

Public Function TrimPosicionador(posic_partnumber as String)
	Dim var as String
	
	var = Left(posic_partnumber,8)
	
	TrimPosicionador = var
End Function