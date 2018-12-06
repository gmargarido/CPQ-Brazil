Sub Porcas(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		var = "N/A"
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		var = "N/A"
	ElseIf Serie(item_partnumber) = "19"
		var = "N/A"
	ElseIf Serie(item_partnumber) = "1A"
		'Es necesario utilizar la query abajo en la tabla CORPO_1A - Celdas CT1:CW25
		var = Select "Porca" from /* CORPO_1A */ where Corpo = TrimCorpo(item_partnumber)
	ElseIf Model(item_partnumber) = "S8000"
		var = "INOX"
	Else
		'Es necesario utilizar la query abajo en la tabla CORPO - Celdas R1:U13
		var = Select "Porca" from /* CORPO */ where Cód = TrimCorpo(item_partnumber)
	EndIf
	
	Parafusos = var
End Sub