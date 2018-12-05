Sub MaterialCorpo(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		'Es necesario utilizar la query abajo en la tabla CORPO_GUILHOTINA - Celdas GX1:HD137
		var = Select "Corpo" from /* CORPO_GUILHOTINA */ where Trim Code = Corpo(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		'Es necesario utilizar la query abajo en la tabla TRIM_BORBOLETA - Celdas FJ1:FP3478
		var = Select "Corpo" from /* TRIM_BORBOLETA */ where Trim = Internos(item_partnumber)
	ElseIf Serie(item_partnumber) = "19"
		'Es necesario utilizar la query abajo en la tabla CORPO_S19 - Celdas EY1:FA28
		var = Select "Material" from /* CORPO_S19 */ where Corpo = Corpo(item_partnumber)
	ElseIf Serie(item_partnumber) = "1A"
		'Es necesario utilizar la query abajo en la tabla CORPO_1A - Celdas CT1:CW25
		var = Select "Material" from /* CORPO_1A */ where Corpo = Corpo(item_partnumber)
	Else
		'Es necesario utilizar la query abajo en la tabla CORPO - Celdas R1:U13
		var = Select "Corpo" from /* CORPO */ where Cód = Corpo(item_partnumber)
	End If
	
	MaterialCorpo = var
End Sub