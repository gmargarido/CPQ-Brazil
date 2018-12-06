Sub MaterialObtur(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) =  "Knife Valve" Then
		'Es necesario utilizar la query abajo en la tabla CORPO_GUILHOTINA - Celdas GX1:HD137
		var = Select "Faca" from /* CORPO_GUILHOTINA */ where Trim Code = TrimCorpo(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		'Es necesario utilizar la query abajo en la tabla TRIM_BORBOLETA - Celdas FJ1:FP3478
		var = Select "Disco" from /* TRIM_BORBOLETA */ where Trim = Internos(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Ball Valve" Then
		If Serie(item_partnumber) = "19"
			'Tabla INTERNOS_S19 - Oja "Dados" > EL1:EW26
			var = Select "Segment" from /* INTERNOS_S19 */ where TRIM S19 = Mid(Internos(item_partnumber),2,3)
		ElseIf Serie(item_partnumber) = "1A"
			'Tabla ESFERA_1A - Oja "Dados" > CX1:CY25
			var = Select "Material" from /* ESFERA_1A */ where ESFERA = TrimEsfera(item_partnumber)
		Else
			'Tabla ESFERA - Oja "Dados" > AA1:AC20
			var = Select "ESFERA" from /* ESFERA_1A */ where Cód = TrimEsfera(item_partnumber)
		EndIf
	End If
	
	MaterialObtur = var
End Sub