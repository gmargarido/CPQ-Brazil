Sub MaterialSede(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		'Es necesario utilizar la query abajo en la tabla CORPO_GUILHOTINA - Celdas GX1:HD137
		var = Select "Sede" from /* CORPO_GUILHOTINA */ where Trim Code = TrimCorpo(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		'Es necesario utilizar la query abajo en la tabla TRIM_BORBOLETA - Celdas FJ1:FP3478
		var = Select "Sede" from /* TRIM_BORBOLETA */ where Trim = Internos(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Ball Valve" Then
		If Serie(item_partnumber) = "19"
			'Tabla INTERNOS_S19 - Oja "Dados" > EL1:EW26
			var = Select "Seat" from /* INTERNOS_S19 */ where TRIM S19 = Mid(Internos(item_partnumber),2,3)
		ElseIf Serie(item_partnumber) = "1A"
			'Tabla SEDE_1A - Oja "Dados" > DB1:DC25
			a = Select "Material" from /* SEDE_1A */ where SEDE = TrimSede(item_partnumber)
			
			'Tabla INSERTO_SEDE_1A	 - Oja "Dados" > DD1:DF25
			b = Select "Vedação" from /* INSERTO_SEDE_1A */ where INSERTO SEDE = Mid(item_partnumber,17,1)
			
			var = a + " / " + b
		Else
			'Tabla ESFERA - Oja "Dados" > AE1:AG16
			var = Select "SEDES" from /* SEDE */ where TrimSede(item_partnumber)
		EndIf
	End If

	MaterialSede = var
End Sub