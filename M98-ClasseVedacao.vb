Sub ClasseVedacao(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		var = "ESTANQUE"
	ElseIf Serie(item_partnumber) = "19"
		'Tabla INTERNOS_S19 - Oja "Dados" > EL1:EW26
		var = Select "Vedação" from /* INTERNOS_S19 */ where TRIM S19 = Mid(Internos(item_partnumber),2,3)
	ElseIf Serie(item_partnumber) = "1A"
		'Tabla INSERTO_SEDE_1A	 - Oja "Dados" > DD1:DF25
		var = Select "Vedação" from /* INSERTO_SEDE_1A */ where INSERTO SEDE = Mid(item_partnumber,17,1)
	Else
		'Tabla SEDE - Oja "Dados" > AE1:AG16
		var = Select "Vedação" from /* SEDE */ where Cód = SeatMaterial(item_partnumber)
	End If

	ClasseVedacao = var
End Sub