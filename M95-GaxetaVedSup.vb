Sub GaxetaVedSup(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		'Es necesario utilizar la query abajo en la tabla CORPO_GUILHOTINA - Celdas GX1:HD137
		var = Select "Gaxeta" from /* CORPO_GUILHOTINA */ where Trim Code = Corpo(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		'Tabla MATERIAL_BORBOLETA - Oja "Dados" > FR1:GA37
		var = Select "Gaxeta" from /* MATERIAL_BORBOLETA */ where Modelo = Serie(item_partnumber)
	ElseIf Serie(item_partnumber) = "19"
		'Tabla INTERNOS_S19 - Oja "Dados" > EL1:EW26
		var = Select "Gasket" from /* INTERNOS_S19 */ where = Mid(Internos(item_partnumber),2,3)
	ElseIf Serie(item_partnumber) = "1A"
		'Tabla GAXETA_1A - Oja "Dados" > EL1:EW26
		var = Select "Material" from /* GAXETA_1A */ where Gaxeta = GasketMaterial(item_partnumber)
	Else	
		'Tabla GAXETA - Oja "Dados" > AI1:AJ7
		var = Select "Gaxeta" from /* GAXETA */ where Cód = GasketMaterial(item_partnumber)	
	End If

	GaxetaVedSup = var 
End Sub