Sub MaterialGuia(item_partnumber as String)
	Dim var as String

	If CheckType(item_partnumber) =  "Knife Valve" Or CheckType(item_partnumber) = "Ball Valve" Then
		var = "N/A"
		
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		'Tabla MATERIAL_BORBOLETA - Oja "Dados" > FR1:GA37
		var = Select "Material da Guia" from /* MATERIAL_BORBOLETA */ where Modelo = Serie(item_partnumber)
	End If
	
	MaterialGuia = var
End Sub