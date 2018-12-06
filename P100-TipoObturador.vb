Sub TipoObturador(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) =  "Knife Valve" Then
		'Tabla MATERIAL_GUILHOTINA - Oja "Dados" > HN1:HR20
		var = Select "Tipo Faca" from /* MATERIAL_GUILHOTINA */ where Modelo = Serie(item_partnumber)
	Else
		'No se aplica a los otros tipos de valvula.
		var = "-"
	End If
	
	TipoObturador = var
End Sub