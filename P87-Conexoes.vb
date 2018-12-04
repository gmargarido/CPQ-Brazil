Sub Conexoes(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		'Tabla MATERIAL_GUILHOTINA - Oja "Dados" > HN1:HR20
		var = Select "Conexão" from /* MATERIAL_GUILHOTINA */ where Modelo = Serie(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		'Tabla MATERIAL_BORBOLETA - Oja "Dados" > FR1:GA37
		var = Select "Conexão" from /* MATERIAL_BORBOLETA */ where Modelo = Serie(item_partnumber)
	Else
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		var = Select "Conexoes" from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)	
	End If
	
	Conexoes = Var
		

End Sub