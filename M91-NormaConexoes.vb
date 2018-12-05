Sub NormaConexoes(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		'Es necesario utilizar la query abajo en la tabla CORPO2_GUILHOTINA - Celdas IA1:IF44
		var = Select "Norma de Furacao" from /* CORPO2_GUILHOTINA */ where Trim Code = Corpo(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		If Diametro(item_partnumber) > "2400"
			var = select "Norma das Conexões > DN 24" from /* MATERIAL_BORBOLETA */ where Modelo = Serie(item_partnumber)
		Else
			var = var = select "Norma das Conexões" from /* MATERIAL_BORBOLETA */ where Modelo = Serie(item_partnumber)
		End If
	Else
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		var = Select "Norma" from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)	
	End If	
	
	NormaConexoes = var
End Sub