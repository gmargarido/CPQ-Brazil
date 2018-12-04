Sub PressureClass(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		var = Mid(item_partnumber,11,2) + "BAR"
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		If (("S"+Serie(item_partnumber) = "S20" Or "S"+Serie(item_partnumber) = "S21") And Internos(item_partnumber) = "127") Or (("S"+Serie(item_partnumber) = "S20" Or "S"+Serie(item_partnumber) = "S21") And Internos(item_partnumber) = "327") Then
			var = "100 PSI"
		ElseIf Model(item_partnumber) = "S30" Or Model(item_partnumber) = "S31" Or Model(item_partnumber) = "S3A" Then
			var = "150 PSI"
		Else
			var = select "Pressão Máxima" from /* MATERIAL_BORBOLETA */ where Modelo = Serie(item_partnumber)
		End If
	ElseIf Serie(item_partnumber) = "19"
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		var = Select "Classe" from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)
	ElseIf Serie(item_partnumber) = "1A"
		classe = Mid(item_partnumber,29,1)
		
		'Es necesario utilizar la query abajo en la tabla CLASSE_1A: Hoja "Dados" / Celdas DK1:DL13
		var = Select "ASME" from /* Tabla */ where CLASSE = classe
	Else
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		var = Select "Classe" from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)
	End If
End Sub