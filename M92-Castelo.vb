Sub Castelo(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Butterfly Valve" Then
		var = select "Norma das Conexões" from /* MATERIAL_BORBOLETA */ where Modelo = Serie(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Ball Valve"
		var = "Definir" 'Dejar el campo con esta mensaje como alerta para que el usuario deba escribir o definir
	Else
		var = "" 'Dejar el campo en blanco
	End If
	Castelo = var
End Sub