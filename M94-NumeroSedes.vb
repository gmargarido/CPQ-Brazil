Sub NumeroSedes(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then	
		'En desarollo
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		var = "01"
	Else
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		var =  Select Sedes from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)
	End If
	
	NumeroSedes = var
End Sub