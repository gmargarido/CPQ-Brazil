Sub SentidoFluxo(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		If Left(Serie(item_partnumber),1) = "9"
			var = "UNIDIRECIONAL"
		Else
			var = "BIDIRECIONAL"
		End If
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		var = "BIDIRECIONAL"
	Else
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		var = Select Sentido from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)
	End If
	
	SentidoFluxo = var
End Sub