Sub MaterialAnelDisco(item_partnumber as String)
	Dim var as String

	If Model(item_partnumber) = "Trilok"
		'Es necesario utilizar la query abajo en la tabla TRIM_BORBOLETA - Celdas FJ1:FP3478
		var = Select "Disc Seal (TriLok)" from /* TRIM_BORBOLETA */ where Trim = Internos(item_partnumber)
	Else
		var = "N/A"
	End If

	MaterialAnelDisco = var
End Sub