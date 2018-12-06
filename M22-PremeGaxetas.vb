Sub PremeGaxetas(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) =  "Knife Valv" Then
		'Es necesario utilizar la query abajo en la tabla CORPO_GUILHOTINA - Celdas GX1:HD137
		var = Select "PremeGaxeta" from /* CORPO_GUILHOTINA */ where Trim Code = Corpo(item_partnumber)
	Else
		var = "-"
	End If
		
	PremeGaxetas = var 
End Sub