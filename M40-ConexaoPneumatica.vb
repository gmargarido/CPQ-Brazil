Sub ConexaoAtuador(item_partnumber as String,actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)
	
	If CheckType(item_partnumber) =  "Knife Valve" Then
		If Left(Atuador,1) = "C"
			'Es necesario utilizar la query abajo en la tabla HASTE_GUILHOTINA - Celdas IV1:IZ91
			conexao = Select "Con. Pneum." from /* HASTE_GUILHOTINA */ where Code = Mid(Atuador,2,3)

			var = conexao + " NPT"
		Else
			var = "-"
		End If
	Else
		var = ""
	End If
	
	ConexaoAtuador = var
End Sub