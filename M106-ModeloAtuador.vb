Sub ModeloAtuador(item_partnumber as String, actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)

	If CheckType(item_partnumber) =  "Knife Valve" Then
		If Left(Atuador,1) = "C" Then
			If SentidoFluxo(item_partnumber) = "UNIDIRECIONAL" Then
				var = "CCGC " + DiamAtuador(item_partnumber, actuator_partnumber)
			Else
				var = "CCGC " + DiamAtuador(item_partnumber, actuator_partnumber)
			End If
		Else
			'Es necesario utilizar la query abajo en la tabla ATUADOR_GUILHOTINA - Celdas HF1:HI12
			var = Select "Nome" from /* ATUADOR_GUILHOTINA */ where Código = Left(Atuador,1)
		End If
	
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		If TipoAtuador(item_partnumber,actuator_partnumber) = "ALAVANCA" Then
			var = "S01"
		ElseIf TipoAtuador(item_partnumber,actuator_partnumber) = "CAIXA DE REDUÇÃO" Then
			var = "S04"
		ElseIf CaracteristicaAbertura(item_partnumber) = "ON-OFF" Or CaracteristicaAbertura(item_partnumber) = "=%"  Then
			var = "S" + Left(actuator_partnumber,2)
		Else
			var = "-"
		End If
	ElseIf CheckType(item_partnumber) = "Ball Valve" Then
		If Left(actuator_partnumber,2) = "93"  Then
			var = "S" + Left(actuator_partnumber,2) + "/" + Atuador + "-" + Mid(actuator_partnumber,6,1)
		Else
			var = "S" + Left(actuator_partnumber,2) + "/" + Atuador
		End If
	End If
	
	ModeloAtuador = var
End Sub

Public Function DiamAtuador(item_partnumber as String, actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)

	If CheckType(item_partnumber) =  "Knife Valve" Then	
		'Es necesario utilizar la query abajo en la tabla HASTE_GUILHOTINA - Celdas IV1:IZ91
		var = Select "Tamanho" from /* HASTE_GUILHOTINA */ where Code = Mid(Atuador,2,3)
	End If
	DiamAtuador = var
End Function