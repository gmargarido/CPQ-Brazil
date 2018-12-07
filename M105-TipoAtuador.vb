Sub TipoAtuador(item_partnumber as String,actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)

	If CheckType(item_partnumber) =  "Knife Valve" Then
		'Es necesario utilizar la query abajo en la tabla ATUADOR_GUILHOTINA - Celdas HF1:HI12
		var = Select "Descrição" from /* ATUADOR_GUILHOTINA */ where Código = Left(Atuador,1)
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		If AtuadorAcion(Atuador) = "BFA" Or Left(Atuador,2) = "01"
			var = "ALAVANCA"
		ElseIf AtuadorAcion(Atuador) = "BFC" Or Left(Atuador,2) = "04"  Then
			var = "CAIXA DE REDUÇÃO"
		ElseIf CaracteristicaAbertura(item_partnumber) = "ON-OFF" Or CaracteristicaAbertura(item_partnumber) = "=%"  Then
			var = "PINHÃO CREMALHEIRA"
		Else
			var = "-"
		End If

	ElseIf CheckType(item_partnumber) = "Ball Valve" Then
		If Left(actuator_partnumber,2) = "92" Or Left(actuator_partnumber,2) = "93"
			var = "PINHÃO CREMALHEIRA"
		ElseIf actuator_partnumber = ""
			var = ""
		Else
			var = "-"
		End If
	
	End If

	TipoAtuador = var
End Sub

Public Function AtuadorAcion(Atuador as String)
	Dim var as String 
	
	If item_partnumber.accionamiento = "Manual"
		var = Left(Atuador,3)
	Else
		var = item_partnumber.accionamiento
	End If
		
	AtuadorAcion = var
End Function