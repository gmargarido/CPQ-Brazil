
Sub FechaCom(item_partnumber as String,actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)
	
	If CheckType(item_partnumber) = "Ball Valve" Then
		If FalhaAtuador(item_partnumber, actuator_partnumber) = "-" Then
			var = "-"
		ElseIf FalhaAtuador(item_partnumber, actuator_partnumber) = "ÚLTIMA POSIÇÃO" Or FalhaAtuador(item_partnumber, actuator_partnumber) = "ABERTA"
			var = AlimentAtuador(item_partnumber,actuator_partnumber)
		Else
			var = "MOLA"
		End If
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Or  CheckType(item_partnumber) = "Knife Valve"
		If FalhaAtuador(item_partnumber, actuator_partnumber) = "ÚLTIMA POSIÇÃO" Or  FalhaAtuador(item_partnumber, actuator_partnumber) = "ABERTA"
			var = AlimentAtuador(item_partnumber,actuator_partnumber)
		ElseIf FalhaAtuador(item_partnumber, actuator_partnumber) = "FECHADA"
			var = "MOLA"
		Else
			var = "-"
		End If
	Else
		var = ""
	End If
	
	FechaCom = var

End Sub