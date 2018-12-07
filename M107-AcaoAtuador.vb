Sub AcaoAtuador(item_partnumber as String,actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)

	If CheckType(item_partnumber) =  "Knife Valve" Then
		If Left(Atuador,1) = "C" Then
			var = "DUPLA AÇÃO"
		Else
			var = "-"
		End If
	Else
		If Left(actuator_partnumber,2) = "93"  Then
			var = "SIMPLES AÇÃO"
		ElseIf Left(actuator_partnumber,2) = "92"  Then
			var = "DUPLA AÇÃO"
		Else
			var = "-"
		End If
	End If

	AcaoAtuador = var
End Sub