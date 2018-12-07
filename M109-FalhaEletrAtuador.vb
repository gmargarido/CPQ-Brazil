Sub FalhaEletrAtuador(item_partnumber as String,actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)
	
	If TemSolenoide = True Then
		If AcaoAtuador(item_partnumber,actuator_partnumber) = "DUPLA AÇÃO" Then
			var = "FECHADA"
		Else
			var = FalhaAtuador(item_partnumber, actuator_partnumber) 
		End If
	Else
		var = "-"
	End
	
	FalhaEletrAtuador = var
End Sub