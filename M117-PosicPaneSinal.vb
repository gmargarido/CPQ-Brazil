Sub PosicPaneSinal(item_partnumber as String,posic_partnumber as String)
	Dim var as String
	
	If (CaracteristicaAbertura(item_partnumber) = "%=" And FalhaEletrAtuador(item_partnumber, actuator_partnumber) = "ÚLTIMA POSIÇÃO" 
		Or (CaracteristicaAbertura(item_partnumber) = "%=" And FalhaEletrAtuador(item_partnumber, actuator_partnumber) = "FECHADA" Then
		var = "FECHADA"
	ElseIf (CaracteristicaAbertura(item_partnumber) = "%=" And FalhaEletrAtuador(item_partnumber, actuator_partnumber) = "ABERTA" Then
		var = "ABERTA"
	Else
		var = "-"
	End
	PosicPaneSinal = var
End Sub