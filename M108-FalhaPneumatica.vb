Sub FalhaPneumatica(item_partnumber as String, actuator_partnumber as String) 
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)

	If CheckType(item_partnumber) =  "Knife Valve" Then
		If AcaoAtuador(item_partnumber, actuator_partnumber) = "DUPLA AÇÃO" Then
			var = "ÚLTIMA POSIÇÃO"
		Else
			var = "-"
		End If
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		If ModeloAtuador(item_partnumber, actuator_partnumber) = "S92" Then
			var = "ÚLTIMA POSIÇÃO" 
		ElseIf FailPosition = "FF" Then 'El campo FailPosition aún no se ha creado. Daniel enviará un correo para hablar de esto.
			var = "FECHADA" 
		ElseIf FailPosition = "FA" Then 'El campo FailPosition aún no se ha creado. Daniel enviará un correo para hablar de esto.
			var = "ABERTA"
		Else
			var = "-"
		End If		
	ElseIf CheckType(item_partnumber) = "Ball Valve" Then
		If Left(Atuador,2) = "92" Then
			var = "ÚLTIMA POSIÇÃO" 
		ElseIf Right(Atuador,2) = "FF"
			var = "FECHADA" 
		ElseIf Right(Atuador,2) = "FA"
			var = "ABERTA" 
		Else
			var = "-"
		EndIf
	EndIf

	FalhaPneumatica = var
End Sub