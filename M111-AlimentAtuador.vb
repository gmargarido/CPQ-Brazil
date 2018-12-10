Sub AlimentAtuador(item_partnumber as String,actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)
	
	If CheckType(item_partnumber) =  "Knife Valve" Then
		If Left(Atuador,1) = "C"
			var = "80 PSI" 'Este campo aún no se ha creado, pero el valor puede ser fijo con "80 PSI"
		Else
			var = "-"
		End If
	Else 
		If actuator_partnumber = "" Then
			var = ""
		ElseIf item_partnumber.accionamiento = "MANUAL" Then
			var = "-"
		Else
			var = "80 PSI" 'Este campo aún no se ha creado, pero el valor puede ser fijo con "80 PSI"
		End If	
	End If
End Sub