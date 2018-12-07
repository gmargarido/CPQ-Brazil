Sub CaracteristicaAbertura(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		'Luis, estoy llamando este campo Accionamiento dentro del objeto item_partnumber.
		var = item_partnumber.accionamiento
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Or CheckType(item_partnumber) = "Ball Valve" Then
		If item_partnumber.accionamiento = "CONTROLE"
			var = "=%"
		Else	
			var = item_partnumber.accionamiento
		End If
	ElseIf CheckType(item_partnumber) = "Check Valve" Then
		If Right(Passagem(item_partnumber),3) = "15º" Or Right(Passagem(item_partnumber),3) = "30º" Or Passagem(item_partnumber)= "SLOTTED" Then
			var = "LINEAR"
		ElseIf Right(Passagem(item_partnumber),3) = "60º" Or Right(Passagem(item_partnumber),3) = "90º" Or Passagem(item_partnumber)= "PORT" Then 
			var = "=%"
		ElseIf Passagem(item_partnumber)= "PLENA" Or Passagem(item_partnumber)= "REDUZIDA" Then 
			var = "ABERTURA RÁPIDA"
		Else
			var = "DEFINIR"
		End If
	Else
		var = ""
	EndIf
	
	CaracteristicaAbertura = var 
End Sub