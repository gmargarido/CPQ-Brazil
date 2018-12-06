Sub FireTested(item_partnumber as String)
	Dim var as String
	
	If Internos(item_partnumber) = "068" Or Internos(item_partnumber) = "068" Or Model(item_partnumber) = "Trilok"
		var = "SIM"
	Else
		var = "NÃO"
	End If

	FireTested = var
End Sub