Sub FiltroReguladorManom(item_partnumber as String)
	Dim var as String

	If IsNull(item_partnumber.filtro) Then
		var = "NÃO"
	Else
		var = "SIM"
	End If
	
	FiltroReguladorManom = var
End Sub