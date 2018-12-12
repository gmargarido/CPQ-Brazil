Sub ValvSolenoide(item_partnumber as String)
	Dim var as String
	
	If IsNull(item_partnumber.solenoide) Then
		var = "NÃO"
	Else
		var = "Ver nota 1"
	End If

	ValvSolenoide = var
End Sub