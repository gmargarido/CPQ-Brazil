Sub FimDeCurso(item_partnumber as String)
	Dim var as String
	
	If IsNull(item_partnumber.fimdecurso) Then
		var = "NÃO"
	Else
		var = "Ver nota 2"
	End If

	FimDeCurso = var
End Sub