Sub Diametro(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		var = Mid(item_partnumber,6,4)
	ElseIf CheckType(item_partnumber) = "Check Valve" Then
		var = Mid(item_partnumber,6,2)
	Else
		var = Mid(item_partnumber,3,4)
	End If
	
	Diametro = var
End Sub