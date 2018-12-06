Sub Internos(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) =  "Knife Valv" Then
		'No se aplica a la valvula de tipo Knife.
		var = ""
	ElseIf CheckType(item_partnumber) =  "Check Valv" Then
		If Len(item_partnumber) = 18
			var = Mid(item_partnumber,12,3)
		ElseIf Len(item_partnumber) = 19
			If Right(item_partnumber,4) = "CAB2"
				var = Mid(item_partnumber,12,3)
			Else
				var = Mid(item_partnumber,13,3)
			End If
		ElseIf Len(item_partnumber) = 20
			var = Mid(item_partnumber,13,3)
		End If
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		var = Right(item_partnumber,3)
	Else
		If Serie(item_partnumber) = "19"
			var = TrimCorpo(item_partnumber) + TrimEsfera(item_partnumber) + TrimHaste(item_partnumber) + TrimSede(item_partnumber)
		ElseIf Serie(item_partnumber) = "1A"
			var = TrimCorpo(item_partnumber) + TrimEsfera(item_partnumber) + TrimHaste(item_partnumber) + TrimSede(item_partnumber) + TrimVedacao(item_partnumber) + TrimGaxeta(item_partnumber) + TrimORing(item_partnumber)
		Else
			var = TrimCorpo(item_partnumber) + TrimEsfera(item_partnumber) + TrimSede(item_partnumber) + TrimGaxeta(item_partnumber)
		EndIf
	End If
	
	Internos = var
End Sub

Public Function TrimEsfera(item_partnumber as String)
	Dim var as String
	
	var = Mid(item_partnumber,14,1)

	TrimEsfera = var
End Function

Public Function TrimHaste(item_partnumber as String)
	Dim var as String
	
	If Serie(item_partnumber) = "19" Or Serie(item_partnumber) = "1A"
		var = Mid(item_partnumber,15,1)
	Else
		var = Mid(item_partnumber,14,1)
	End If

	TrimHaste = var
End Function

Public Function TrimSede(item_partnumber as String)
	Dim var as String
	
	If Serie(item_partnumber) = "19" Or Serie(item_partnumber) = "1A"
		var = Mid(item_partnumber,16,1)
	Else
		var = Mid(item_partnumber,15,1)
	End If

	TrimSede = var
End Function

Public Function TrimVedacao(item_partnumber as String)
	Dim var as String
	
	If Serie(item_partnumber) = "1A"
		var = Mid(item_partnumber,17,1)
	Else
		var = 0
	End If

	TrimEsfera = var
End Function

Public Function TrimGaxeta(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) =  "Butterfly Valv" Then
		var = Select Gaxeta from MATERIAL_BORBOLETA (Oja Dados - FR1:GA37) where Mod = Serie(item_partnumber)
	ElseIf Serie(item_partnumber) = "1A"
		var = Mid(item_partnumber,18,1)
	Else
		var = Mid(item_partnumber,16,1)
	End If
	
	TrimGaxeta = var
End Function

Public Function TrimORing(item_partnumber as String)
	Dim var as String
	
	If Serie(item_partnumber) = "1A"
		var = Mid(item_partnumber,19,1)
	Else
		var = 0
	End If

	TrimORing = var
End Function