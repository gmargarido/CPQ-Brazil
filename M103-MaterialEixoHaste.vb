Sub MaterialEixoHaste(item_partnumber as String, actuator_partnumber as String)
	Dim var as String
	Dim Atuador as String
	
	Atuador = TrimAtuador(item_partnumber, actuator_partnumber)
	
	If CheckType(item_partnumber) =  "Knife Valve" Then
		If Atuador = ""
			var = "DEFINIR"
		ElseIf Left(Atuador,1) = "C"
			'Es necesario utilizar la query abajo en la tabla HASTE_GUILHOTINA - Celdas IV1:IZ91
			var = Select "Material Haste" from /* HASTE_GUILHOTINA */ where Code = Mid(Atuador,2,3)
		Else
			'Es necesario utilizar la query abajo en la tabla MAT_ACIONAMENTO_GUILHOTINA - Celdas IN1:IT13
			var = Select "Haste" from /* MAT_ACIONAMENTO_GUILHOTINA */ where Trim Code = Mid(Atuador,2,3)
		EndIf
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		'Es necesario utilizar la query abajo en la tabla TRIM_BORBOLETA - Celdas FJ1:FP3478
		var = Select "Haste" from /* TRIM_BORBOLETA */ where Trim = Internos(item_partnumber)
	ElseIf CheckType(item_partnumber) = "Ball Valve" Then
		If Serie(item_partnumber) = "1A"
			'Tabla HASTE_1A - Oja "Dados" > CZ1:CZ25
			var = Select "Material" from /* SEDE_1A */ where SEDE = TrimHaste(item_partnumber)
		Else
			'Tabla ESFERA - Oja "Dados" > AA1:AC20
			var = Select "HASTE" from /* ESFERA_1A */ where Cód = TrimHaste(item_partnumber)
		End If
	End If
	
	MaterialEixoHaste = var
End Sub

Public Function TrimAtuador(item_partnumber as String, actuator_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) =  "Knife Valve" Then
		If actuator_partnumber = "" Then
			var = Mid(item_partnumber,28,5)
		Else
			var = Mid(actuator_partnumber,14,5)
		End If
	ElseIf CheckType(item_partnumber) =  "Butterfly Valve" Then	
		var = Right(item_partnumber,8)
	Else
		var = Mid(actuator_partnumber,3,3)
	End If
	
	TrimAtuador = var
End Function