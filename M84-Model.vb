Sub Model(item_partnumber as String)
	Dim var as String

	If CheckType(item_partnumber) = "Knife Valve" Then
		var = "FIG" + Mid(item_partnumber,2,3)
	ElseIf CheckType(item_partnumber) = "Check Valve" Then
		If Right(item_partnumber,4) = "CAB2" Then
			var = "210"
		Else
			var = Right(item_partnumber,3)
		End If
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		trilok = select TRILOK from /* MATERIAL_BORBOLETA */ where Modelo = Serie(item_partnumber)
		
		If trilok = true Then
			var = "Trilok"
		Else
			var = "S" + Serie(item_partnumber)
		End If
	Else
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		modelo = Select MODELO from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)
		If modelo = "S7000/S8000"
			If TrimCorpo(item_partnumber) = "C"
				var = "S8000"
			Else
				var = "S7000"
			End If
		End If
	End If
	
	Model = var
End Sub

Public Function CheckType(item_partnumber as String)
	Dim var as String
	Dim item_NCM as String
	
	'En el cadastro de productos hay un campo con NCM [El código unico que utilizaremos para buscar el "Tipo" (Type) ]
	item_NCM = Select NCM from /* Productos */ where codigo = item_partnumber
	
	'La tabla abajo esta en la hoja "Impostos"
	var = Select Tipo from /* NCM */ where NCM = item_NCM
	
End Function

Public Function BaseNumber(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		var = Mid(item_partnumber,14,7)
	Else
		var = Mid(item_partnumber,8,5)
	End If
	
	BaseNumber = var
End Function

Public Function IndexBaseNumber(BaseNumber as String)
	Dim idx as String
	
	idx = 'Select Base from Dados!A3:A130 where Base = BaseNumber 
End Function

Public Function IndexSerie(Serie as String)
	Dim idx as String
	
	idx = 'Select Serie from Dados!A3:A130 where Serie = Serie
End Function

Public Function Serie(item_partnumber as String)
	Dim var as String
	
	If CheckType(item_partnumber) = "Knife Valve" Then
		var = Mid(item_partnumber,2,3)
	Else
		If BaseNumber(item_partnumber) = "13000" 'Sub Function
			var = Mid(item_partnumber,1,2) + "H"
		ElseIf Right(item_partnumber,3) = "3JY" Or Right(item_partnumber,3) = "3RL"
			var = Mid(item_partnumber,1,2) + "R"
		Else
			var = Mid(item_partnumber,1,2)
		End If
	End If
	
	Serie = var
	
End Function

Public Function TrimCorpo(item_partnumber as String)
	Dim var as String
	
	If Item = "Knife Valve" Then
		var = Mid(item_partnumber,15,3)
	Else
		var = Mid(item_partnumber,13,1)
	End If
	
	Corpo = var

End Function

Public Function Construction(item_partnumber as String)
	'Tabla en la hoja "Dados" - Celdas B3:M943
	'Utilizada para las valvulas Ball y Check
	Dim var as String
	
	'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
	var = Select MODELO from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)
	
	Construction = var
End



