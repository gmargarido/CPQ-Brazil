Sub Passagem(item_partnumber as String)
	Dim var as String

	If CheckType(item_partnumber) = "Knife Valve" Then
		var = "PLENA"
	ElseIf CheckType(item_partnumber) = "Butterfly Valve" Then
		var = "N/A"
	ElseIf Serie(item_partnumber) = "19"
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		var = Select Passagem from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)
	ElseIf Serie(item_partnumber) = "1A"
		classe = Mid(item_partnumber,29,1)
		
		'Es necesario utilizar la query abajo en la tabla PASSAGEM_1A: Hoja "Dados" / Celdas DK1:DM25
		var = Select "PASSAGEM" from /* Tabla */ where CLASSE = classe
	ElseIf Serie(item_partnumber) = "17"
		'Es necesario utilizar la query abajo en la tabla ESPECIAL: Hoja "Dados" / Celdas BM1:BQ94
		var = Select "Feature 1" from /* Tabla */ where SPECIAL FEATURE NUMBER = PassagemEspecial(item_partnumber)
		
	ElseIf AcionamentoEsfera(item_partnumber) = "NN"
		'Es necesario utilizar la query abajo en la tabla Celdas B3:M943 
		var = Select Passagem from /* Tabla */ where Serie = Serie(item_partnumber) and Base = BaseNumber(item_partnumber)
	ElseIf AcionamentoEsfera(item_partnumber) = "NA"
		var = "V-PORT 30°"
	ElseIf AcionamentoEsfera(item_partnumber) = "NB"
		var = "V-PORT 60°"
	ElseIf AcionamentoEsfera(item_partnumber) = "NJ"	
		var = "V-PORT 15°"
	ElseIf AcionamentoEsfera(item_partnumber) = "NF"
		var = "V-PORT 90°"
	ElseIf AcionamentoEsfera(item_partnumber) = "ND"
		var = "SLOTTED"
	End If
	
	Passagem = var
End Sub

Public Function PassagemEspecial(item_partnumber as String)
	Dim var as String
	
	If Serie(item_partnumber) = "19"
		var = Mid(item_partnumber,14,4)
	ElseIf Serie(item_partnumber) = "1A"
		var = Mid(item_partnumber,25,4)
	Else
		var = Mid(item_partnumber,21,4)
	End If
	
	PassagemEspecial = var
End Function

Public Function AcionamentoEsfera(item_partnumber as String)
	Dim var as String
	
	If Serie(item_partnumber) <> "1A"
		var = Mid(item_partnumber,18,2)
	End If
	
	Acionamento = var
End Function