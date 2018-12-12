Sub FluidoEstado(item_partnumber as String)
	Dim var as String
	
	If IsNull(item_partnumber.fluidoestado) Then
		var = "-"
	Else
		var = item_partnumber.fluidoestado
	End If

	FluidoEstado = var
End Sub