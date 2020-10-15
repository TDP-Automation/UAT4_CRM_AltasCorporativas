Call Log_Out()

Sub Log_Out()
	
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
			wait 3
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Click
		End If @@ hightlight id_;_30463552_;_script infofile_;_ZIP::ssf1.xml_;_
		
		If JavaWindow("Ejecutivo de interacción").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").JavaMenu("Archivo").JavaMenu("Salida").Select
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").Exist(1)  Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").JavaButton("Sí").Click
			End If
		End If
End Sub

'Call FinExe()
'Call ExportExcel()

