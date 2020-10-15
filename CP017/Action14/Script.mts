

If ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "POSTPAGO" Then
	Call Potspago()
	ElseIf ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "PREPAGO" Then
	Call Prepago()
	ElseIf ucase(DataTable("Tipo_Plan", dtLocalSheet)) = "PORTABILIDAD" Then
	Call Portabilidad()
End If



Sub Potspago()
While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = false
	wait 1
Wend
wait 3
 Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
 wait 3
 While Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist = false
 	wait 1
 Wend
 wait 3
 Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Click
 wait 3
While  (Window("Ejecutivo de interacción").InsightObject("InsightObject_16").Exist or Window("Ejecutivo de interacción").InsightObject("InsightObject_24").Exist) = false
	wait 1
Wend
wait 3
If Window("Ejecutivo de interacción").InsightObject("InsightObject_24").Exist = true Then
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorServicio.png", True
	imagenToWord "Error al Consultar Servicio Web", RutaEvidencias() & "ErrorServicio.png"
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_26").Click
	wait 3
While JavaWindow("Ejecutivo de interacción").JavaDialog("Autenticación del Cliente").JavaDialog("Error Message").Exist= False
	wait 1
Wend
wait 3
JavaWindow("Ejecutivo de interacción").JavaDialog("Autenticación del Cliente").JavaDialog("Error Message").JavaButton("Cancelar").Click
DataTable("s_Resultado","A_StandAlone") ="Fallido"
DataTable("s_Detalle","A_StandAlone") ="Error al consultar servicio web"
Reporter.ReportEvent micFail, DataTable("s_Resultado","A_StandAlone"), DataTable("s_Detalle","A_StandAlone")
ExitTestIteration
End If
Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ScoreCalculado.png", True
imagenToWord "Score Calculado", RutaEvidencias() & "ScoreCalculado.png"
Set shell = CreateObject("Wscript.Shell") 
shell.SendKeys "{PGDN}"
wait 3
If Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist =True Then
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Click
End If 
wait 3
Set shell = CreateObject("Wscript.Shell") 
shell.SendKeys "{PGDN}"
If Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist = True Then
	wait 3
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Exist = true Then
		Call Validacion()
    End If
    wait 3
   	Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Click
End If
wait 5
Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidacionDatos.png", True
imagenToWord "Validación de Datos Exitosa", RutaEvidencias() & "ValidacionDatos.png"
wait 3
If Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Exist = True Then
		wait 3
		Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
End If



End Sub
Sub Validacion()
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_19").Click
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_21").Click
	wait 3
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "prueba"
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_22").Type micCtrlDwn + micAltDwn + "q" + micCtrlUp + micAltUp
	wait 3
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "gmail.com"
	wait 3

End Sub

Sub Prepago()
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Exist = false
		wait 1
	Wend
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Prepago.png", True
	imagenToWord "Click Prepago", RutaEvidencias() & "Prepago.png"
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Click
	wait 3
	If Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist=True Then
		wait 3
		Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Click
	End If
	

	wait 3
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN}"
	wait 3
	If  Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Exist=True Then
		wait 3
		If Window("Ejecutivo de interacción").InsightObject("InsightObject_23").Exist = true Then
			wait 3
			Window("Ejecutivo de interacción").InsightObject("InsightObject_23").Click
			wait 3
			Window("Ejecutivo de interacción").InsightObject("InsightObject_21").Click
			wait 3
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "prueba"
			wait 3
			Window("Ejecutivo de interacción").InsightObject("InsightObject_22").Type micCtrlDwn + micAltDwn + "q" + micCtrlUp + micAltUp
			wait 3
			Set shell = CreateObject("Wscript.Shell") 
			shell.SendKeys "gmail.com"
			wait 3
		End If
		wait 3
		 Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Click
		 
	End If

	wait 5
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ValidacionDatos.png", True
	imagenToWord "Validación de Datos Exitosa", RutaEvidencias() & "ValidacionDatos.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
End Sub
Sub Portabilidad()
	Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Exist = False
		wait 1
	Wend
	If  ucase(DataTable("Tipo_Venta", dtLocalSheet)) = "CHIP" Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Click
	ElseIf ucase(DataTable("Tipo_Venta", dtLocalSheet)) = "COMBO" Then
		Window("Ejecutivo de interacción").InsightObject("InsightObject_10").Click
	End If
	Window("Ejecutivo de interacción").InsightObject("InsightObject_11").Click
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys DataTable("Num_Porta", dtLocalSheet)
	Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Click
	wait 1
	If ucase(DataTable("Producto_Origen", dtLocalSheet)) = "POSTPAGO" Then
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN 2}"
		shell.SendKeys "{ENTER}"
	ElseIf ucase(DataTable("Producto_Origen", dtLocalSheet)) = "PREPAGO" Then
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN 1}"
		shell.SendKeys "{ENTER}"
	End If
	Window("Ejecutivo de interacción").InsightObject("InsightObject_14").Click
	If ucase(DataTable("Operador_Origen", dtLocalSheet)) = "OPERADOR VIRTUAL" Then
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN 8}"
		shell.SendKeys "{ENTER}"
	End If

	Window("Ejecutivo de interacción").InsightObject("InsightObject_15").Click


	
End Sub


	
