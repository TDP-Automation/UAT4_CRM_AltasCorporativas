



While Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = false
	wait 1
Wend
If Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Exist= True Then
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Incidencia.png", True
	imagenToWord "Aviso: Actualmente mos encontramos en incidencia", RutaEvidencias() & "Incidencia.png"
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Click
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Click
Else
	Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContratacionValidacion.png", True
	imagenToWord "Contratación y Validación de Identidad", RutaEvidencias() & "ContratacionValidacion.png"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN 2}"
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Exist = false
		wait 1
	Wend
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Click
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Exist = False
		wait 1
	Wend
	wait 3
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DescargarContrato.png", True
	imagenToWord "Descargar Contrato", RutaEvidencias() & "DescargarContrato.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Click
	wait 3
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_5").Exist = False
		wait 1
	Wend
	wait 3
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "DescargaCompleta.png", True
	imagenToWord "Descarga Completa", RutaEvidencias() & "DescargaCompleta.png"
	Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Click
End If



