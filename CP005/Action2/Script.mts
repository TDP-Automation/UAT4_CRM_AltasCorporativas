Dim ValidaContr, shell, TipoCliente
Dim Opcion1,Opcion2,Opcion3,Opcion4,Opcion5
Dim Padre, Madre, Nacimiento


ValidaContr = 0
TipoCliente = Datatable("e_TipoDocCliente", dtLocalSheet)


Call IniciaProceso()
If ucase(TipoCliente) = "CE" Then
	Call ValidaCE()
ElseIf ucase(TipoCliente) = "DNI" Then
	Call ValidaDNI()
End If


Sub IniciaProceso()
	Do 
		
		If Window("Ejecutivo de interacción").InsightObject("InsightObject").Exist = True Then
			Call Captura("Actualiza para ver los contratos","ActContratos")
			Window("Ejecutivo de interacción").InsightObject("InsightObject").Click
			Wait 5
		End If
		If Window("Ejecutivo de interacción").InsightObject("InsightObject_2").Exist = True Then
			ValidaContr = 1
		End If
		wait 3
		If Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Exist = True Then
			wait 2
			ValidaContr = 1
		End If
	Loop While Not ValidaContr = 1
	
End Sub
Sub ValidaCE()

	Call Captura("Se realizara la Validacion con Ciudadano Extranjeria","ValidaCE")
	Window("Ejecutivo de interacción").InsightObject("InsightObject_3").Click
	Call PageDown()
	Call Contratos()
	Call PageDown()
	while Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Exist = false
		wait 1
	wend
	Call Captura("Descarga de documentos completa","ContratosCompl")
	Call PageDown()
	Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Click

End Sub
Sub ValidaDNI()
	wait 3
	Window("Ejecutivo de interacción").InsightObject("InsightObject_10").Click
	wait 2
	Call Captura("Se realizara la Validacion con DNI - Discapacitado Huella Desgastada","ValidaDNI")
	Call PageDown()
	wait 2
	Window("Ejecutivo de interacción").InsightObject("InsightObject_11").Click
	
	
	Call ValidacionCuestionario()
	Call PageDown()
	Call DeclaracionJurada()
	Call Contratos()
	Call PageDown()
	Call ValidacionCompleta()
End Sub
Sub ValidacionCuestionario()

	Padre = Datatable("e_NombrePadre", dtLocalSheet)
	Madre = Datatable("e_NombreMadre", dtLocalSheet)
	Nacimiento = DataTable("e_Nacimiento", dtLocalSheet)
	
	
	
	For Iterator = 1 To 3 Step 1

		While Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Exist = False
			wait 1
		Wend
		
		
		Opcion1 = Window("Ejecutivo de interacción").InsightObject("InsightObject_12").GetVisibleText(50, 6, 200, 32)
		Opcion1 = replace(Opcion1,"5","S")
		Opcion2 = Window("Ejecutivo de interacción").InsightObject("InsightObject_13").GetVisibleText(58, 7, 260, 33)
		Opcion2 = replace(Opcion2,"5","S")
		Opcion3 = Window("Ejecutivo de interacción").InsightObject("InsightObject_14").GetVisibleText(62, 4, 320, 31)
		Opcion3 = replace(Opcion3,"5","S")
		Opcion4 = Window("Ejecutivo de interacción").InsightObject("InsightObject_15").GetVisibleText(45, 6, 324, 31)
		Opcion4 = replace(Opcion4,"5","S")
		Opcion5 = Window("Ejecutivo de interacción").InsightObject("InsightObject_16").GetVisibleText(53, 8, 361, 36)
		Opcion5 = replace(Opcion5,"5","S")
		
		
		Select Case ucase(Padre)
			Case UCASE(Opcion1)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Click
				Opcion1 =""
			Case UCASE(Opcion2)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_13").Click
				Opcion2 =""
			Case UCASE(Opcion3)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_14").Click
				Opcion3 =""
			Case UCASE(Opcion4)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_15").Click
				Opcion4 =""
			Case UCASE(Opcion5)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_16").Click
				Opcion5 =""
		End Select
		Select Case ucase(Madre)
			Case UCASE(Opcion1)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Click
				Opcion1 =""
			Case UCASE(Opcion2) 
				Window("Ejecutivo de interacción").InsightObject("InsightObject_13").Click
				Opcion2 =""
			Case UCASE(Opcion3)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_14").Click
				Opcion3 =""
			Case UCASE(Opcion4)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_15").Click
				Opcion4 =""
			Case UCASE(Opcion5)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_16").Click
				Opcion5 =""
		End Select
		Select Case ucase(Nacimiento)
			Case UCASE(Opcion1)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_12").Click
				Opcion1 =""
			Case UCASE(Opcion2)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_13").Click
				Opcion2 =""
			Case UCASE(Opcion3)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_14").Click
				Opcion3 =""
			Case UCASE(Opcion4)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_15").Click
				Opcion4 =""
			Case UCASE(Opcion5)
				Window("Ejecutivo de interacción").InsightObject("InsightObject_16").Click
				Opcion5 =""
		End Select
		Call Captura("Seleccion "&Iterator,"Sel"&Iterator)
		Window("Ejecutivo de interacción").InsightObject("InsightObject_17").Click
	Next
	
	
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_24").Exist = False
		wait 1
	Wend
	Call Captura("Cuestionario Aprobado","Aprobado")
	
	Window("Ejecutivo de interacción").InsightObject("InsightObject_25").Click
	

End Sub
Sub DeclaracionJurada()
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_26").Exist = false
		wait 1
	Wend
	Window("Ejecutivo de interacción").InsightObject("InsightObject_26").Click
	While Window("Ejecutivo de interacción").InsightObject("InsightObject_27").Exist = False
		wait 1
	Wend
	
	Call Captura("Se visualiza Declaracion Jurada","DecJurada")
	
	Window("Ejecutivo de interacción").InsightObject("InsightObject_28").Click

	
End Sub
Sub Contratos()
	wait 2
	Call PageDown()
	Window("Ejecutivo de interacción").InsightObject("InsightObject_4").Click

	While Window("Ejecutivo de interacción").InsightObject("InsightObject_6").Exist = False
		wait 1
	Wend
	Call Captura("Se visualizan los contratos","Contratos")
	Window("Ejecutivo de interacción").InsightObject("InsightObject_7").Click
	
End Sub
Sub ValidacionCompleta()
		
	While ((Window("Ejecutivo de interacción").InsightObject("InsightObject_8").Exist) OR (Window("Ejecutivo de interacción").InsightObject("InsightObject_29").Exist)) = False
		wait 1
	Wend
	Call Captura("Se realizo la validacion Correctamente","ValCorrecta")
	Window("Ejecutivo de interacción").InsightObject("InsightObject_9").Click

End Sub


'------- Sub's de ayuda para el desarrollo-----------------------
Sub Captura(Texto,Imagen)
	Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Imagen&".png", True
	imagenToWord Texto,RutaEvidencias() &Imagen&".png"
End Sub
Sub Down(cantidad)
	For Iterator = 1 To cantidad Step 1
		WAIT 1
		Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{DOWN}"
	Next
End Sub
Sub Tab()
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{TAB}"
End Sub
Sub PageDown()
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGDN}"
End Sub
Sub EsperaElemento(Elemento)
	t = 0
	While (Elemento.Exist) = False
		Wait 1
		
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtlocalSheet) = "Fallido"
			DataTable("s_Detalle", dtlocalSheet) = "No se realizo la Carga del componente Correctamente"
			Window("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error.png", True
			imagenToWord "No se realizo la Carga del componente Correctamente",RutaEvidencias() &Num_Iter&"_"&"Error.png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtlocalSheet), DataTable("s_Detalle", dtlocalSheet)
			Wait 2
		'
		End If
	Wend
End Sub
