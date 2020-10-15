'Option Explicit

Dim tiempo, var1, var2, varfilas,varpag, var, vardisp, varsim, varlog, varfinan, varpagoinm, Num_Iter, flag, nroreg, shell, varhab, varasig, varasig2,mensaje, varValidaRespuestaCumplimiento, varsap, varselec, varfin, filas, Iterator
Dim str_titulo
Dim str_tipo_alta
Dim str_motivo_alta
Dim str_departamento
Dim str_provincia
Dim str_modeloCel
Dim str_plan_comp
Dim str_paquete
Dim str_tipoSIM
Dim str_mediopago
'Dim str_tipofinan
Dim str_idDispositivo
Dim str_idSim
Dim str_valEstadoOrden
Dim str_Wic
Dim str_RUC
Dim varRUC

'Se relacionan variables con DataTable
Num_Iter 		 	= 	Environment.Value("ActionIteration") 
str_RUC				=   DataTable("e_NumDocumento", "Buscar_Cliente")
str_tipo_alta		=	DataTable("e_Tipo_Alta", dtLocalSheet)
str_motivo_alta		=	DataTable("e_Motivo_Alta", dtLocalSheet)
str_departamento	=	DataTable("e_Departamento", dtLocalSheet)
str_provincia		=	DataTable("e_Provincia", dtLocalSheet) 
str_modeloCel		=	DataTable("e_ModeloCelular", dtLocalSheet)
str_plan_comp		=	DataTable("e_Plan", dtLocalSheet)
str_paquete			=	DataTable("e_paquete", dtLocalSheet)
str_tipoSIM			=	DataTable("e_TipoSIM", dtLocalSheet)
str_mediopago		=   DataTable("e_MedioPago", dtLocalSheet)
str_cant_cuota    	= 	DataTable("e_Cant_Cuota", dtLocalSheet)
str_tipofinan		=   DataTable("e_Tipo_Financiamiento", dtLocalSheet)
str_idDispositivo	=	DataTable("e_ID_Dispositivo", dtLocalSheet)
str_idSim			=	DataTable("e_ID_SIM", dtLocalSheet)
str_Wic				=	DataTable("e_WIC_ValidaCli", dtLocalSheet) @@ hightlight id_;_31334378_;_script infofile_;_ZIP::ssf8.xml_;_

'Métodos
Call SeleccionarUbicacion()
Call SeleccionarEquipo()
Call SeleccionarPlan()
Call ParametrosAlta()
Call RecursosAlta()
Call TipoEnvio()	
Call Financiamiento()
Call GeneracionOrden()
'If DataTable("e_Ambiente", "Login [Login]")<>"PROD" Then
	Call PagoManual()
'End If
Call GestionLogistica()
'If DataTable("e_Ambiente", "Login [Login]")<>"PROD" Then
	Call EmpujeOrden()
'End If
Call OrdenCerrado()
Call DetalleActividadOrden()

Sub SeleccionarUbicacion()
wait 10
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_3").JavaStaticText("Número de documento(st)").Exist)=False
			wait 1
		Wend
		
		Select Case str_tipo_alta
			Case "Alta Nueva Solo Linea"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#5"
			Case "Alta Nueva Equipo + Linea"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaTable("Titulo").ActivateRow "#4"
		End Select
		If str_Wic = "SI" Then
		

RunAction "WIC", oneIteration
		End If
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Departamento:").Exist) = False
			wait 1
		Wend
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Departamento:").WaitProperty "enabled", true, 10000 @@ hightlight id_;_14234044_;_script infofile_;_ZIP::ssf9.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Departamento:").Select DataTable("e_Departamento", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaList("Provincia:").Select DataTable("e_Provincia", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Ubicación"&".png", True
		imagenToWord "Ubicación ", RutaEvidencias() &Num_Iter&"_"&"Ubicación"&".png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Seleccionar Ubicación").JavaButton("Siguiente >").Click
		
		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaList("ComboBoxNative$1").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").Exist)) = False
			wait 1
		Wend
		
		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Buscar").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaButton("Buscar").Exist)) = False
			wait 1
		Wend
End Sub
Sub SeleccionarEquipo()
	If str_tipo_alta<>"Alta Nueva Solo Linea" Then
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaList("ComboBoxNative$1").Select "Celulares"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaEdit("TextFieldNative$1").Set DataTable("e_ModeloCelular", dtLocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Buscar").Click
		wait 7
		
		If DataTable("e_ModeloCelular", dtLocalSheet)<>"HUAWEI P10 NEGRO" Then
		
			tiempo=0
			Do
			tiempo=tiempo+1
				If tiempo>=120 Then
					DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
					DataTable("s_Detal0le", dtLocalSheet) = "El Equipo Móvil: "&DataTable("e_ModeloCelular", dtLocalSheet)&"no se encuentra"
					Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				Else
					Reporter.ReportEvent micPass, "Exito","Se encontro el equipo móvil buscado"
				End If
				wait 1
			Loop While Not (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Agregar al carrito").Exist(4) or JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(2) or JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(2))
			
		else
			tiempo=0
			Do
			tiempo=tiempo+1
				If tiempo>=120 Then
					DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
					DataTable("s_Detalle", dtLocalSheet) = "El Equipo Móvil: "&DataTable("e_ModeloCelular", dtLocalSheet)&"no se encuentra"
					Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				Else
					Reporter.ReportEvent micPass, "Exito","Se encontro el equipo móvil buscado"
				End If
				wait 1
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Agregar al carrito_2").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist))
			
		End If	
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(1) Then
			var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
			var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
			var1=replace(var1, "<html>", "")
			var1=replace(var1, "</html>", "")
			var1=replace(var1, "<br>", "")
			var1=replace(var1, "&#8203","")
			var1=replace(var1, "?;","")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=var1
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para LUDER").JavaButton("Cerrar").Click
			ExitActionIteration
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)= varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		   	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		   	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para LUDER").JavaButton("Cerrar").Click
		   	wait 2
		   	ExitActionIteration
		End If	
		
'	
'		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(1) Then
'			var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
'			var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
'			var1=replace(var1, "<html>", "")
'			var1=replace(var1, "</html>", "")
'			var1=replace(var1, "<br>", "")
'			var1=replace(var1, "&#8203","")
'			var1=replace(var1, "?;","")
'			DataTable("s_Resultado",dtLocalSheet)="Fallido"
'			DataTable("s_Detalle",dtLocalSheet)=var1
'			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para LUDER").JavaButton("Cerrar").Click
'			ExitActionIteration
'		End If
'		
'		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
'			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
'			DataTable("s_Resultado",dtLocalSheet)="Fallido"
'			DataTable("s_Detalle",dtLocalSheet)= varsap
'			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
'		   	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
'		   	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para LUDER").JavaButton("Cerrar").Click
'		   	wait 2
'		   	ExitActionIteration
'		End If
		
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EquipoMovil"&".png", True
		imagenToWord "Equipo Móvil", RutaEvidencias() &Num_Iter&"_"&"EquipoMovil"&".png"
		
		If DataTable("e_ModeloCelular", dtLocalSheet)<>"HUAWEI P10 NEGRO" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Agregar al carrito").Click
		else 
			'MsgBox "Selecciona Equipo Móvil"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para WILSON").JavaButton("Agregar al carrito_2").Click
		End If
		
			tiempo=0
			Do
			tiempo=tiempo+1
				If tiempo>=120 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Resultado", dtLocalSheet) = "No cargo la ventana 'Nuevo Plan (Para..)'"
			  		Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet), DataTable("s_Resultado", dtLocalSheet)
					ExitActionIteration
				else
					Reporter.ReportEvent micPass,"Exito","El combo para escoger los Planes Móviles cargo correctamente"
				End If
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))
			
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
			varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)= varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		   	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		   	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo equipo (Para LUDER").JavaButton("Cerrar").Click
		   	wait 2
		   	ExitActionIteration
		End If
	End If	
End Sub
Sub SeleccionarPlan()
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 7000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").Select DataTable("e_Tipo_Subcategoria", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 6000 @@ hightlight id_;_6652116_;_script infofile_;_ZIP::ssf13.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").Select DataTable("e_Tipos_Categoria_Plan", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaList("ComboBoxNative$1").WaitProperty "enabled", true, 10000 @@ hightlight id_;_6652116_;_script infofile_;_ZIP::ssf13.xml_;_
	
	If str_tipo_alta<>"Alta Nueva Solo Linea" Then
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaEdit("Equipo seleccionado:").Set DataTable("e_Plan", dtLocalSheet)
		else
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaEdit("TextFieldNative$1").Set DataTable("e_Plan", dtLocalSheet)
		wait 2
	End If 
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaButton("Buscar").Click
	wait 2
	
	tiempo = 0
		Do
		tiempo = tiempo + 1
			If tiempo>=120 Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se encuentra el plan: "&DataTable("e_TipoDePlan", dtLocalSheet)
				Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaCheckBox("Seleccionar").Exist(5) @@ hightlight id_;_7128919_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 2
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"PlanTarifario"&".png", True
	imagenToWord "Plan Tarifario", RutaEvidencias() &Num_Iter&"_"&"PlanTarifario"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaCheckBox("Seleccionar").Set "ON" @@ hightlight id_;_22129898_;_script infofile_;_ZIP::ssf2.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Nuevo Plan (Para WILSON").JavaButton("Siguiente >").Click
	
	tiempo = 0
		Do
		tiempo = tiempo + 1
			If tiempo>=120 Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Actualizar Atributos'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			else
			Reporter.ReportEvent micPass,"OK","Continuar Flujo"
			End If
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist(2)
End Sub
Sub ParametrosAlta()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select str_tipo_alta
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_motivo_alta
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Código de Centro Poblado").Set "0101010001"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaCheckBox("Tiene cobertura").Set "ON" @@ hightlight id_;_13164893_;_script infofile_;_ZIP::ssf3.xml_;_
	wait 2
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ParametrosAlta"&".png", True
	imagenToWord "Parametros de la Alta", RutaEvidencias() &Num_Iter&"_"&"ParametrosAlta"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click
	
		tiempo = 0
		Do
		tiempo = tiempo + 1
			If tiempo>=120 Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Negociar Configuración'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
			End If
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist))
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set @@ hightlight id_;_2822126_;_script infofile_;_ZIP::ssf4.xml_;_
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Siguiente >").Click
	End If
End Sub
Sub RecursosAlta()
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Calcular").Exist) =False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Asignación de número"
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Exist)=False
		wait 1
	Wend
	wait 1
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Set "920%%%%%%"
'	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Set DataTable("e_ID_Servicio",dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Proponer números").Click
	wait 2
	
		tiempo=0
			Do
				tiempo=tiempo+1
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Exist Then
					varasig=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").GetROProperty("enabled")
				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Exist Then
					varasig2=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").GetROProperty("enabled")
				End If
				wait 1
		Loop  While Not ((varasig="1") Or (varasig2="1") Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist))

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
		DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
		DataTable("s_Detalle", dtLocalSheet) = "No hay ningún número devuelto"
	    Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"No hay números disponibles"&".png", True
		imagenToWord "No hay números disponibles", RutaEvidencias() &Num_Iter&"_"&"No hay números disponibles"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaList("Motivo:").Select "Pedido de Cliente"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaButton("Aceptar").Click
		
		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar Cancelar").Exist)) = False
			wait 1
		Wend
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar Cancelar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar Cancelar").Click
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_3").JavaButton("Finalizar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción_3").JavaButton("Finalizar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaMenu("Archivo").JavaMenu("Salida").Select
			wait 2
		End If
		ExitTest
	End If
	wait 1
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Click
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Click
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("SearchJTable").Output CheckPoint("SearchJTable") @@ hightlight id_;_23150386_;_script infofile_;_ZIP::ssf16.xml_;_
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NumeroTelefonico"&".png", True
	imagenToWord "Número Telefónico", RutaEvidencias() &Num_Iter&"_"&"NumeroTelefonico"&".png"	

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Configuración" @@ hightlight id_;_26466124_;_script infofile_;_ZIP::ssf6.xml_;_
	wait 2
	If str_paquete= "SI" Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "_Before_Sub_Productos_Disponible.png", True
		imagenToWord "-SubProducto Disponibles-",RutaEvidencias() & "_Before_Sub_Productos_Disponible.png"
		wait 2	
		Dim objS
		set objS = CreateObject("WScript.Shell")
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand "#0;Móvil;Promociones y Descuentos Especiales;Bono 2GB x 5 dias (VR S/ 10.00)"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Select "#0;Móvil;Promociones y Descuentos Especiales;Bono 2GB x 5 dias (VR S/ 10.00)"
		wait 2
		objS.SendKeys "{ENTER}"
		wait 3
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "_SubProductosDisponibles_.png", True
		imagenToWord "-Eligir Promociones y Descuentos Especiales-",RutaEvidencias() & "_SubProductosDisponibles_.png"
	End If

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set "Tipo de SIM"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Click
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 8
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Exist(7) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 1
	End If
	varfilas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
			
	For Iterator = varfilas-1 To 0 Step -1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SelectRow "#"&Iterator
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").PressKey "C",micCtrl
			JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
			str_titulo=JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("text")
			str_titulo = Replace(str_titulo,"Nombre    Valor    Por única vez    Mensual     ","")
			If str_titulo="Grupo de SIM    NA            " Then
				str_titulo = Left(str_titulo,12)
				else
				str_titulo = Left(str_titulo,11)
			End If
			wait 1
					Select Case DataTable("e_Tipo_Alta", dtLocalSheet)
						Case "Alta Nueva Equipo + Linea"
							If str_titulo="Tipo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_TipoSIM", dtLocalSheet) 
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Mostrar_Atributos_"&Num_Iter&".png", True
								wait 2
								Exit For
							End  If
							
						Case "Alta Nueva Solo Linea"
							If str_titulo="Tipo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_TipoSIM", dtLocalSheet) 
								wait 2
							End  If
							If str_titulo="Grupo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", "Estandar"
								wait 2
							End  If
							If str_titulo="Número IMEI" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", "811111111111111"
								wait 2
								Exit For
							End If
						End Select	
					wait 1
		Next

		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
		wait 8
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTree("Subproductos disponibles").Expand "#0;Móvil;Promociones y Descuentos Especiales"
		wait 5 
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Promoe.png", True
			imagenToWord "No hay bonificación Gratitud", RutaEvidencias() &"Promoe.png"
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se encuentra el Bono gratitud"
			Reporter.ReportEvent micFail , DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
        Dim bono
		bono =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Agregar").GetROProperty("enabled")

		If bono <>"1" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cancelar oferta").Click
			wait 1
			While JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").JavaStaticText("Se cancelará esta orden").Exist=False
				wait 1	
			Wend
			wait 1	
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar una opción").JavaButton("Sí").Click
			wait 1
			
			While JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaTable("Acciones de orden que").Exist=False
				wait 2
			Wend
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaButton("Aceptar").Click
			wait 1
			While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaEdit("TextAreaNative$1").Exist=False
				wait 1
			Wend
		
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &"Promoer.png", True
			imagenToWord "Orden cancelada", RutaEvidencias() &"Promoer.png"
			ExitActionIteration
			
		End If
		wait 8

		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Tipo_SimCard"&".png", True
		imagenToWord "Tipo de SimCard ", RutaEvidencias() &Num_Iter&"_"&"Tipo_SimCard"&".png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
	
			tiempo = 0
			Do
			tiempo = tiempo + 1
				If tiempo>=120 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Negociar dirección'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				else
					Reporter.ReportEvent micPass,"OK","Continuar Flujo"
			End If
			wait 1
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist))
			
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(1) Then
				var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
				var1=replace(var1, "<html>", "")
				var1=replace(var1, "</html>", "")
				var1=replace(var1, "<br>", "")
				var1=replace(var1, "&#8203","")
				var1=replace(var1, "?;","")
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)=var1
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaList("Motivo:").Select "Pedido de Cliente"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaButton("Aceptar").Click
				wait 2
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar Cancelar").Exist(1) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar Cancelar").Click
					wait 3
				End If
				While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Exist)) = False
					wait 1
				Wend
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
					wait 2
				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Cerrar").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Cerrar").Click
					wait 2
				End If
				ExitActionIteration
		End If

		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
			var2=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
			var2=replace(var1,"<html>", "")
			var2=replace(var1,"</html>","")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = var1
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			var1 = Mid(var1,1,47 )
			
				If  var1="La regla rule with the following details failed" Then
					JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
					wait 2
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
				End If
				
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Exist(1) Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
				wait 2
			End If
			
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Exist(1) Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaList("Motivo:").Select "Pedido de Cliente"
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Configuración").JavaButton("Aceptar").Click
				wait 2
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar Cancelar").Exist(1) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar Cancelar").Click
					wait 3
				End If
				While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").Exist)) = False
					wait 1
				Wend
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
					wait 2
				End If
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Cerrar").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Cerrar").Click
					wait 2
				End If
				ExitActionIteration
			End If
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		End If
End Sub
Sub TipoEnvio()

	Select Case DataTable("e_MetodoEntrega", dtLocalsheet)
	
		Case "Delivery"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Delivery").Set "ON"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png" , True
			imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
			wait 1
				While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Exist) = False
						wait 1
				Wend
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Click
			wait 1
				While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Exist) = False
					wait 1
				Wend
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaEdit("TextFieldNative$1_2").Set "CAMINO REAL"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Buscar ahora").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
			wait 1
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png" , True
			imagenToWord "Dirección de Envio", RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
			wait 1
				While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist))= False
					wait 1
				Wend
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
				varDel=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)=varDel
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png", True
				imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Cerrar").Click
					While(JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Exist)=False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
					While(JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar dirección de").JavaButton("Aceptar").Exist)=False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar dirección de").JavaList("Motivo:").Select "Pedido de Cliente"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CancelarOrden"&".png", True
				imagenToWord "Cancelar Orden", RutaEvidencias() &Num_Iter&"_"&"CancelarOrden"&".png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar dirección de").JavaButton("Aceptar").Click
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Exist)=False
						wait 1
					Wend
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenCancelada"&".png", True
				imagenToWord "Orden Cancelada", RutaEvidencias() &Num_Iter&"_"&"OrdenCancelada"&".png"
				varOrden=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
				varOrden=Replace(varOrden,"Orden ","")
				varOrden="La orden: "&varOrden&" se cancelo debido a la validación: "&varDel
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
				ExitActionIteration
			End If	
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Buscar detalles de contacto").Click
			wait 1
				While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Buscar ahora").Exist)=False
					wait 1
				Wend
			wait 1	
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaList("ComboBoxNative$1").Select "DNI"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaEdit("TextFieldNative$1").Set "77286567"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Buscar ahora").Click
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
			wait 1
				While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
					wait 1
				Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "QA"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Número de teléfono del").Set "999999999"
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png", True
			imagenToWord "Entrega Delivery", RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click
		Case "En Tienda"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"TipoEnvio"&".png", True
			imagenToWord "Tipo de Envio ", RutaEvidencias() &Num_Iter&"_"&"TipoEnvio"&".png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click
	End Select
	
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("ID del Acuerdo de Facturación:").WaitProperty "editable", 1, 10000

			Do While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated_2").Exist) = False
				wait 1
					c=c+1 
					If (c=30) Then exit Do 
			Loop
		wait 1
	

		tiempo=0
		Do
		tiempo=tiempo+1
			If tiempo>=130 Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se habilito el 'Nombre y Dirección de Facturación'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			else
				Reporter.ReportEvent micPass, "Exito","'Nombre y Dirección de Facturación' se habilito correctamente"
			End If
		wait 2
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaEdit("Nombre y Dirección de").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist))

		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(1) Then
			var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
			var1=replace(var1, "<html>", "")
			var1=replace(var1, "</html>", "")
			var1=replace(var1, "<br>", "")
			var1=replace(var1, "&#8203","")
			var1=replace(var1, "?;","")
			MsgBox var1
			wait 1
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=var1
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
			ExitActionIteration
		End If
	

		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
			var1=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se puede seleccionar método de entrega"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar dirección de").JavaList("Motivo:").Select "Pedido de Cliente"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar dirección de").JavaButton("Aceptar").Click
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar Cancelar").Click
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Cerrar").Click
			wait 2
			ExitActionIteration
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").Exist(1) Then
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").WaitProperty "enabled", True, 8000
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaRadioButton("Nuevo").Set @@ hightlight id_;_16599652_;_script infofile_;_ZIP::ssf8.xml_;_
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaList("Mostrar:").Select  "Acciones de orden activas "

			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Siguiente >").Click
			wait 2
		End If

			tiempo = 0
			Do
			tiempo = tiempo + 1
				If tiempo>=120 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla 'Negociar Pago'"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					ExitActionIteration
				else
					Reporter.ReportEvent micPass,"OK","Continuar Flujo"
				End If
			wait 1
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist))
			
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
			mensaje=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=mensaje
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
			wait 2
			ExitActionIteration
		End If
			
End Sub
Sub Financiamiento()

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Exist Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Límite de Compra").Exist Then
			varfinan=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Límite de Compra").GetROProperty("enabled")
			wait 1
			If varfinan="1" Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Límite de Compra").Click
			End If
		End If
		
			tiempo=0
			Do 
				tiempo=tiempo+1
				varpagoinm=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").GetROProperty("enabled")
				wait 2
			Loop While Not (varpagoinm="1")
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Pago inmediato").Click
			wait 2
				While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
					wait 1
				Wend
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
		DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)=varsap
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacción"&".png", True
		imagenToWord "Mensaje de Validacción", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacción"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NegociarPago"&".png", True
		imagenToWord "Negociar Pago", RutaEvidencias() &Num_Iter&"_"&"NegociarPago"&".png"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Siguiente >").Click
		
			While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist)Or(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist))=False
				wait 1
			Wend
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Cerrar").Click
'			wait 2
'			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
'			wait 2
'			ExitActionIteration
	End If
	wait 5
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Exist Then
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").Exist(2) Then
			varRUC = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").GetROProperty("text")
			If (varRUC = "") Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").SetFocus
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Numero RUC").Set str_RUC
			End If		
			Wait 2
		End If
		
		
		
		wait 2
		Dim Iterator, Count
		Dim rs
	Count = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").GetROProperty ("items count")
	'MsgBox 	Count
	For Iterator = 0 To Count-1
	 	rs = 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").GetItem (Iterator)
	 	'MsgBox rs
		If rs = DataTable("e_MedioPago", dtLocalSheet) Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select DataTable("e_MedioPago", dtLocalSheet)
			    wait 1
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Cantidad de cuotas:").Select DataTable("e_Cant_Cuota" , dtLocalSheet)
					wait 1
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png" , True
					imagenToWord "Financiamiento", RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png"
			
			Exit for
		ElseIf Iterator = Count-1 Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select "Externo"
			Exit for
		End if	
	Next
	wait 1
'		If (str_mediopago= "Pago a la Factura") Then
'			wait 2
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select "Pago a la Factura"
'			wait 1
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Cantidad de cuotas:").Select str_cant_cuota
'			wait 1
''			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Calcular").Click
''			wait 3
'		End If
'		
'		
'		
		
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
		wait 5
			While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Siguiente >").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist)) = False
				wait 1
			Wend
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
			varfin=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
			varfin=replace(varfin, "<html>", "")
			varfin=replace(varfin, "</html>", "")
			varfin=replace(varfin, "<br>", "")
			varfin=replace(varfin, "&#8203","")
			varfin=replace(varfin, "?;","")
			varfin=replace(varfin, "&nbsp;","")
			DataTable("s_Resultado",dtLocalSheet)="Fallido"
			DataTable("s_Detalle",dtLocalSheet)=varfin
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Cancelar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
			wait 5
			ExitActionIteration
		End  If
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Siguiente >").Click
			While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist)) =False
				wait 1
			Wend
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
		varpag=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varpag
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 4
		ExitActionIteration
	End If
		
End Sub
Sub GeneracionOrden()
	
	tiempo = 0
	Do
		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
			If DataTable("e_WIC_ContrCli",dtLocalSheet)="SI" Then
					
RunAction "WIC2", oneIteration
			
				Exit Do
			End If
			wait 2
		Wend
	
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist Then
			wait 3
			var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
	   		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
	   		wait 2	
		End If
			
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
			wait 2
		End If
		wait 1
			
			If tiempo>=180 Then
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = "No se a cargado el contrato correctamente'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				ExitActionIteration
			else
				Reporter.ReportEvent micPass,"Contrato Exitoso","Se a cargado el contrato correctamente"
			End If
	wait 2
	Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist) or (var1 = "Contratos no Generados") or (var1="0"))
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist(1) Then
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
		wait 1
	End If

	'Bucle que espera "Enviar orden"
	t = 0
	While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist) = False
		Wait 1
		t = t + 1
		If (t >= 180) Then
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = "No se habilitó el botón -Enviar orden-"
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ErrorbtnEnviarOrden_"&Num_Iter&".png", True
			imagenToWord "No se habilitó el botón -Enviar orden_"&Num_Iter, RutaEvidencias() & "ErrorbtnEnviarOrden_"&Num_Iter&".png"
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			ExitActionIteration
		End If
	Wend
	Wait 1

	'Click en "Enviar orden"
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist(2) Then
		'Damos clic en el boton "Enviar Orden"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		Wait 3
	End If
	
		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
			wait 1
		Wend	
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		wait 1
		varVen=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validar").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ValidacionVendedor"&".png", True
		imagenToWord "Validación Vendedor", RutaEvidencias() &Num_Iter&"_"&"ValidacionVendedor"&".png"
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaButton("Seleccionar").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		wait 1	
	End If
	
	DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
	flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
	MsgBox flag
	DataTable("s_Nro_Orden", dtLocalSheet) = replace (DataTable("s_Nro_Orden", dtLocalSheet),"Orden ","")
	Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
	wait 1

'	Select Case DataTable("e_Ambiente", "Login [Login]")
'		Case "UAT8"
'				DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
'				flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'				DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),8)
'				Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada"&".png", True
'				imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Ubicación"&".png"
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
'				wait 2
'		Case "UAT4"
'				DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
'				flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'				DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'				Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada"&".png", True
'				imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Ubicación"&".png"
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
'				wait 2
'		Case "UAT6"	
'				DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
'				flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'				DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'				Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada"&".png", True
'				imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Ubicación"&".png"
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
'				wait 2
'		Case "UAT10"	
'				DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
'				flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'				DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),6)
'				Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada"&".png", True
'				imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Ubicación"&".png"
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
'				wait 2
'		Case "UAT13"	
'				DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
'				flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'				DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),6)
'				Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada"&".png", True
'				imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Ubicación"&".png"
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
'				wait 2
'		Case "PROD"	
'				DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").GetROProperty("text")
'				flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'				DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),10)
'				Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada"&".png", True
'				imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Ubicación"&".png"
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 2302300A").JavaButton("Cerrar").Click
'				wait 2
'	End Select
	
	If DataTable("e_MetodoEntrega", dtLocalSheet)="Delivery" Then
		ExitActionIteration
	End If
	
End Sub
Sub PagoManual()
		
	If (str_mediopago<>"Pago a la Factura") Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_24061018_;_script infofile_;_ZIP::ssf2.xml_;_
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist) = False
			wait 1
		Wend
		wait 1
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
			wait 2
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
		wait 2
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
		wait 2		
		
			tiempo=0
			Do	
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("attached text")
					tiempo=tiempo+1
					wait 1
				End If
				If (tiempo >= 180) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Pago de la Orden"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					ExitActionIteration
				End If
				Loop While Not (nroreg="1 Registros")
				
		wait 1	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Exist) = False
			wait 1
		Wend
	
			tiempo=0
			Do
				var = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").GetROProperty("enabled")
				tiempo=tiempo+1
					If (tiempo >= 8) Then
						DataTable("s_Resultado", dtLocalSheet) = "Exito"
			  			DataTable("s_Detalle", dtLocalSheet) = "La orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" ya fue pagado"
						Reporter.ReportEvent micPass, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Cancelar").Click
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
						wait 2
						Exit Do
					End If
 @@ hightlight id_;_23686897_;_script infofile_;_ZIP::ssf5.xml_;_
			Loop While Not (var <> "0")
			wait 1
	
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Exist Then
		var=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").GetROProperty("enabled")
			If var = 1 Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Enviar").Click
				Else 
				JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").JavaButton("Cancelar").Click
			End If
			wait 1
		End If
	End If
	
End Sub
Sub GestionLogistica()
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
	wait 1
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist) = False
		wait 1
	Wend
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click

		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("1 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 180) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar la Gestion Logistica"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					ExitActionIteration
			End If
		Loop While Not (nroreg="1 Registros")
		wait 1
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").SelectRow "#0" @@ hightlight id_;_20884557_;_script infofile_;_ZIP::ssf5.xml_;_
		wait 1
		
		tiempo=0
		Do
		
		If (DataTable("s_Detalle", dtLocalSheet)="Por favor rellenar todas las identificaciones de equipos") or (DataTable("s_Detalle", dtLocalSheet)<>"Por favor rellenar todas las identificaciones de equipos") Then
			
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Exist(1) Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Click
				wait 2
			End If
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Gestionar logística").Click
			tiempo=tiempo+1
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").Exist) = False
				wait 1
			Wend
			
			
				If str_tipo_alta="Alta Nueva Equipo + Linea" Then
			
					vardisp=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData(1,4)
					If vardisp<>str_idDispositivo Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
						Set shell = CreateObject("Wscript.Shell") 
						shell.SendKeys "{ENTER}"
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_Dispositivo", dtLocalSheet)
						wait 2
					End If
				
					varsim=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData(2,4)
					If varsim<>str_idSim Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
						Set shell = CreateObject("Wscript.Shell") 
						shell.SendKeys "{ENTER}"
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",DataTable("e_ID_SIM", dtLocalSheet)
						wait 2
					End If
					
				else
				
					varsim=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData(1,4)
					If varsim<>str_idSim Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
						Set shell = CreateObject("Wscript.Shell") 
						shell.SendKeys "{ENTER}"
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_SIM", dtLocalSheet)
						wait 2
					End If
				End If
			
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Materiales"&".png", True
			imagenToWord "Materiales", RutaEvidencias() &Num_Iter&"_"&"Materiales"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Validar y Crear Factura").Object.doClick()
			
			tiempo = 0
			Do
					varhab=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").GetROProperty("enabled")					
					wait 3
			Loop While Not((JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (varhab="1"))
			
				If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(2) or JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(2) Then
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(0) Then
						varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
					End If
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(0) Then
						varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
					End If
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		       		DataTable("s_Detalle", dtLocalSheet) = varlog
		       		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
		     		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
		     			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"validacionData"&".png", True
						imagenToWord "Validación Data", RutaEvidencias() &Num_Iter&"_"&"validacionData"&".png"
		        		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Click
		        	End If
		        	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
		        		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		        	End If
		        	wait 2
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Exist(1) Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Click
						wait 2
					End If
		     		If DataTable("s_Detalle", dtLocalSheet)<>"Por favor rellenar todas las identificaciones de equipos" Then
						If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist Then
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
						End If
						If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist(2) Then
							JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
							ExitActionIteration
						End If	
		     		End  If
		    	End If
		End  If
		
		If tiempo>=16 Then
			Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
			DataTable("s_Resultado",dtLocalSheet) = "Fallido"
			DataTable("s_Detalle",dtLocalSheet) = "Luego de 3 intentos no se pudo realizar la Asignación de Series"
			ExitActionIteration
		else
			Reporter.ReportEvent micPass, "Exito", "Se realizo la Asignación de Series correctamente"
		End If
		Loop While Not varhab = "1"
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Click
	End If
End Sub
Sub EmpujeOrden()
		
	If DataTable("e_Tipo_De_DATA_Sim", dtLocalSheet) = "DATA LOGICA" Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_19748072_;_script infofile_;_ZIP::ssf61.xml_;_
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist) = False
			wait 1
		Wend
		
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
			wait 3
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
		wait 2
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo" @@ hightlight id_;_25130440_;_script infofile_;_ZIP::ssf62.xml_;_
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
		wait 2
		
			tiempo=0
			Do 
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("attached text")
					tiempo=tiempo+1
					wait 1
				End If
				
				If (tiempo >= 80) Then
						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
						DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"
						Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
						ExitActionIteration
				End If
			Loop While Not(nroreg="1 Registros")
			wait 1
			
'			tiempo=0
'			Do 
'				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
'					wait 2
'					tiempo=tiempo+1
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").Output CheckPoint("Equipo usuario:")
'					varValidaRespuestaCumplimiento=Environment("s_ValidaManejarRespuestaCumplimiento")
'					wait 1
'				End If
'				
'				If (tiempo >= 80) Then
'						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'						DataTable("s_Detalle", dtLocalSheet) = "La actividad 'Manejar Respuesta de Cumplimiento' no cargo"
'						Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
'						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
'						ExitActionIteration
'				End If
'			Loop While Not varValidaRespuestaCumplimiento="Manejar Respuesta de Cumplimiento" @@ hightlight id_;_2317921_;_script infofile_;_ZIP::ssf3.xml_;_
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Exist) = False
			wait 1
		Wend
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo Parcial"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		wait 2
	End If
	
End Sub
Sub OrdenCerrado()
	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select @@ hightlight id_;_17809817_;_script infofile_;_ZIP::ssf6.xml_;_
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist) = False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF"
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
	wait 5
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
	Reporter.ReportEvent micPass,"Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
	
	tiempo = 0
		Do
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
		tiempo = tiempo +1
		wait 3
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:")
			If (tiempo >= 10) Then
				DataTable("s_Resultado",dtLocalSheet) = "Fallido" 
				DataTable("s_Detalle", dtLocalSheet) = "La Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" no culmino en estado Cerrado"
				Reporter.ReportEvent micFail,DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
				If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
					Exit Do
					wait 1
				End If			
				'ExitActionIteration
			else
				Reporter.ReportEvent micPass, "Se valida el estado de la orden", DataTable("s_ValEstadoOrden", dtLocalSheet)
			End If
		wait 1
		Loop While Not DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado"
		DataTable("s_Resultado", dtLocalSheet) = "Exito"
		DataTable("s_Detalle", dtLocalSheet) = "La orden culmino correctamente"
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		DataTable("s_Resultado", dtLocalSheet) = "Exito"
		DataTable("s_Detalle", dtLocalSheet) = "La orden culmino correctamente"
		Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"OrdenCerrada"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"OrdenCerrada"&".png"
	
End Sub
Sub DetalleActividadOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaList("Ver por:").Select "Acciones de orden"
	wait 3
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").DoubleClickCell 0, "#8", "LEFT"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
	wait 1
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2957187A").JavaEdit("Fecha de vencimiento:").Exist)=False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2957187A").JavaTab("Nombre del cliente:").Select "Actividad"
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2957187A").JavaTable("SearchJTable").Exist)=False
		wait 1	
	Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png"
	
	shell.SendKeys "{PGDN}"
	wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png"
	
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2957187A").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = 0 To filas-1
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2957187A").JavaTable("SearchJTable").GetCellData(Iterator,0)
	Next
	
	If varselec<>"Cerrar Acción de Orden" Then
	 	DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en la Actividad "&varselec&""
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2957187A").JavaButton("Cancelar").Click
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
		ExitActionIteration
		wait 1
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2957187A").JavaButton("Cancelar").Click
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	
End Sub



		

