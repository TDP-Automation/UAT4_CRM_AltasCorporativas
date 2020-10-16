
Option  Explicit

Dim Valor, Fila, Valor2, i, Filas, varsap, tiempo, var, var1, flag, vartext, vardisp, varsim, shell, varlog, varhab, Num_Iter, varfinan, varfila, varmsgval, varId
Dim varValidaRespuestaCumplimiento, nroreg, varpagoinm, varnuevo, varprob, varasig, varasig2, varacuer, vardepo, varlogis, Iterator, varselec, str_titulo, varmsg, varvend
Dim str_tipoalta
Dim str_motivoalta
Dim str_tipoSIM
Dim str_metodo_entrega
Dim str_tipofinan
Dim str_planfinan
Dim str_mediopago
Dim str_cant_cuota
Dim str_tipodata
Dim str_idDispositivo
Dim str_idSim
Dim c

Num_Iter 		 	= Environment.Value("ActionIteration") 
str_tipoalta     	= DataTable("e_Tipo_Alta", dtLocalSheet)
str_motivoalta   	= DataTable("e_Motivo_Alta", dtLocalSheet)
str_tipoSIM       	= DataTable("e_TipoSIM", dtLocalSheet)
str_metodo_entrega	= DataTable("e_Metodo_Entrega", dtLocalSheet)
str_tipofinan    	= DataTable("e_Tipo_Financiamiento", dtLocalSheet)
str_planfinan		= DataTable("e_PlanFinanciamiento",dtLocalSheet)
str_mediopago		= DataTable("e_MedioPago", dtLocalSheet)
str_cant_cuota    	= DataTable("e_Cant_Cuota", dtLocalSheet)
str_tipodata      	= DataTable("e_Tipo_De_DATA_Sim", dtLocalSheet)
str_idDispositivo 	= DataTable("e_ID_Dispositivo", dtLocalSheet)
str_idSim         	= DataTable("e_ID_SIM", dtLocalSheet)
Num_Iter		  	= Environment.Value("ActionIteration")

Call EncontrarAcuerdoComercial()
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

Sub EncontrarAcuerdoComercial()
	wait 10
		Do
		tiempo=0
		Do
			tiempo=tiempo+1
			varacuer = JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Visión General del Cliente").GetROProperty("enabled")
			wait 1
		Loop While Not (varacuer="1")
	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
	JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Visión General del Cliente").Select
	wait 1
	
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("SearchJTable").Exist) = False
			wait 1
		Wend
	
	Valor = ""
		filas = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("SearchJTable").GetROProperty("rows")
		
		dim Iterator , filas	
		
		For Iterator = 0 To filas-1 step 1 
			
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("SearchJTable").ClickCell Iterator , 4
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("SearchJTable").DoubleClickCell "#"&Iterator ,"#4", "LEFT"
				
				
'				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("SearchJTable").PressKey "C",micCtrl
				JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").PressKey "V",micCtrl
				'MsgBox Valor
				Valor = JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").GetROProperty("value")
				'MsgBox Valor
				Valor = Left(Valor,21)
				Valor2 = DataTable("e_Grupo_Negocio","A_Comercial")
		If Valor = Valor2 Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaTable("SearchJTable").ActivateRow "#"&Iterator
			'MsgBox Valor2
			Reporter.ReportNote "Se encontro el valor"
        	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
   			Exit For
    	End If
		If Iterator = filas -1 Then
			Reporter.ReportNote "No se encontró el valor, se repetira la acción"
        	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
        	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaButton("Cerrar").Click
    	Exit For
		End If
    	JavaWindow("Ejecutivo de interacción").JavaEdit("Titulo").Set ""
		Next
	Loop While Not Valor = Valor2
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"AcuerdoComercialSeleccionado"&".png" , True
	imagenToWord "Acuerdo Comercial Seleccionado", RutaEvidencias() &Num_Iter&"_"&"AcuerdoComercialSeleccionado"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").JavaButton("Alta Individual").Click
	wait 2
End Sub	
Sub ParametrosAlta()
	
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").JavaButton("<html>Dar de alta</html>").Exist )= False
			wait 1
		Wend
	If str_tipoalta<>"Alta Nueva solo SIM" Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").JavaTable("Mensual").SelectRow "#0" @@ hightlight id_;_30573475_;_script infofile_;_ZIP::ssf59.xml_;_
		'MsgBox "Selecciona Equio Móvil"
		wait 1
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Proporciopnar_Suscriptor_"&".png" , True
	imagenToWord "Proporcionar Suscriptor", RutaEvidencias() &Num_Iter&"_"&"Proporciopnar_Suscriptor_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").JavaButton("<html>Dar de alta</html>").Click @@ hightlight id_;_11903749_;_script infofile_;_ZIP::ssf60.xml_;_
	'MsgBox "Selecciona Dar de alta"
	
		tiempo=0
		Do
		tiempo=tiempo+1
			If tiempo>=180 Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="No cargo la ventana Actualizar Atributo no cargo"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaActualizarAtributos"&".png" , True
				imagenToWord "Error Carga Actualizar Atributos", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaActualizarAtributos"&".png"
				ExitActionIteration
			else
				Reporter.ReportEvent micPass, "Exito","La ventana 'Actualizar Atributo' correctamente"
			End If
			wait 1
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Exist) or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist))
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		varsap=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		  	DataTable("s_Detalle", dtLocalSheet) = varsap
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error_SAP_"&".png", True
			imagenToWord "Error SAP", RutaEvidencias() &Num_Iter&"_"&"Error_SAP_"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
			SystemUtil.CloseProcessByName "javaw.exe"
			SystemUtil.CloseProcessByName "jp2launcher.exe"
			ExitTest
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Problema"&".png", True
		imagenToWord "No Hay Números Disponible", RutaEvidencias() &Num_Iter&"_"&"Problema"&".png"
		var=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
		var= replace(var,"<html>","")
		var= replace(var,"</html>","")
		var= replace(var,"<br>","")
		var= replace(var,"?&#8203","")
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = var
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		wait 1
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Proporcionar suscriptor").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Close
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Click
			wait 2
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 1
			End If
			JavaWindow("Ejecutivo de interacción").JavaMenu("Archivo").JavaMenu("Salida").Select
		End If
		ExitTest
	End If
	
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").WaitProperty "enabled", true, 10000
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaList("Motivo:").Select str_tipoalta @@ hightlight id_;_440740_;_script infofile_;_ZIP::ssf65.xml_;_
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Texto del motivo:").Set str_motivoalta
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaEdit("Código de Centro Poblado").Set "0101010001"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaCheckBox("Tiene cobertura").Set "ON"
	wait 1
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Parametros_Alta_"&".png", True
	imagenToWord "Parametros de la Alta", RutaEvidencias() &Num_Iter&"_"&"Parametros_Alta_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Actualizar Atributos de").JavaButton("Siguiente >").Click
	
		tiempo=0
		Do
		tiempo=tiempo+1
			If tiempo>=180 Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="La ventana 'Parametrización del Producto' no cargo"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaParametrizaelProducto"&".png" , True
				imagenToWord "Error Carga Parametriza el Producto", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaParametrizaelProducto"&".png"
				ExitActionIteration
			else
				Reporter.ReportEvent micPass, "Exito", "La ventana 'Parametrización del Producto cargo correctamente'"
			End If
			wait 1
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist(1)
		
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaRadioButton("Contacto por Defecto").Set
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Parametriza el Producto").JavaButton("Siguiente >").Click
	End If
		
		tiempo=0
		Do
		tiempo=tiempo+1
			If tiempo>=180 Then
				DataTable("s_Resultado",dtLocalSheet)="Fallido"
				DataTable("s_Detalle",dtLocalSheet)="La ventana 'Negociar Configuración del Producto Móvil no cargo'"
				Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Proporciopnar_Suscriptor_"&".png" , True
				imagenToWord "Proporcionar Suscriptor", RutaEvidencias() &Num_Iter&"_"&"Proporciopnar_Suscriptor_"&".png"
				ExitActionIteration
			else
				Reporter.ReportEvent micPass, "Exito", "La ventana 'Negociar Configuración del Producto Móvil' cargo correctamente"
			End If
			wait 1
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Exist(1)
End Sub
Sub RecursosAlta()

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Asignación de número"
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend
'	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Set "6%%%%%%%%"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("TextFieldNative$1").Set "920%%%%%%"
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
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"NoHayNúmeroDisponible_"&".png", True
		imagenToWord "No Hay Números Disponible", RutaEvidencias() &Num_Iter&"_"&"NoHayNúmeroDisponible_"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
	  	DataTable("s_Detalle", dtLocalSheet) = "No hay ningún número devuelto desde Resource Manager System."
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Close
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaMenu("Archivo").JavaMenu("Salida").Select
		End If
		ExitTest
	End If
	wait 1
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Distribuir Número").Click
		wait 1
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Asignar número").Click
		wait 1
	End If
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Número_Telefónico_"&".png", True
	imagenToWord "Número Telefónico", RutaEvidencias() &Num_Iter&"_"&"Número_Telefónico_"&".png"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("SearchJTable").Output CheckPoint("SearchJTable") @@ hightlight id_;_2328002_;_script infofile_;_ZIP::ssf2.xml_;_
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTab("JXTabbedPane").Select "Configuración"			
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaEdit("Buscar por nombre:").Set "Tipo de SIM"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Buscar").Click
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 8
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 1
	End If
	
	varId = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetCellData (1,1) 
	varfila=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").GetROProperty("rows")
	For Iterator = varfila-1 To 0 Step -1
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
					Select Case DataTable("e_Tipo_Alta", dtLocalSheet)
						Case "Alta Nueva Equipo + SIM"
							If str_titulo="Tipo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_TipoSIM", dtLocalSheet) 
								JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Mostrar_Atributos_"&Num_Iter&".png", True
								wait 1
								Exit For
							End  If
							
						Case "Alta Nueva solo SIM"
							If str_titulo="Tipo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", DataTable("e_TipoSIM", dtLocalSheet) 
								wait 1
							End  If
							If str_titulo="Grupo de SIM" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", "Estandar"
								wait 1
							End  If
							If str_titulo="Número IMEI" Then
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").DoubleClickCell Iterator, "#1", "LEFT"
								JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaTable("Mostrar atributos:").SetCellData Iterator, "#1", "111111111111111"
								wait 1
								Exit For
							End If
						End Select	
					wait 1
		Next

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Validar").Click
	wait 8
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Tipo_SimCard_"&".png", True
	imagenToWord "Tipo de SIM Card", RutaEvidencias() &Num_Iter&"_"&"Tipo_SimCard_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Siguiente >").Click
	
		tiempo = 0
		Do
		tiempo = tiempo + 1
			If tiempo>=180 Then
		 	 	DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		  	 	DataTable("s_Detalle", dtLocalSheet) = "No se cargo la pantalla correctamente"
		   		 Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
		   		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaMetdosEntrega"&".png" , True
				imagenToWord "Error Carga Metdos Entrega", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaMetdosEntrega"&".png"
			 	ExitActionIteration
			else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
			End If
			wait 1
		Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist))
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
			varprob=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
			varprob= replace(varprob,"<html>","")
			varprob= replace(varprob,"</html>","")
			varprob= replace(varprob,"<br>","")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = varprob
			Reporter.ReportEvent micWarning, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error"&".png" , True
			imagenToWord "Error", RutaEvidencias() &Num_Iter&"_"&"Error"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").Close
			wait 2
			ExitActionIteration
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
			varmsgval = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = varmsgval
			Reporter.ReportEvent micFail , DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error"&".png" , True
			imagenToWord "Error", RutaEvidencias() &Num_Iter&"_"&"Error"&".png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Configuración").JavaButton("Cerrar").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
			wait 4
			ExitActionIteration
		End If 
		
End Sub
Sub TipoEnvio()
	
	Select Case str_metodo_entrega
		Case "En tienda"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("En tienda").Set "ON"
				wait 1
		Case "Delivery"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Delivery").Set "ON"
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png" , True
				imagenToWord "Canal de Venta", RutaEvidencias() &Num_Iter&"_"&"CanalVenta"&".png"
				
					While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Lookup-Validated_2").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png" , True
				imagenToWord "Dirección de Envio", RutaEvidencias() &Num_Iter&"_"&"DireccionEnvio"&".png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Buscar detalles de contacto").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaList("ComboBoxNative$1").Select DataTable ("e_docuDelivery",dtLocalSheet)
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaEdit("TextFieldNative$1").Set DataTable ("e_numDocuDelivery",dtLocalSheet)
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Buscar ahora").Click
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaTable("SearchJTable").SelectRow "#0"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DireccionReceptor"&".png" , True
				imagenToWord "Direccion de Receptor", RutaEvidencias() &Num_Iter&"_"&"DireccionReceptor"&".png"
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de_2").JavaButton("Seleccionar").Click
				wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Exist) = False
						wait 1
					Wend
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Instrucciones del envío:").Set "QA"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaEdit("Número de teléfono del").Set "999999999"
				wait 1
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png", True
				imagenToWord "Entrega Delivery", RutaEvidencias() &Num_Iter&"_"&"EntregaDelivery"&".png"
		Case "Recojo en otra tienda"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaRadioButton("Recojo en otra tienda").Set "ON"
			wait 1
		
	End Select
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Siguiente >").Click
	wait 1
	
'		Do While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución").JavaButton("Lookup-Validated").Exist) = False
'				wait 1	
'				Dim c
'				c=c+1
'					If (c=8) Then exit do		
'		Loop 
'		wait 3
	
		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 1221256A").JavaButton("Pago inmediato").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist)) = False
			wait 1
			
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaObject("WindowsInternalFrameTitlePane_2").Exist(1) Then
					If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Cancelar").Exist(3) Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Cancelar").Click
					End If
				End If   
		Wend
		
		
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
	    var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
	    var1 = replace(var1,"<html>","")
		var1 = replace(var1,"</html>","")
	   ' vardispo2 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(1,1)
	    Reporter.ReportEvent micFail, "Fallido", var1
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png" , True
		imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		ExitActionIteration
	End If 

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		varmsg=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = varmsg
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error"&".png" , True
		imagenToWord "Error", RutaEvidencias() &Num_Iter&"_"&"Error"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar dirección de").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 2
		ExitActionIteration
	End If

	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
	    vardispo1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
	    vardispo2 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(1,1)
	    Reporter.ReportEvent micFail, "Fallido", vardispo2
	    JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Error"&".png" , True
		imagenToWord "Error", RutaEvidencias() &Num_Iter&"_"&"Error"&".png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		ExitActionIteration
	End If   
'	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").Exist=False
'		wait 1
'	Wend
'	Dim text
'	text=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").GetROProperty("text")
'	While text=""
'		wait 1
'		text=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").GetROProperty("text")
'		
'	Wend	
	

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").Exist Then
		
			tiempo = 0
			Do
				tiempo=tiempo+1
					If tiempo>= 620 Then
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					    	DataTable("s_Detalle", dtLocalSheet) = "El Nombre y Dirección del Acuerdo de Facturación no se habilito"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
							JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaNegociarDistribución"&".png" , True
							imagenToWord "Error Carga Negociar Distribución", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaNegociarDistribución"&".png"
							ExitActionIteration
					else
							Reporter.ReportEvent micPass,"Exito","El Nombre y Dirección del Acuerdo de Facturación cargo correctamente"
					End If
			wait 1
			Loop While Not ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("Nombre y Dirección de").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Seleccionar").Exist))
			wait 2
			
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Seleccionar").Exist(2) Then
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Cancelar").Click
		wait 1
	End If		
		
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaEdit("ID del Acuerdo de Facturación:").WaitProperty "editable", 1, 10000
	
			Do While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaButton("Lookup-Validated").Exist) = False
				wait 1
					c=c+1 
					If (c=30) Then exit Do 
			Loop
		wait 1
'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaButton("Lookup-Validated").Exist Then
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png" , True
'			imagenToWord "Negociar Distribución", RutaEvidencias() &Num_Iter&"_"&"Negociar Distribución"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaButton("Lookup-Validated").Click
'		End If
'		
'		Dim t
'		t = 0
'		Do 
'			Wait 1
'			t = t + 1
'			If (t >= 15) Then
'				Exit Do
'			End If
'		Loop While Not (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Siguiente >").Exist)
'		
'		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Seleccionar").Exist Then
'			Dim btnSel, btnReg
'			btnSel = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Seleccionar").GetROProperty("enabled")		
'			If btnSel = "0"	Then
'				btnReg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("0 Registros").GetROProperty("label")
'				If btnReg = "0 Registros" Then
'					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DetalleDistribución"&".png" , True
'					imagenToWord "Detalle Distribución", RutaEvidencias() &Num_Iter&"_"&"DetalleDistribución"&".png"
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_3").JavaButton("Cancelar_2").Click
'					wait 2
'				End If
'			End If	
'			wait 2		
'		End If
		
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaRadioButton("Nuevo").Set
		wait 5	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaList("Mostrar:").Select "Acciones de orden activas "
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Distribución_2").JavaButton("Siguiente >_2").Click	
		wait 5 
		

			tiempo = 0
			Do
				tiempo = tiempo + 1
	
				If tiempo>=680 Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No cargo la pantalla Neogicar Pago"
					Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ErrorCargaNegociarPago"&".png" , True
					imagenToWord "Error Carga Negociar Pago", RutaEvidencias() &Num_Iter&"_"&"ErrorCargaNegociarPago"&".png"
					ExitActionIteration
				else
				Reporter.ReportEvent micPass,"OK","Continuar Flujo"
				End If
				wait 1
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 1221256A").JavaButton("Pago inmediato").Exist
			wait 1
	End If

End Sub
Sub Financiamiento()
    While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaEdit("ID del cliente:").Exist=False
    	wait 1
    Wend
	Dim textID
	textID=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaEdit("ID del cliente:").GetROProperty ("text")		
    While textID=""
    	wait 1
    	textID=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaEdit("ID del cliente:").GetROProperty ("text")
    Wend
    	wait 1
    	

    Dim finExterno
	finExterno =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaCheckBox("Financiamiento Externo").GetROProperty("enabled")
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaCheckBox("Financiamiento Externo").Exist = True and finExterno = "1" Then
          JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaCheckBox("Financiamiento Externo").Set "OFF"
	End If	
    
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaButton("Límite de Compra").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Negociar Pago"&".png" , True
		imagenToWord "Negociar Pago", RutaEvidencias() &Num_Iter&"_"&"Negociar Pago"&".png"
		varfinan=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaButton("Límite de Compra").GetROProperty("enabled")
		wait 1
		If varfinan = "1" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaButton("Límite de Compra").Click	
		End If
		wait 2
	End If
	
		tiempo=0
			Do 
				tiempo=tiempo+1
				varpagoinm=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 1221256A").JavaButton("Pago inmediato").GetROProperty("enabled")
				wait 2	
		Loop While Not (varpagoinm="1")
		wait 1
	
	If str_tipofinan="Financiado" Then
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaCheckBox("Financiamiento Externo").Set "ON"
		
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaEdit("Importe de Cuota Inicial:").Exist)=False
				wait 1
			Wend
			
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaList("Plan de Financiamiento:").Select str_planfinan
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaEdit("Importe de Cuota Mayor:").Set "1"

		varfinan=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaButton("Límite de Compra").GetROProperty("enabled")
		wait 1
		If varfinan = "1" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaButton("Límite de Compra").Click	
		End If
		wait 2
	End If

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 1221256A").JavaButton("Pago inmediato").Click
	wait 1
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Exist) = False
			wait 1
		Wend
	wait 1	
	 While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Nombre a Facturar BAR").Exist=False
	 	wait 1
	 Wend
	 
	 Dim nom
	 nom=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Nombre a Facturar BAR").GetROProperty("text")
	 While nom=""
	 	wait 1
	 	nom=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaEdit("Nombre a Facturar BAR").GetROProperty("text")
	 	
	 Wend
	 
	

	
		Dim Iterator, Count
		Dim rs
	Count = 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").GetROProperty ("items count")
	'MsgBox 	Count
	For Iterator = 0 To Count-1
	 	rs = 	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").GetItem (Iterator)
	 	'MsgBox rs
		If rs = DataTable("e_MedioPago", dtLocalSheet) Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select DataTable("e_MedioPago", dtLocalSheet)
			    wait 1
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Cantidad de cuotas:").Select DataTable("e_Cant_Cuota" , dtLocalSheet)
					wait 1
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Calcular").Click
					JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png" , True
					imagenToWord "Financiamiento", RutaEvidencias() &Num_Iter&"_"&"Financiamiento"&".png"
			
			Exit for
		ElseIf Iterator = Count-1 Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaList("Medio de pago").Select "Externo"
			Exit for
		End if	
	Next
	wait 3
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"DetallePago"&".png" , True
	imagenToWord "Detalle Pago", RutaEvidencias() &Num_Iter&"_"&"DetallePago"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Enviar").Click
	wait 5
		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 1221256A").JavaButton("Siguiente >").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist)) = False
			wait 1
		Wend
	wait 2
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist Then
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Problema"&".png", True
		imagenToWord "No Hay Números Disponible", RutaEvidencias() &Num_Iter&"_"&"Problema"&".png"
		var=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("text")
		var= replace(var,"<html>","")
		var= replace(var,"</html>","")
		var= replace(var,"<br>","")
		var= replace(var,"-&#8203;","")
		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
		DataTable("s_Detalle", dtLocalSheet) = var
		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Pago Inmediato").JavaButton("Cancelar").Click
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 2986153A").JavaButton("Cerrar").Click
		wait 2
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Exist)=False
				wait 1
			Wend
		wait 1	
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("Cancelar orden").Click

			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Pago (Orden 860227A").JavaList("Motivo:").Exist)=False
				wait 1
			Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Pago (Orden 860227A").JavaList("Motivo:").Select "Pedido de Cliente"
		wait 2
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Cancelarorden"&".png", True
		imagenToWord "Cancelar Orden", RutaEvidencias() &Num_Iter&"_"&"Cancelarorden"&".png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Negociar Pago (Orden 860227A").JavaButton("Aceptar").Click
		wait 3
			While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Exist)) =False
				wait 1
			Wend
			
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").Exist Then
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Cancelarorden"&".png", True
			imagenToWord "Cancelar Orden", RutaEvidencias() &Num_Iter&"_"&"Cancelarorden"&".png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaButton("Enviar Cancelar").Click
		End If
		
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Exist Then
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Cancelarorden"&".png", True
			imagenToWord "Cancelar Orden", RutaEvidencias() &Num_Iter&"_"&"Cancelarorden"&".png"
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de cliente empresarial").Close
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Finalizar").Click
			wait 2
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 2
			End If
		End If
		JavaWindow("Ejecutivo de interacción").JavaMenu("Archivo").JavaMenu("Salida").Select
		ExitActionIteration
	End If

	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 1221256A").JavaButton("Siguiente >").Click
	
		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist) or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist)) = False
			wait 1
		Wend
	
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").Exist Then
		var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaTable("SearchJTable").GetCellData(0,1)
		var1 = replace(var1,"<html>","")
		var1 = replace(var1,"</html>","")
		DataTable("s_Resultado",dtLocalSheet) = "Fallido"
		DataTable("s_Detalle",dtLocalSheet)=var1
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensajes de validación").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Negociar Pago (Orden 1221256A").JavaButton("Cerrar").Click
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cerrar negociación de").JavaButton("No guardar").Click
		wait 5
		ExitActionIteration
	End If
End Sub

Sub CodigoVendedor()
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaButton("Validar").Click
	wait 1
	
	Dim text
	text =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaList("ComboBoxNative$1_2").GetROProperty ("text")
    While text= ""
    	wait 1
    	text =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaList("ComboBoxNative$1_2").GetROProperty ("text") 	
    Wend
    wait 1
 
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaList("ComboBoxNative$1").Select "Tienda"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaEdit("TextFieldNative$1").Set "4%%%%%%%"
	wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Buscar ahora").Click
	
	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable_2").Exist = False
		wait 1
	Wend
	
	wait 1
    JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable_2").SelectRow "#0"
    wait 1
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Seleccionar").Click
	
	Dim textCodigo
	textCodigo=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaEdit("Código de vendedor").GetROProperty  ("text")
	While textCodigo=" "
		wait 1
		textCodigo=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaEdit("Código de vendedor").GetROProperty  ("text")
	Wend
End Sub

Sub GeneracionOrden()


	While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Exist = False
		wait 1
	Wend
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
	
	wait 2
	
    If DataTable("e_WIC_Activa", dtLocalSheet) = "SI" Then
				'RunAction "WIC2", oneIteration

		Else 
				
				While JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist = False
				 	wait 1
				Wend

			 If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist= True Then
			 	 	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "LinkDocu.png", True
					imagenToWord "LinkDeDocumentación_"&Num_Iter,RutaEvidencias() & "LinkDocu.png"
					JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
						wait 2
					
					 Dim var
					 var = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").GetROProperty("enabled")
					If var = "1"  Then
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
					End If
					
			End If

			
	End If

'	
'	tiempo = 0
'	Do
'		While((JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist)) = False
'			wait 1
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Validade y Ver Contrato").Click
'			tiempo = tiempo + 1
'			If DataTable("e_WIC_Activa",dtLocalSheet)="SI" Then
'				'RunAction "WIC2", oneIteration
'				Exit Do
'			End If
'			wait 3
'		Wend
'		
'		If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
'			wait 3
'			var1 = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("attached text")
'			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ContratosNoGenerados_"&".png", True
'			imagenToWord "Contratos no Generados", RutaEvidencias() &Num_Iter&"_"&"ContratosNoGenerados_"&".png"
'	   	 	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
'	   	 	wait 1
'		End  If
'		
'		If JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Exist Then
'			JavaWindow("Ejecutivo de interacción").JavaDialog("Error interno").Close
'			wait 2
'		End If
'		wait 1
'			
'			If tiempo>=180 Then
'				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'				DataTable("s_Detalle", dtLocalSheet) = "No se a cargado el contrato correctamente"
'				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
'				ExitActionIteration
'			else
'				Reporter.ReportEvent micPass,"Contrato Exitoso","Se a cargado el contrato correctamente"
'			End If
'	wait 1
'	Loop While Not (JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist or (var1 = "Contratos no Generados") or (var1 = "0"))
'
'	If JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Exist Then
'		JavaWindow("Ejecutivo de interacción").JavaDialog("Resumen de la orden (Orden").Close
'		wait 2
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaCheckBox("El cliente firmó.").Set "ON"
'	End If

		tiempo = 0
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
		
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Exist(2) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
		Wait 1
			If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Por favor valide el Código").Exist = True Then
				JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
				wait 1
				Call CodigoVendedor()
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Enviar orden").Click
			
			End If
	End If

		While((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)) = False
			wait 1
		Wend
	wait 3
	If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist Then
		varvend=JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
		If varvend="Por favor valide el código del Vendedor." Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click	
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaButton("Validar").Click
'			wait 1
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaEdit("TextFieldNative$1").Set "41523813"
'			wait 1
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaTable("SearchJTable_2").SelectRow "#0"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden").JavaButton("Seleccionar").Click
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de la orden (Orden_2").JavaButton("Enviar orden").Click
			wait 2
		End If
	End  If
	
	
	wait 2
	DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").GetROProperty("text")
	flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
	DataTable("s_Nro_Orden", dtLocalSheet) = replace (DataTable("s_Nro_Orden", dtLocalSheet),"Orden ","")
	Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
	imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
	wait 1

'	Select Case DataTable("e_Ambiente", "Login [Login]")
'		Case "UAT8"
'			DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),8)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
'			imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
'			wait 1
'		Case "UAT4"
'			DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
'			imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
'			wait 1
'		Case "UAT3"
'			DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),8)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
'			imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
'			wait 1
'		Case "UAT6"
'			DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
'			imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
'			wait 1	
'		Case "UAT10"
'			DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
'			imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
'			wait 1
'		Case "UAT13"
'			DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),7)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
'			imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
'			wait 1	
'		Case "PROD"
'			DataTable("s_Nro_Orden", dtLocalSheet) =JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").GetROProperty("text")
'			flag = InStr(DataTable("s_Nro_Orden", dtLocalSheet), "correctamente")
'			DataTable("s_Nro_Orden", dtLocalSheet) = Right (DataTable("s_Nro_Orden", dtLocalSheet),10)
'			Reporter.ReportEvent micPass, "Numero de Orden", "Se generó el Numero de Orden: "&DataTable("s_Nro_Orden", dtLocalSheet)
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png", True
'			imagenToWord "Orden Generada", RutaEvidencias() &Num_Iter&"_"&"Orden_Generada_"&".png"
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Orden 1207161A").JavaButton("Cerrar").Click
'			wait 1
'	End Select
	
		If str_metodo_entrega="Delivery" Then
			ExitActionIteration
		End If
	
End Sub
Sub PagoManual()

'	If (str_mediopago<>"Pago a la Factura") Then
'		wait 1
			tiempo=0
			Do 
				tiempo=tiempo+1
				vardepo=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").GetROProperty("enabled")
				wait 2
			Loop While Not (vardepo="1")
		
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_19748072_;_script infofile_;_ZIP::ssf61.xml_;_
		wait 1
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
			wait  1
		Wend
		wait 1
		
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
			wait 2
		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
		wait 3
	
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
			If (tiempo >= 120) Then
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Pago de la Orden"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					ExitActionIteration
			End If
		Loop While Not(nroreg="1 Registros")
		wait 1
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Exist) = False
			wait  1
		Wend
		
		var = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").GetROProperty("enabled")
			tiempo=0
				While(var <> "0") = False	
					wait 1
					tiempo=tiempo+1
					If (tiempo >= 180) Then
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").Close
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				  			DataTable("s_Detalle", dtLocalSheet) = "La orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" ya fue pagado"
							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").Close
							ExitActionIteration
					End If
				Wend
		
		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
		Reporter.ReportEvent micDone, "Pago Correcto", "El número de orden : "&DataTable("s_Nro_Orden", dtLocalSheet)&" fué correctamente pagado"

'	End If
	
'	If (str_tipofinan<>"Financiado")  Then
'		wait 1
'			tiempo=0
'			Do 
'				tiempo=tiempo+1
'				vardepo=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").GetROProperty("enabled")
'				wait 2
'			Loop While Not (vardepo="1")
'		
'		wait 1
'		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
'		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_19748072_;_script infofile_;_ZIP::ssf61.xml_;_
'		wait 1
'		
'		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
'			wait  1
'		Wend
'		wait 1
'		
'		Do
'			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
'			wait 2
'		Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
'		wait 2
'	
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
'		wait 2
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
'		wait 2
'		
'		tiempo=0
'		Do 
'			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
'				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
'				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("attached text")
'				tiempo=tiempo+1
'				wait 1
'			End If
'			If (tiempo >= 120) Then
'					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Pago de la Orden"
'					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
'					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
'					ExitActionIteration
'			End If
'		Loop While Not(nroreg="1 Registros")
'		wait 1
'		
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
'		wait 1
'		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
'		
'		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Exist) = False
'			wait  1
'		Wend
'		
'		var = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").GetROProperty("enabled")
'			tiempo=0
'				While(var <> "0") = False	
'					wait 1
'					tiempo=tiempo+1
'					If (tiempo >= 180) Then
'							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").Close
'							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
'				  			DataTable("s_Detalle", dtLocalSheet) = "La orden: "&DataTable("s_Nro_Orden", dtLocalSheet)&" ya fue pagado"
'							Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
'							JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes").Close
'							ExitActionIteration
'					End If
'				Wend
'		
'		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_2").JavaButton("Enviar").Click
'		Reporter.ReportEvent micDone, "Pago Correcto", "El número de orden : "&DataTable("s_Nro_Orden", dtLocalSheet)&" fué correctamente pagado"
'
'	End If
End Sub
Sub GestionLogistica()
		
		tiempo=0
			Do
				varlogis=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").GetROProperty("enabled")
				wait 2
			Loop While Not (varlogis="1")

		wait 1
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		wait 1
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist)=False
			wait 1
		Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
		wait 1
		
		tiempo=0
		Do 
			If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Exist Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
				nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("0 Registros").GetROProperty("attached text")
				tiempo=tiempo+1
				wait 1
			End If
			If (tiempo >= 120) Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"
					Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
					ExitActionIteration
			End If
		Loop While Not (nroreg="1 Registros")
		wait 2

		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").SelectRow "#0"
		
		tiempo=0
			Do
				If (DataTable("s_Detalle", dtLocalSheet)="Por favor rellenar todas las identificaciones de equipos") or (DataTable("s_Detalle", dtLocalSheet)<>"Por favor rellenar todas las identificaciones de equipos") Then
					If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Exist(1) Then
						JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Cancelar").Click
						wait 2
					End If
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Gestionar logística").Click
					tiempo=tiempo+1
					wait 1
					While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").Exist) = False
						wait 1
					Wend
					
						If str_tipoalta="Alta Nueva Equipo + SIM" Then
								vardisp=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (1,1)
								
								If vardisp="Dispositivo" Then
								
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_Dispositivo", dtLocalSheet)
									wait 2
								Else 
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#1","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#1","#4",DataTable("e_ID_SIM", dtLocalSheet)
									wait 2
									
									
								End If
					
								varsim=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").GetCellData (2,1)
								
								If varsim="Tarjeta SIM" Then
								
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",DataTable("e_ID_SIM", dtLocalSheet)
									wait 2
								Else 
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").DoubleClickCell "#2","#4"
									Set shell = CreateObject("Wscript.Shell") 
									shell.SendKeys "{ENTER}"
									JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaTable("SearchJTable").SetCellData "#2","#4",DataTable("e_ID_Dispositivo", dtLocalSheet)
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
						
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png", True
					imagenToWord "Ingreso de Materiales", RutaEvidencias() &Num_Iter&"_"&"Ingreso_Materiales_"&".png"
					wait 1
					JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Validar y Crear Factura").Object.doClick()
					 wait 1
					
					tiempo = 0
					Do
						tiempo=tiempo+1
							varhab=JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").GetROProperty("enabled")					
							wait 3
					Loop While Not((JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist) Or (varhab="1"))
					
				
						'If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(1) or Window("Ejecutivo de interacción").Window("Buscar: Orden > Solicitar").Window("Mensaje").Exist(1) or JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
						If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(1) or JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(1) Then
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").Exist(0) Then
								varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text") 
							End If
							If JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist(0) Then
								varlog = JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
							End If
							DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				       		DataTable("s_Detalle", dtLocalSheet) = varlog
				       		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				     		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
				     			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png", True
								imagenToWord "Mensaje de Validación", RutaEvidencias() &Num_Iter&"_"&"MensajeValidacion"&".png"
				        		JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaDialog("Mensaje").JavaButton("OK").Click
							End If
							If 	JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Exist(1) Then
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
								If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist(1) Then
									JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
									ExitActionIteration
								End If
				     		End  If
				    	End If
				End  If
				If tiempo>=20 Then
					Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)  
					DataTable("s_Resultado",dtLocalSheet) = "Fallido"
					DataTable("s_Detalle",dtLocalSheet) = "Luego de 20 intentos no se pudo realizar la Asignación de Series"
					ExitActionIteration
				else
					Reporter.ReportEvent micPass, "Exito", "Se realizo la Asignación de Series correctamente"
			End If
		Loop While Not varhab = "1"
		wait 2
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Exist(2) Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Orden > Solicitar").JavaButton("Enviar").Click
		End If
End Sub
Sub EmpujeOrden()
	
		If DataTable("e_Tipo_De_DATA_Sim", dtLocalSheet) = "DATA LOGICA" Then
		
			tiempo=0
			Do 
				tiempo=tiempo+1
				vardepo=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").GetROProperty("enabled")
				wait 2
			Loop While Not (vardepo="1")
		
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
			JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Depósito de Ordenes").Select @@ hightlight id_;_19748072_;_script infofile_;_ZIP::ssf61.xml_;_
			wait 1 @@ hightlight id_;_19748072_;_script infofile_;_ZIP::ssf61.xml_;_
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Exist) = False
				wait 1
			Wend
			
			
			Do
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTab("Equipo usuario:").Select "Tareas pendientes del equipo"
				wait 2
			Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Finalizar compra y activar").Exist
			wait 3
			
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").SetFocus
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet)
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
			wait 2
			
			tiempo=0
			Do 
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					nroreg = JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("1 Registros").GetROProperty("text")
					tiempo=tiempo+1
					wait 1
				End If
				
				If (tiempo >= 120) Then
						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
						DataTable("s_Detalle", dtLocalSheet) = "No se encuentra la orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" para realizar el Empuje de la Orden"
						Reporter.ReportEvent micFail,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
						wait 2
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
						wait 2
						ExitActionIteration
						wait 2
				End If
			
			Loop While Not(nroreg="1 Registros")
			wait 1
			
			tiempo=0
			Do
				If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Exist Then
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Buscar ahora").Click
					wait 2
					tiempo = tiempo+1
					JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").Output CheckPoint("Equipo usuario:") @@ hightlight id_;_22747135_;_script infofile_;_ZIP::ssf1.xml_;_
					varValidaRespuestaCumplimiento = Environment("s_ValidaManejarRespuestaCumplimiento")
					wait 1
				End If
					If (tiempo >= 120) Then
						DataTable("s_Resultado",dtLocalSheet)="Fallido"
						DataTable("s_Detalle",dtLocalSheet)="La actividad 'Manejar Respuesta de Cumplimiento' no cargo"	
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
						JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
						ExitTestIteration
					End If 
			Loop While Not varValidaRespuestaCumplimiento = "Manejar Respuesta de Cumplimiento"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaTable("Equipo usuario:").SelectRow "#0"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Gestión manual").Click
			
			While(JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_3").JavaList("Estado de la gestión manual:").Exist) = False
				wait  1
			Wend
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_3").JavaList("Estado de la gestión manual:").Select "Cumplimiento Completo"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_3").JavaList("Motivo de la gestión manual").Select "Manejo manual: Manejo Manual OSS"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaDialog("Buscar: Grupo de órdenes_3").JavaButton("Enviar").Click
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").Exist) = False
				wait  1
			Wend
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
End Sub
Sub OrdenCerrado()

		tiempo=0
			Do
				tiempo=tiempo+1
				varlogis=JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").GetROProperty("enabled")
				wait 2
			Loop While Not (varlogis="1")
			
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Buscar").JavaMenu("Pedidos").JavaMenu("Órdenes").Select
		wait 1 @@ hightlight id_;_27981779_;_script infofile_;_ZIP::ssf1.xml_;_
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Exist) = False
			wait 1
		Wend

		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaEdit("TextFieldNative$1").Set DataTable("s_Nro_Orden", dtLocalSheet) @@ hightlight id_;_19889480_;_script infofile_;_ZIP::ssf3.xml_;_
		wait 1 @@ hightlight id_;_22342896_;_script infofile_;_ZIP::ssf8.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaCheckBox("Solo órdenes pendientes").Set "OFF" @@ hightlight id_;_23592401_;_script infofile_;_ZIP::ssf22.xml_;_
		wait 1
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click @@ hightlight id_;_12440768_;_script infofile_;_ZIP::ssf4.xml_;_
		wait 8
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:") @@ hightlight id_;_32738597_;_script infofile_;_ZIP::ssf67.xml_;_
		Reporter.ReportEvent micPass,"Se valida el estado de la orden",  DataTable("s_ValEstadoOrden", dtLocalSheet)
		
		tiempo = 0
		Do
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Buscar ahora").Click
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").Output CheckPoint("Ver por:") @@ hightlight id_;_24634058_;_script infofile_;_ZIP::ssf1.xml_;_
			tiempo = tiempo + 1
				If (tiempo>=150) Then		
					DataTable("s_Resultado", dtLocalSheet) = "Fallido"
					DataTable("s_Detalle", dtLocalSheet) = "La Orden:"&DataTable("s_Nro_Orden", dtLocalSheet)&" no culmino en estado Cerrado"
					Reporter.ReportEvent micFail,"Error al finalizar la orden","Es probable que la orden termine con tiempo excedido"
					If DataTable("s_ValEstadoOrden", dtLocalSheet) = "Enviado" Then
						Exit Do
						wait 1
					End If	
					'ExitActionIteration
				else
				Reporter.ReportEvent micPass, "Se valida el estado de la orden",  DataTable("s_ValEstadoOrden", dtLocalSheet)
				End If
			Loop While Not DataTable("s_ValEstadoOrden", dtLocalSheet) = "Cerrado"
			DataTable("s_Resultado", dtLocalSheet)="Exito"
			DataTable("s_Detalle", dtLocalSheet)="La orden finalizó correctamente"
			Reporter.ReportEvent micPass,DataTable("s_Resultado", dtLocalSheet),DataTable("s_Detalle", dtLocalSheet)

		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Orden_Cerrada_"&".png", True
		imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"Orden_Cerrada_"&".png"

End Sub
Sub DetalleActividadOrden()
	wait 2
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaList("Ver por:").Select "Acciones de orden"
	wait 3
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaTable("Ver por:").DoubleClickCell 0, "#8", "LEFT"
	Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{ENTER}"
	wait 1
	
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2956696A").JavaEdit("Fecha de vencimiento:").Exist)=False
		wait 1
	Wend
	
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2956696A").JavaTab("Nombre del cliente:").Select "Actividad"
	
	While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2956696A").JavaTable("SearchJTable").Exist)=False
		wait 1	
	Wend
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_1"&".png"
	
	shell.SendKeys "{PGDN}"
	wait 1
	
	JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png", True
	imagenToWord "Orden Cerrada", RutaEvidencias() &Num_Iter&"_"&"ActividadesdeOrden_2"&".png"
	
	filas=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2956696A").JavaTable("SearchJTable").GetROProperty("rows")
	For Iterator = 0 To filas-1
		varselec=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2956696A").JavaTable("SearchJTable").GetCellData(Iterator,0)
	Next
	
	If varselec<>"Cerrar Acción de Orden" Then
	 	DataTable("s_Resultado",dtLocalSheet)="Fallido"
		DataTable("s_Detalle",dtLocalSheet)="La orden "&DataTable("s_Nro_Orden",dtLocalSheet)&" culmino en la Actividad "&varselec&""
		Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2956696A").JavaButton("Cancelar").Click

		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
			wait 2
		End If
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(1) Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
		End If
		ExitActionIteration
		wait 1
	End If
	JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Ver acción de orden: 2956696A").JavaButton("Cancelar").Click

	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Exist Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Orden").JavaButton("Cerrar").Click
		wait 2
	End If
	
	If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Exist(1) Then
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
	End If
	
End Sub




