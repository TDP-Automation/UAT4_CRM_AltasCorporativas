Option Explicit

Dim varerror, vartext, varresumen, varacuerdo, varfactu, varmens, filas
Dim varenable
Dim tiempo
Dim x
Dim var_GN

Dim str_nom_acuer
Dim str_con_corp
Dim str_plan_comp
Dim str_modeloCel
Dim str_tipoPlan
Dim str_alquiler
Dim str_dias
Dim str_grupo_negocio

str_nom_acuer 		= DataTable("e_Nombre_Acuerdo", dtLocalSheet)
str_con_corp  		= DataTable("e_Contrato_Corporativo", dtLocalSheet) 
str_plan_comp 		= DataTable("e_Plan_Compartido", dtLocalSheet)
str_modeloCel 		= DataTable("e_ModeloCelular", dtLocalSheet)
str_tipoPlan  		= DataTable("e_Plan_Tarifario", dtLocalSheet)
str_alquiler  		= DataTable("e_Alquiler", dtLocalSheet)
str_dias      		= DataTable("e_Dias", dtLocalSheet)
str_grupo_negocio   = DataTable("e_Grupo_Negocio", dtLocalSheet)

If str_grupo_negocio=Empty Then
	Call CrearAcuerdoComercial()
	Call AgregarAcuerdoFacturacion()
	Call ConfigurarTerminosGenerales()
	Call AgregarPlanCompartido()
	Call AnadirGrupoNegocio()
	Call AnadirPlanTarifario()
	Call AnadirEquipoMovil()
	Call LineaAlquiler()
	Call EnviarOrdenAcuerdo()
	
End If
wait 1

Sub CrearAcuerdoComercial()

		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaStaticText("Número de documento(st)").Exist )= False
				wait 1
		Wend
		wait 2
			
			tiempo=0
				Do
						tiempo=tiempo+1
						varacuerdo=JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").JavaMenu("Pedidos").JavaMenu("Acuerdo Comercial").GetROProperty("enabled")
						wait 2
				Loop While Not (varacuerdo="1")

		wait 1
		JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").JavaMenu("Pedidos").Select
		JavaWindow("Ejecutivo de interacción").JavaMenu("Crear").JavaMenu("Pedidos").JavaMenu("Acuerdo Comercial").Select
		wait 1
		
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaStaticText("Acuerdo Comercial(st)").Exist )= False
				wait 1
		Wend
		
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Missing Permissions").JavaButton("OK").Exist Then
			varerror=JavaWindow("Ejecutivo de interacción").JavaDialog("Missing Permissions").JavaObject("JPanel").GetROProperty("attached text")
				DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				DataTable("s_Detalle", dtLocalSheet) = varerror
				Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Error.png", True
				imagenToWord "Error de Permisos", RutaEvidencias() & "Error.png"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Missing Permissions").JavaButton("OK").Click
				SystemUtil.CloseProcessByName "javaw.exe"
				SystemUtil.CloseProcessByName "jp2launcher.exe"
				ExitTest
		End If	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Renombrar la Orden").WaitProperty "enabled", true, 10000 @@ hightlight id_;_19402115_;_script infofile_;_ZIP::ssf8.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Renombrar la Orden").Click @@ hightlight id_;_2841166_;_script infofile_;_ZIP::ssf3.xml_;_
		While (JavaWindow("Ejecutivo de interacción").JavaDialog("Cambiar el nombre de Acuerdo").JavaEdit("Nombre del pedido:").Exist )= False
				wait 1
		Wend
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cambiar el nombre de Acuerdo").JavaEdit("Nombre del pedido:").Set str_nom_acuer @@ hightlight id_;_7649579_;_script infofile_;_ZIP::ssf4.xml_;_
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "NombreAcuerdo.png", True
		imagenToWord "Nombre del Acuerdo", RutaEvidencias() & "NombreAcuerdo.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Cambiar el nombre de Acuerdo").JavaButton("Actualizar").Click
		
		While (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Agregar Términos Generales").Exist )= False
			wait 1
		Wend
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Agregar Términos Generales").Click @@ hightlight id_;_17276892_;_script infofile_;_ZIP::ssf6.xml_;_
		While ((JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaEdit("Condiciones Generales").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").Exist)) = False
			wait 1
		Wend
		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").Exist(2) Then
			varerror=JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaObject("JPanel").GetROProperty("attached text")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = varerror
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "Error.png", True
			imagenToWord "Error", RutaEvidencias() & "Error.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Problema").JavaButton("OK").Click
			SystemUtil.CloseProcessByName "javaw.exe"
			SystemUtil.CloseProcessByName "jp2launcher.exe"
			ExitTest
		End If
	
		If JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaDialog("Mensaje").Exist Then
			varmens=JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaDialog("Mensaje").JavaObject("JPanel").GetROProperty("text")
			DataTable("s_Resultado", dtLocalSheet) = "Fallido"
			DataTable("s_Detalle", dtLocalSheet) = varmens
			Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "MensajeValidacion.png", True
			imagenToWord "Error", RutaEvidencias() & "MensajeValidacion.png"
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaDialog("Mensaje").JavaButton("OK").Click
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaButton("Descartar").Click
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Cerrar").Click
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Panel de Interacción").JavaButton("Cerrar").Click
			wait 1
			ExitTest
		End If
	
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaEdit("Condiciones Generales").Set str_con_corp @@ hightlight id_;_1246132_;_script infofile_;_ZIP::ssf7.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaButton("Buscar").Click @@ hightlight id_;_5787833_;_script infofile_;_ZIP::ssf8.xml_;_
		wait 2
				tiempo = 0
				Do
					tiempo = tiempo + 1
						If tiempo >= 180 Then
								DataTable("s_Resultado", dtLocalSheet) = "Fallido"
						   	 	DataTable("s_Detalle", dtLocalSheet) = "El contrato corporativo no se encuentra disponible"
								Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
								ExitActionIteration
						else
								Reporter.ReportEvent micPass, "Busqueda Exitosa", "La busqueda resulto exitosa"
						End If
				Loop While Not JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaCheckBox("Seleccionar").Exist(2)

		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaCheckBox("Seleccionar").Set "ON" @@ hightlight id_;_17428215_;_script infofile_;_ZIP::ssf9.xml_;_
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ContratoCorporativo.png", True
		imagenToWord "Contrato Corporativo", RutaEvidencias() & "ContratoCorporativo.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Términos Generales").JavaButton("Agregar Elemento Seleccionado").Click @@ hightlight id_;_7728626_;_script infofile_;_ZIP::ssf10.xml_;_
		wait 2

		While ((JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Realizar Distribuición").Exist) Or (JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Agregar Acuerdo de Facturación").Exist)) = False
			wait 1
		Wend
End Sub 
Sub AgregarAcuerdoFacturacion()

		wait 2
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Realizar Distribuición").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Realizar Distribuición").Click
			wait 1
		End If
		
		While JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaStaticText("<html>Términos Generales").Exist = False
		 wait 1
		Wend
		
			
		If JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Agregar Acuerdo de Facturación").Exist Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Agregar Acuerdo de Facturación").Click	
			wait 1
		End If
		
		wait 2
				tiempo = 0
				Do
					tiempo=tiempo+1
							If tiempo>= 220 Then
									DataTable("s_Resultado", dtLocalSheet) = "Fallido"
							    	DataTable("s_Detalle", dtLocalSheet) = "El Nombre y Dirección del Acuerdo de Facturación no se habilito"
									Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
									ExitActionIteration
							else
									Reporter.ReportEvent micPass,"Exito","El Nombre y Dirección del Acuerdo de Facturación cargo correctamente"
							End If
				wait 1
				Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial_2").JavaEdit("Nombre y Dirección de").Exist(2)
	
					tiempo=0
						While(vartext<>"") = False
							wait 1
							tiempo=tiempo+1
							vartext=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial_2").JavaEdit("Nombre y Dirección de").GetROProperty("text")
							If(tiempo>=160) Then
									DataTable("s_Resultado", dtLocalSheet) = "Fallido"
									DataTable("s_Detalle", dtLocalSheet) = "No carga el Nombre y Dirección de Facturación"
										Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet), DataTable("s_Detalle", dtLocalSheet)
										JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Buscar: Grupo de órdenes").JavaButton("Cerrar").Click
										ExitActionIteration
							End  If
						Wend
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial_2").JavaButton("Lookup-Validated").WaitProperty "enabled", true, 10000
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial_2").JavaRadioButton("Nuevo").Set
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "AcuerdoFacturacion.png", True
		imagenToWord "Asignación Acuerdo Facturación", RutaEvidencias() & "AcuerdoFacturacion.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial_2").JavaButton("Asignar acuerdo de facturación").Click @@ hightlight id_;_26086573_;_script infofile_;_ZIP::ssf13.xml_;_
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Configurar").Exist)=False
			wait 1
		Wend
	
End Sub
Sub ConfigurarTerminosGenerales()
		
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Configurar").Click @@ hightlight id_;_32908966_;_script infofile_;_ZIP::ssf15.xml_;_
	
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaButton("Agregar").Exist)=False
			wait 1
		Wend
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTree("Subproductos disponibles").Expand "#0;Términos Generales e Información (GTI)"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTree("Subproductos disponibles").Select "#0;Términos Generales e Información (GTI);Cargo único por Activación de M2M" @@ hightlight id_;_10399305_;_script infofile_;_ZIP::ssf16.xml_;_
		wait 3
		JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaButton("Agregar").Click
		wait 2
		filas=JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").GetROProperty("rows")
		Select Case filas
			Case "6"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").DoubleClickCell 2, 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").SetCellData 2,1,"1"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").DoubleClickCell 3, 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").SetCellData 3,1,"1"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").DoubleClickCell 4, 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").SetCellData 4,1,"1"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").DoubleClickCell 5, 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").SetCellData 5,1,"1"
				wait 1
			Case "5"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").DoubleClickCell 2, 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").SetCellData 2,1,"1"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").DoubleClickCell 3, 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").SetCellData 3,1,"1"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").DoubleClickCell 4, 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").SetCellData 4,1,"1"
				wait 1
		End Select
		If DataTable("e_Alquiler",dtLocalSheet) = "Alquiler" Then
			JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").DoubleClickCell 5, 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Mostrar atributos:").SetCellData 5,1,"1"	
		End If
		
		JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaButton("Validar").Click
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaTable("Notificaciones").Exist)=True
			wait 1
		Wend
			
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ConfigurarTérminosGenerales.png", True
		imagenToWord "Configurar Términos Generales", RutaEvidencias() & "ConfigurarTérminosGenerales.png"
		wait 1		
		JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar Términos Generales").JavaButton("Guardar  Cerrar").Click @@ hightlight id_;_30680188_;_script infofile_;_ZIP::ssf23.xml_;_
		
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Configurar").Exist)=False
			wait 1
		Wend
	
End Sub
Sub AgregarPlanCompartido()

		If str_nom_acuer="ACUERDO CON BOLSA" Then
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Agregar plan compartido").Click @@ hightlight id_;_19105288_;_script infofile_;_ZIP::ssf24.xml_;_
				
				While(JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Plan de Servicio").JavaEdit("Planes Compartidos Disponibles").Exist)=False
					wait 1
				Wend
				JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Plan de Servicio").JavaEdit("Planes Compartidos Disponibles").SetFocus
				JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Plan de Servicio").JavaEdit("Planes Compartidos Disponibles").Set str_plan_comp
				JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Plan de Servicio").JavaButton("Buscar").Click @@ hightlight id_;_18972276_;_script infofile_;_ZIP::ssf26.xml_;_

					tiempo=0
						Do
						tiempo=tiempo+1
								If tiempo>=180 Then
										DataTable("s_Resultado",dtLocalSheet)="Fallido"
										DataTable("s_Detalle",dtLocalSheet)="La Busqueda de 'Planes Tarifarios' no mostros registros"
										Reporter.ReportEvent micFail, DataTable("s_Resultado",dtLocalSheet), DataTable("s_Detalle",dtLocalSheet)
										ExitActionIteration
								else	
										Reporter.ReportEvent micPass, "Exito","La búsqueda de Planes Tarifario resultdo exitosa"
								End If
									wait 1
						Loop While Not JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Plan de Servicio").JavaCheckBox("Seleccionar").Exist(1)
					
				JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Plan de Servicio").JavaCheckBox("Seleccionar").Set "ON" @@ hightlight id_;_31436341_;_script infofile_;_ZIP::ssf27.xml_;_
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "PlanCompartido.png", True
				imagenToWord "Plan Compartido", RutaEvidencias() & "PlanCompartido.png"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccionar Plan de Servicio").JavaButton("Agregar Elemento Seleccionado").Click
				wait 2
				While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Configurar_2").Exist)=False
					wait 1
				Wend	
				wait 7			
				JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Configurar_2").Click @@ hightlight id_;_32652698_;_script infofile_;_ZIP::ssf29.xml_;_
				
				While(JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaButton("Expandir todo").Exist)=False
					wait 1
				Wend
				
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTree("Subproductos disponibles").Select "#0;Plan de Servicios Compartidos (SSP);Bolsa de Planes de Voz Compartidos" @@ hightlight id_;_8309088_;_script infofile_;_ZIP::ssf30.xml_;_
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaButton("Agregar").Click @@ hightlight id_;_14454280_;_script infofile_;_ZIP::ssf31.xml_;_
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTree("Subproductos disponibles").Expand "#0;Plan de Servicios Compartidos (SSP);Bolsa de Planes de Voz Compartidos" @@ hightlight id_;_8309088_;_script infofile_;_ZIP::ssf32.xml_;_
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTree("Subproductos disponibles").Select "#0;Plan de Servicios Compartidos (SSP);Bolsa de Planes de Voz Compartidos;Bolsa Comunidad Movistar"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaButton("Agregar").Click @@ hightlight id_;_14454280_;_script infofile_;_ZIP::ssf34.xml_;_
				wait 2
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTable("Mostrar atributos:").DoubleClickCell 4, 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTable("Mostrar atributos:").SetCellData 4,1,"1" @@ hightlight id_;_32998441_;_script infofile_;_ZIP::ssf35.xml_;_
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTable("Mostrar atributos:").DoubleClickCell 5,1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTable("Mostrar atributos:").SetCellData 5,1,"1" @@ hightlight id_;_32998441_;_script infofile_;_ZIP::ssf36.xml_;_
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTable("Mostrar atributos:").DoubleClickCell 6,1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaTable("Mostrar atributos:").SetCellData 6,1,"1"
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaButton("Validar").Click @@ hightlight id_;_22873928_;_script infofile_;_ZIP::ssf37.xml_;_
				wait 5
				JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ConfiguracionPlanCompartido.png", True
				imagenToWord "Configuración Plan Compartido", RutaEvidencias() & "ConfiguracionPlanCompartido.png"
				wait 1
				JavaWindow("Ejecutivo de interacción").JavaDialog("Configurar el plan de").JavaButton("Guardar  Cerrar").Click @@ hightlight id_;_8969145_;_script infofile_;_ZIP::ssf38.xml_;_
				
				While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Ańadir Grupo de negocio").Exist)=False
					wait 1
				Wend
			End If
End Sub
Sub AnadirGrupoNegocio()
	
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Ańadir Grupo de negocio").Click @@ hightlight id_;_3394965_;_script infofile_;_ZIP::ssf39.xml_;_
		x=RandomNumber (100,1000)
		var_GN= "ACUERDO COMERCIAL "+CStr(x)
		DataTable("e_Grupo_Negocio",dtLocalSheet)=var_GN
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Grupo de Negocio").JavaEdit("Nombre").Exist)=False
			wait 1
		Wend
	
		JavaWindow("Ejecutivo de interacción").JavaDialog("Grupo de Negocio").JavaEdit("Nombre").Set DataTable("e_Grupo_Negocio", dtLocalSheet)
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "NombreAcuerdoComercial.png", True
		imagenToWord "Nombre del Acuerdo Comercial", RutaEvidencias() & "NombreAcuerdoComercial.png"
		JavaWindow("Ejecutivo de interacción").JavaDialog("Grupo de Negocio").JavaButton("Ańadir Grupo de negocio").Click @@ hightlight id_;_20769500_;_script infofile_;_ZIP::ssf41.xml_;_
	
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Ańadir Plantilla de Oferta").Exist)=False
			wait 1
		Wend
End Sub
Sub AnadirPlanTarifario()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Ańadir Plantilla de Oferta").Click @@ hightlight id_;_24777236_;_script infofile_;_ZIP::ssf42.xml_;_
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione plantilla de").JavaEdit("Ofertas de Negocios Disponible").Exist)=False
			wait 1
		Wend
		
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione plantilla de").JavaEdit("Ofertas de Negocios Disponible").SetFocus
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione plantilla de").JavaEdit("Ofertas de Negocios Disponible").Set str_tipoPlan @@ hightlight id_;_18741357_;_script infofile_;_ZIP::ssf43.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione plantilla de").JavaButton("Buscar").Click @@ hightlight id_;_10534324_;_script infofile_;_ZIP::ssf46.xml_;_

				tiempo = 0
				Do
				tiempo = tiempo + 1
					If tiempo >=180  Then
						DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				    	DataTable("s_Detalle", dtLocalSheet) = "No se encuentro planes tarifarios disponibles"
						Reporter.ReportEvent micFail , DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
						ExitActionIteration
					else
				   	 Reporter.ReportEvent micPass, "Exito", "Se encontro el equipo móvil filtrado"
					End If
				wait 1
				Loop While Not JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione plantilla de").JavaCheckBox("Seleccionar").Exist(1)
			
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione plantilla de").CaptureBitmap RutaEvidencias() & "PlanTarifario.png", True
		imagenToWord "Plan Tarifario", RutaEvidencias() & "PlanTarifario.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione plantilla de").JavaCheckBox("Seleccionar").Set "ON" @@ hightlight id_;_4961342_;_script infofile_;_ZIP::ssf47.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione plantilla de").JavaButton("Ańadir oferta seleccionada").Click @@ hightlight id_;_6132322_;_script infofile_;_ZIP::ssf48.xml_;_
		
				tiempo = 0
				Do
				tiempo = tiempo + 1
					If tiempo >= 120 Then
				  		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				   		DataTable("s_Detalle", dtLocalSheet) = "No se habilito el boton Agregar Dispositivos"
				   		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				   		ExitActionIteration
				   else
				   		Reporter.ReportEvent micPass, "Exito", "Se habilito el boton Agregar Dispositivos"
					End If
				wait 1
				Loop While Not JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Agregar dispositivos").Exist(1)
End Sub
Sub AnadirEquipoMovil()
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Agregar dispositivos").Click @@ hightlight id_;_13673009_;_script infofile_;_ZIP::ssf49.xml_;_
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").JavaEdit("TextFieldNative$1").Exist)=False
			wait 1
		Wend
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").JavaEdit("TextFieldNative$1").Set str_modeloCel @@ hightlight id_;_8785105_;_script infofile_;_ZIP::ssf51.xml_;_
		wait 5
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").JavaButton("Buscar").Click @@ hightlight id_;_11774235_;_script infofile_;_ZIP::ssf52.xml_;_
		wait 2
		If str_modeloCel<>"HUAWEI P10 NEGRO" Then
			tiempo = 0
				Do
				tiempo = tiempo + 1
					If tiempo >= 180 Then
				  		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				   		DataTable("s_Detalle", dtLocalSheet) = "No se encuentra el equipo movil de la busqueda"
				   		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				   		ExitActionIteration
				   else
				   		Reporter.ReportEvent micPass, "Exito", "Se habilito el boton Agregar Dispositivos"
					End If
				wait 1
				Loop While Not JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").JavaCheckBox("Seleccionar").Exist(0)
			wait 3
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").CaptureBitmap RutaEvidencias() & "EquipoMovil.png", True
			imagenToWord "Equipo Móvil", RutaEvidencias() & "EquipoMovil.png"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").JavaCheckBox("Seleccionar").Set "ON"
			wait 1
		Else 
			tiempo = 0
				Do
				tiempo = tiempo + 1
					If tiempo >= 180 Then
				  		DataTable("s_Resultado", dtLocalSheet) = "Fallido"
				   		DataTable("s_Detalle", dtLocalSheet) = "No se encuentra el equipo movil de la busqueda"
				   		Reporter.ReportEvent micFail, DataTable("s_Resultado", dtLocalSheet) , DataTable("s_Detalle", dtLocalSheet)
				   		ExitActionIteration
				   else
				   		Reporter.ReportEvent micPass, "Exito", "Se habilito el boton Agregar Dispositivos"
					End If
				wait 1
				Loop While Not JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").JavaCheckBox("Seleccionar_2").Exist
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").CaptureBitmap RutaEvidencias() & "EquipoMovil.png", True
			imagenToWord "Equipo Móvil", RutaEvidencias() & "EquipoMovil.png"
			wait 1
			JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").JavaCheckBox("Seleccionar_2").Set "ON"
			wait 1
		End If
 @@ hightlight id_;_3031813_;_script infofile_;_ZIP::ssf53.xml_;_
		JavaWindow("Ejecutivo de interacción").JavaDialog("Seleccione Dispositivos").JavaButton("Asignar Seleccionado").Click @@ hightlight id_;_14967096_;_script infofile_;_ZIP::ssf54.xml_;_
	
			While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaCheckBox("Alquiler").Exist)=False
				wait 1
			Wend
End Sub
Sub LineaAlquiler()
		If str_alquiler="Alquiler" Then
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaCheckBox("Alquiler").Set "ON"
			wait 2
			JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaList("duración de compromiso").Select str_dias
			wait 2
			JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ConfiguracionAlquiler.png", True
			imagenToWord "Configuración Alquiler", RutaEvidencias() & "ConfiguracionAlquiler.png"
			wait 1
		End If
End Sub
Sub EnviarOrdenAcuerdo()
		wait 2
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Crear Acuerdo Comercial").JavaButton("Ir al resumen del pedido").Click
		While(JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de Acuerdo Comercial(T").JavaButton("Enviar orden de acuerdo").Exist)=False
			wait 1
		Wend
		
		tiempo=0
			Do 
				tiempo=tiempo+1
				varresumen=JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de Acuerdo Comercial(T").JavaButton("Enviar orden de acuerdo").GetROProperty("enabled")
				wait 2	
		Loop While Not (varresumen="1")
		wait 2
		
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ResumenAcuerdoComercial.png", True
		imagenToWord "Resumen de Acuerdo Comercial", RutaEvidencias() & "ResumenAcuerdoComercial.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaInternalFrame("Resumen de Acuerdo Comercial(T").JavaButton("Enviar orden de acuerdo").Click
		
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Confirme Aceptar acuerdo").JavaButton("Confirme Aceptar acuerdo").Exist)=False
			wait 1
		Wend
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "ConfirmacionAcuerdoComercial.png", True
		imagenToWord "Confirmación Acuerdo Comercial", RutaEvidencias() & "ConfirmacionAcuerdoComercial.png"		
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Confirme Aceptar acuerdo").JavaButton("Confirme Aceptar acuerdo").Click
		While(JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").Exist)=False
			wait 1
		Wend	
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaStaticText("Pedido 1203554A se ha").Output CheckPoint("Pedido 1203554A se ha enviado correctamente.(st)") @@ hightlight id_;_25346215_;_script infofile_;_ZIP::ssf56.xml_;_
		DataTable("s_Nro_Pedido", dtLocalSheet) = Left(DataTable("s_Nro_Pedido",dtLocalSheet), 15) 
		DataTable("s_Nro_Pedido", dtLocalSheet) = Right(DataTable("s_Nro_Pedido",dtLocalSheet), 8) 
		JavaWindow("Ejecutivo de interacción").CaptureBitmap RutaEvidencias() & "NumeroPedido.png", True
		imagenToWord "Numero de Pedido", RutaEvidencias() & "NumeroPedido.png"
		wait 1
		JavaWindow("Ejecutivo de interacción").JavaDialog("Mensaje").JavaButton("OK").Click
End Sub


