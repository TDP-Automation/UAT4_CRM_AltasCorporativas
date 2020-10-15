Option Explicit

'Call IniExe()
Dim str_usuario, str_password, str_idioma, UAT8, UAT5, UAT4, UAT6, UAT10, UAT13, PROD, intStartTime, intStopTim, UAT3
intStartTime = Timer
str_usuario = DataTable("e_Usuario", dtLocalsheet)
str_password = DataTable("e_Password", dtLocalsheet)
str_idioma = "español (Perú)"

SystemUtil.CloseProcessByName "javaw.exe"
SystemUtil.CloseProcessByName "jp2launcher.exe"

Call SeleccionarAmbiente()
Call ValidacionesLogin()
Call IngresoDatosLogin()
Call EsperaCargaLogin()

Sub SeleccionarAmbiente()

	Select Case DataTable("e_Ambiente", dtLocalsheet)
	
				Case "UAT8"
				UAT8="http://tfpwltst04:30828/Crm/CRM/Crm.jnlp"
				SystemUtil.Run "iexplore.exe", UAT8
				wait 10
				Case "UAT4"
				UAT4="http://tfpwltst02:30428/Crm/CRM/Crm.jnlp"
				SystemUtil.Run "iexplore.exe", UAT4
				wait 10
				Case "UAT5"	
				UAT5="http://tfpwltst03:30528/Crm/CRM/Crm.jnlp"
				SystemUtil.Run "iexplore.exe", UAT5
				wait 10
				Case "UAT6"
				UAT6="http://tfpwltst03:30628/Crm/CRM/Crm.jnlp"
				SystemUtil.Run "iexplore.exe", UAT6
				wait 10		
				Case "UAT10"
				UAT10="http://tfpwltst05:31028/Crm/CRM/Crm.jnlp"
				SystemUtil.Run "iexplore.exe", UAT10		
				wait 10		
				Case "UAT13"
				UAT13="http://tfpwltst07:31328/Crm/CRM/Crm.jnlp"
				SystemUtil.Run "iexplore.exe", UAT13		
				wait 10					
				Case "PROD"
				PROD="http://10.4.55.14/Crm/CRM/Crm.jnlp"
				SystemUtil.Run "iexplore.exe", PROD
				wait 10		
				Case "UAT3"
				UAT3="http://tfpwltst02:30328/Crm/CRM/Crm.jnlp"
				SystemUtil.Run "iexplore.exe", UAT3
				wait 10	
	End Select

End Sub
Sub ValidacionesLogin()

	While (JavaWindow("Iniciar sesión").JavaList("Iniciar sesión en:").Exist) = False
		wait 1
		Dim a
		a=a+1
				'If ((Browser("No se puede acceder a").Page("No se puede acceder a").Exist) or (JavaDialog("Error de Aplicación").JavaButton("Detalles").Exist)) Then
				If (JavaDialog("Error de Aplicación").JavaButton("Detalles").Exist) Then

					SystemUtil.CloseProcessByName "javaw.exe"
					SystemUtil.CloseProcessByName "jp2launcher.exe"
					SystemUtil.CloseProcessByName "iexplore.exe"
					
					Select Case DataTable("e_Ambiente", dtLocalsheet)
					Case "UAT8", "uat8"
					UAT8="http://tfpwltst04:30828/Crm/CRM/Crm.jnlp"
					SystemUtil.Run "iexplore.exe", UAT8
					Case "UAT4", "uat4"
					UAT4="http://tfpwltst02:30428/Crm/CRM/Crm.jnlp"
					SystemUtil.Run "iexplore.exe", UAT4
					Case "UAT5", "uat5"
					UAT5="http://tfpwltst03:30528/Crm/CRM/Crm.jnlp"
					SystemUtil.Run "iexplore.exe", UAT5
					Case "UAT6", "uat6"
					UAT6="http://tfpwltst03:30628/Crm/CRM/Crm.jnlp"
					SystemUtil.Run "iexplore.exe", UAT6
					Case "UAT10", "uat10"
					UAT10="http://tfpwltst05:31028/Crm/CRM/Crm.jnlp"
					SystemUtil.Run "iexplore.exe", UAT10		
					Case "UAT13", "uat13"
					UAT13="http://tfpwltst07:31328/Crm/CRM/Crm.jnlp"
					SystemUtil.Run "iexplore.exe", UAT13		
					Case "PROD", "prod"
					UAT6="http://10.4.55.14/Crm/CRM/Crm.jnlp"
					SystemUtil.Run "iexplore.exe", PROD
					End Select
				
				End If
	
	If JavaDialog("Advertencia de Seguridad").JavaCheckBox("Acepto los riesgos y deseo").Exist Then
		JavaDialog("Advertencia de Seguridad").JavaCheckBox("Acepto los riesgos y deseo").Set "ON"
		JavaDialog("Advertencia de Seguridad").JavaButton("Ejecutar").Click
	End If

	If JavaDialog("Advertencia - Versión").JavaButton("Ejecutar con Versión Más").Exist Then
		JavaDialog("Advertencia - Versión").JavaButton("Ejecutar con Versión Más").Click
	End If
				
	If JavaDialog("Información de Seguridad").JavaObject("JLayeredPane").Exist Then
		If JavaDialog("Información de Seguridad").JavaCheckBox("No volver a mostrar esto").Exist Then
			JavaDialog("Información de Seguridad").JavaCheckBox("No volver a mostrar esto").Set "ON"
		End If
		JavaDialog("Información de Seguridad").JavaButton("Ejecutar").Click
	End If
	
	If JavaDialog("Advertencia de Seguridad").Exist Then
		JavaDialog("Advertencia de Seguridad").JavaCheckBox("Acepto los riesgos y deseo").Set "ON"
		JavaDialog("Advertencia de Seguridad").JavaButton("Ejecutar").Click
	End If
	
	If Dialog("Java Update Necesario").Exist Then
		Dialog("Java Update Necesario").WinObject("Java Update Necesario").WinCheckBox("No volver a preguntar").Set "ON"
		Dialog("Java Update Necesario").WinButton("Más tarde").Click
	End If

	If (a=180) Then
		Call NextStep()
	End If
	
Wend
	
End Sub
Sub IngresoDatosLogin()

	JavaWindow("Iniciar sesión").JavaList("Iniciar sesión en:").Select "Ejecutivo de interacción del cliente de Amdocs" @@ hightlight id_;_6155557_;_script infofile_;_ZIP::ssf5.xml_;_
	JavaWindow("Iniciar sesión").JavaEdit("Nombre Login").Set str_usuario
	JavaWindow("Iniciar sesión").JavaEdit("Contraseńa").Set str_password
	JavaWindow("Iniciar sesión").JavaList("Idioma:").Select str_idioma
	JavaWindow("Iniciar sesión").JavaCheckBox("Usar fomato de idioma").Set "ON"
	JavaWindow("Iniciar sesión").JavaButton("Iniciar sesión").Click
	JavaWindow("Iniciar sesión").CaptureBitmap RutaEvidencias() & "Login.png", True
	imagenToWord "Login CRM", RutaEvidencias() & "Login.png" 
End Sub
Sub EsperaCargaLogin()
	
	While (JavaWindow("Ejecutivo de interacción").JavaStaticText("JLabel(st)").Exist) = False
		wait 1	
		Dim b
		b=b+1
			If b=120 Then
				Call NextStep()
			End If		
	Wend
	
	JavaWindow("Ejecutivo de interacción").Maximize
	
End Sub

