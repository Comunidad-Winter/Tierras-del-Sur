Attribute VB_Name = "CLI_General"
Option Explicit

'IP a la cual se conecta el server

Public PingPerformanceTimer As New clsPerformanceTimer
Public profileClicks As Boolean

Private ultimo As Byte

Sub SetConnected()

  Dim i As Integer
  '*****************************************************************
  'Sets the client to "Connect" mode
  '*****************************************************************
  'Set Connected

    Connected = True
    'Call SaveGameini
    'Unload the connect form
    Unload frmConnect
    
    frmMain.lblUserName.Caption = UserName
    
    If Not isTengoClan() Then
        frmMain.lblClan.Caption = ""
    Else
        frmMain.lblClan.Caption = CharList(UserCharIndex).Clan
    End If
    
    Dim Color As Long
    Color = getHexaColorByPrivForInterface(CharList(UserCharIndex))
    
    frmMain.lblClan.ForeColor = Color
    frmMain.lblUserName.ForeColor = Color

    Call CambiarCursor(frmMain)

    For i = 0 To 20
        If Listaintegrantes(i) <> "" Then
            Listaintegrantes(i) = ""
        End If
    Next

    frmMain.Visible = True
    frmMain.SoundFX.Enabled = EfectosSonidoActivados
    
    frmMain.Pasarsegundo.Enabled = True
    frmMain.pasarMinuto.Enabled = True
    frmMain.comienzoMinutoCheat = timeGetTime
    
    
    frmMain.SendTxt.Visible = False
    frmMain.lblIndicadorEscritura.Visible = False
    ultimo = Second(Time) ' anticheat

    Istrabajando = False
    UserStats(SlotStats).UserCentinela = False
    frmMain.IconoDyd = ""
End Sub

Public Sub cargarRecursos_Interface()
    ' Cargo estructura de la interface
    Call pakGUI.Cargar(RecursosPath & "Interface.TDS")
    Call LogDebug("Cargada Interface.")
End Sub

Public Sub cargarRecursos()

    #If testeo = 0 Then
       ' If MD5String(MD5File(RecursosPath & "Graficos.TDS") & "SALTO222") <> "b7317dee4e33e7fd0965a78beedf4b24" Then
       '     Call MsgBox("Cliente corrupto. Bájelo de nuevo desde www.tierrasdelsur.cc.", vbApplicationModal + vbCritical + vbOKOnly, "Error al ejecutar")
       '     End
       ' End If

       ' If MD5String(MD5File(RecursosPath & "Mapas.TDS") & "SALTO222") <> "8def45f26c818f667067607a3e9f691b" Then
       '     Call MsgBox("Cliente corrupto. Bájelo de nuevo desde www.tierrasdelsur.cc.", vbApplicationModal + vbCritical + vbOKOnly, "Error al ejecutar")
       '     End
       ' End If

        'If MD5String(MD5File(RecursosPath & "\Graficos.ind") & "SALTO222") <> "4a7cecc731b69dceb8fd7713ee174231" Then
        '    Call MsgBox("Cliente corrupto. Bájelo de nuevo desde www.tierrasdelsur.cc.", vbApplicationModal + vbCritical + vbOKOnly, "Error al ejecutar")
        '    End
        'End If

        'If MD5String(MD5File(RecursosPath & "\Cabezas.ind") & "SALTO222") <> "78db2f685c222fca888a3514351222f2" Then
        '    Call MsgBox("Cliente corrupto. Bájelo de nuevo desde www.tierrasdelsur.cc.", vbApplicationModal + vbCritical + vbOKOnly, "Error al ejecutar")
        '    End
        'End If
    #End If

    ' Cargo la estructura de los mapas
    'Call CargarPakMapas(RecursosPath & "Mapas.TDS")
    'Call LogDebug("Cargados Mapas.")
    
    ' Cargos los efectos de pisada
    Call CLI_EfectosPisadas.Cargar_EfectosPisadas
    Call LogDebug("Cargados Efectos de Pisada.")
    
    ' Cargo Gráficos
    Call CLI_Carga_Inits.Cargar_Graficos
    Call LogDebug("Cargados Graficos.")
        
    ' Cargo Armas
    Call CLI_Carga_Inits.Cargar_Armas
    Call LogDebug("Cargadas Armas.")
    
    ' Cargo Escudos
    Call CLI_Carga_Inits.Cargar_Escudos
    Call LogDebug("Cargados Escudos.")
    
    ' Cargo Cuerpos
    Call CLI_Carga_Inits.Cargar_Cuerpos
    Call LogDebug("Cargados Cuerpos.")
    
    ' Cargo Cabezas
    Call CLI_Carga_Inits.Cargar_Cabezas
    Call LogDebug("Cargadas Cabezas.")
    
    ' Cargo Cascos
    Call CLI_Carga_Inits.Cargar_Cascos
    Call LogDebug("Cargados Cascos.")
    
    ' Cargo Efectos
    Call CLI_Carga_Inits.Cargar_Efectos
    Call LogDebug("Cargados Efectos.")
    
    ' Cargo Pisos
    Call CLI_Carga_Inits.Cargar_Pisos
    Call LogDebug("Cargados Pisos.")

End Sub
Public Sub iniciarEstructuras()

    ' Tipo de separador que uitliza el Sistema
    Call DecimalSeparator
    
    ' Efectos
    ReDim FXList(0)
    
    ' Semilla de números random
    Randomize timeGetTime _
                            Xor 4 _
                            Xor 8 _
                            Xor 15 _
                            Xor 16 _
                            Xor 23 _
                            Xor 42 Xor 48151623 'LOST
                            
    ' Manejador de Bits
    Call BS_Init_Table
    
    ' Cache de mapas
    Call Cli_CacheMapas.iniciar
    Call LogDebug("  Manejador de mapas dinámicos iniciado.")
End Sub


Private Function calcularCOnwi(text As String) As Integer

Dim loopLetra As Integer
Dim ba() As Byte

ba() = StrConv(text, vbFromUnicode)

For loopLetra = LBound(ba) To UBound(ba)
    calcularCOnwi = calcularCOnwi + Font_Default.HeaderInfo.CharWidth(ba(loopLetra)) '* scalea
Next

End Function

Public Function cortarTexto(ByRef Font As Byte, ByVal texto As String, longitud As Integer, collection As collection)
    Dim Str As String
    Dim Tmp As Integer
    Dim posEnter As Integer
    Dim Pos As Integer
    Dim posEspacio As Integer
    Dim acumulado As Integer
    Dim i As Integer
    Dim Char As Byte
        
    If LenB(texto) = 0 Then Exit Function
    
    For i = 1 To Len(texto)
        Char = Asc(mid$(texto, i, 1))
        
        If Char = vbKeyReturn Then
            posEnter = i
            Exit For
        ElseIf acumulado >= longitud Then
            Exit For
        ElseIf Char = vbKeySpace Then
            posEspacio = i
        End If
        
        acumulado = acumulado + Fonts(Font).Font.HeaderInfo.CharWidth(Char)
    Next i
    
    If acumulado < longitud Then
        Call collection.Add(texto)
        Exit Function
    End If
        
    If (posEnter > 0 And posEnter < posEspacio) Then
        Call collection.Add(left$(texto, posEnter - 1) & Chr(13))
        Call cortarTexto(Font, mid$(texto, posEnter + 1), longitud, collection)
    ElseIf posEspacio > 0 And (posEspacio < posEnter Or posEnter >= 0) Then
        Call collection.Add(left$(texto, posEspacio - 1) & Chr(13))
        Call cortarTexto(Font, mid$(texto, posEspacio + 1), longitud, collection)
    ElseIf longitud < Len(texto) Then
        Call collection.Add(left$(texto, longitud) & Chr(13))
        Call cortarTexto(Font, mid$(texto, longitud + 1), longitud, collection)
    Else
        Call collection.Add(texto)
    End If

   

End Function

Sub Main()

    ' Establecmos directorio de trbajao
    ChDir app.Path
    ChDrive app.Path

    ' Iniciamos los Paths donde se encuentra las cosas
    Call Configuracion_Usuario.iniciarPaths
    Call LogDebug("  Paths Iniciados.")
    
    ' Cargamos la configuracion del usuario
    Call Configuracion_Usuario.cargarConfiguracionUsuario
    
    ' Establecemos la resoluion del juego
    Call Engine_Resolution.setResolucionJuego(ResolucionJuego)
    
    ' Cambiamos la resolucion de la pc
    Call SetResolutionPantalla(forzarFullScreen, Engine_Resolution.pixelesAncho, Engine_Resolution.pixelesAlto)
        
    Call CLI_General.cargarRecursos_Interface
    
    ' TODO
    'base_light_techo = D3DColorXRGB(150, 150, 150) And &HFFFFFF
    
    frmPres.Show
    
    frmPres.checkJuego
    
    ' Inicializo las Encriptacion
    Call CryptoInitInicial
    
    UserMac = GetMACAddress()
    
    Call CryptoInit
    Call LogDebug("  Encriptación configurada.")
                  
    ' Cargamos el lenguaje por defecto
    If Not CargarLenguaje(Configuracion_Usuario.lenguaje) Then
        MsgBox "No se pudo cargar el lenguaje. Por favor, re instale el juego.", vbCritical
        End
    End If
    
    ' Iniciamos estructura de juego
    Call CLI_General.iniciarEstructuras
    
    ' Muestro pantalla de Presentacion
    
    DoEvents
    
    ' Cargamos los recursos del juego
    Call CLI_General.cargarRecursos
    
    ' Cosas Hacordeadas
    Call Iniciar_Constantes_De_Juego
        
    '
    Call Iniciar_Motor_Grafico
     

    LogDebug "  Instancia de Audio iniciada."
    
    ' Cargamos logins servers
    Call modLogin.iniciarLogins
    
    ' Muestro pantalla de conectar
    
    Do While Not frmPres.isReady
        Sleep 100
        DoEvents
    Loop
    
    Load frmConnect
    
    frmConnect.Show
    
    Call frmPres.SetFocus
     
    DoEvents
     
    ' Iniciamos el registro de eventos
    Call LogDebug_Iniciar
    
    ' Se corre el Engine
    Call comenzarConectar
    Call Engine_Start
    
    If Grabando Then FinalizarGrabacion
          
    Call Terminar_Motor_Grafico
    
    ' Cierro la aplicacion
    Call UnloadAllForms
    End
     
     Exit Sub
     

ManejadorErrores:
        LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
        End

End Sub


Public Function GenerarFotoDenuncia() As String
    ' TODO. Falta activar
    GenerarFotoDenuncia = Dialogos.GDialogos
End Function

Public Function checksum(cadena As String, key As Byte) As String
Dim Salto As String

Salto = "01luoq"

checksum = mid(MD5String(cadena & key & Salto), 3, 11)
Debug.Print checksum
End Function

