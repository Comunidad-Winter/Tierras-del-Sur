Attribute VB_Name = "ME_Render"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim T(1 To 10) As clsPerformanceTimer

Public Const ClientWindowWidth = 672
Public Const ClientWindowHeight = 672

Public DRAWCLIENTAREA As Byte
Public DRAWGRILLA As Byte

Public WalkMode As Boolean

Dim TiempoPresent As Single
Dim TiempoMapa As Single
Dim TiempoLuces As Single
Public TiempoLucesLightmaps As Single
Public TiempoAguatierra As Single

#If medir Then
    Public mostrarTiempos As Boolean
#End If

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long


Public Function GetElapsedTimeME() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTimeME = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub CrearCharWalkMode()
    If UserCharIndex = 0 Then
        UserCharIndex = NextOpenChar
    End If
    
    If UserPos.x = 0 Then
        UserPos.x = 10
        UserPos.y = 10
    End If
        
    Call MakeChar(UserCharIndex, 1, 1, SOUTH, UserPos.x, UserPos.y, 0, 0, 0)
    Call DeactivateChar(CharList(UserCharIndex))

    CharList(UserCharIndex).nombre = CDM.cerebro.Usuario.nombre
    Engine_Extend.char_act_color UserCharIndex
End Sub



Public Sub ToggleWalkMode()
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************
On Error GoTo fin:
If WalkMode = False Then
    WalkMode = True
Else
    frmMain.ver_char.checked = False
    WalkMode = False
End If

If WalkMode = False Then
    'Erase character
    DeactivateChar CharList(UserCharIndex)
    'CharList(UserCharIndex).Velocidad.x = 40
    'CharList(UserCharIndex).Velocidad.y = 40
Else
    'MakeCharacter
    If PuedoCaminar(UserPos.x, UserPos.y, E_Heading.NONE, False, False) Then
        Call CrearCharWalkMode
        Call ActivateChar(CharList(UserCharIndex))
        'CharList(UserCharIndex).Velocidad
        
        frmMain.ver_char.checked = True
    Else
        MsgBox "ERROR: Ubicacion ilegal."
        WalkMode = False
    End If
End If
fin:
End Sub

Private Sub chequearActualizacionPendiente()

    Dim posibleExePendiente As String
    Dim comando As String
    
    posibleExePendiente = app.Path & "\EditorTDS.exe_"
    comando = Chr$(34) & app.Path & "\Updater.exe" & Chr$(34) & " 5 " & Chr$(34) & app.Path & "\EditorTDS.exe" & Chr$(34)
     
    If FileExist(posibleExePendiente, vbNormal) Then
        Call Shell(comando, vbNormalFocus) ' Reset
        End
    End If
            
End Sub

Sub Main()

ChDir app.Path
ChDrive app.Path

' Chequeamos que no tengamos una actualizaicon pendiente
Call chequearActualizacionPendiente

' Inicimaos modulo de enriptacion
Call CryptoInit

' Iniciamos el Cerebro de mono
Call CDM_Iniciar(frmCDM.Inet1, frmCDM.Timer1, "TDSEditor/" & VERSION_EDITOR & "/" & app.Major & "." & app.Minor & "." & app.Revision & "/" & GetIdentificacionPC() & "/" & "2a6dfb1e9514102315c66f2470f57197")

' Mostramos la pantalla de login
Call frmCDMLogin.Show(vbModal)

'Si no se conecto, cierro
If Not CDM.cerebro.estado = eEstadoCDM.conectado Then
    If Not IsIDE Then
        MsgBox "Para poder ingresar al Editor del Mundo es necesario que inicies sesión. ", vbExclamation, "Tierras del Sur - Editor del Mundo"
        End
    End If
End If

' Cargo la configuracion del usuario
Call ME_Configuracion_Usuario.cargarConfiguracionUsuario(app.Path & "/Me.ini")

' Aplicamos la Configuracion
Call ME_General.AplicarConfiguracion

'Creo los directorios necesarios
Call ME_Configuracion_Usuario.crearDirectorios

' Validamos que haya aceptado el contrato
'If Not modContrato.contratosAceptados() Then End

'Cargo la configuracion de la pantalla
Call modPantalla.Pantalla_Iniciar

'Versionados de los elementos de los archivos de recursos
Call versionador.iniciar_versionador

' Iniciamos el modulo que me da soporte para comunicarme http
If Not modHttp.iniciar Then
    Call MsgBox("Es necesario reiniciar el equipo para aplicar cambios en la configuración que permita que el Editor funcione correctamente. Estos cambios no alteraran el funcionamiento del equipo o de otras aplicaciones.", vbExclamation)
    End
End If

'Cargo climas
Call ME_Climas.cargarClimasDisponibles

'Cargo las distintas zonas que se pueden editar con el mundo
Call ME_Mundo.cargarZonasPosibles

' Cargo Modulo Auxiliar para el manejo de bytes
Call BS_Init_Table

' Cargamos el formulario
load frmMain

Call modPantalla.Pantalla_AcomodarElementos

' Mostramos el formulario
frmMain.Show

' Lo dejamos msotrarse
DoEvents

LogDebug "Iniciando motor grafico"

' Inicio el motor gráfico
Call Iniciar_Motor_Grafico

' Muestro el Editor
miniMapInit
miniMap_Redraw

Sonido_Ambiental_Iniciar 400

'Cargo las propiedades de los mapas de los mapas
Call ME_Mapas.cargarInformacionMapas

LogDebug "Cargando indices:     Tilesets"
Call CargarTilesetsIni(DBPath & "tilesets.ini")      'Tilesets

LogDebug "                      Triggers"            'Triggers
Call ME_Tools_Triggers.LoadTriggersRaw(DBPath & "triggers.ini")
Call ME_Tools_Triggers.cargarTriggersALista(frmMain.lstTriggers)

LogDebug "                      NPC"
Call ME_obj_npc.cargarInformacionNPCs                 ' Npcs
Call ME_obj_npc.cargarListaNPC

LogDebug "                      Efectos de Pisadas"
Call Me_indexar_EfectosPisadas.cargarInformacionEfectosPisadas

LogDebug "                      Hechizos"
Call Me_Hechizos.cargarInformacionHechizos                ' Hechizoz

LogDebug "                      Objetos"

'De los objetos
Call ME_obj_npc.cargarInformacionObjetos            'Objetos
Call ME_obj_npc.cargarListaObjetos

LogDebug "                      Presets"
Call CargarPresets                                  ' Presets

LogDebug "                      Graficos comunes"
Call CargarListaGraficosComunes                     'Graficos

LogDebug "                      Mapas"
Call CargarPakMapas(app.Path & "\Datos\Mapas\MapasME.TDS", ClientPath & "Graficos\Mapas.TDS")

LogDebug "                      Aspectos"
Call Me_indexar_Pixels.cargarDesdeIni

LogDebug "  Indices OK!         Limpiando Mapa e iniciando editor"

'Inicio las tools
Call ME_Tools.iniciar

' Inicio el portapeles de areas del mapa
Call Me_Tools_Seleccion.iniciarPortapapeles

MouseTileX = 50
MouseTileY = 50

init_map_editor

Call NuevoMapa

THIS_MAPA.editado = True
If FileExist(app.Path & "\Datos\tmpmap.cache") Then
    ME_FIFO.Cargar_Mapa_ME app.Path & "\Datos\tmpmap.cache"
End If

If THIS_MAPA.editado = True Then
    THIS_MAPA.editado = False
Else
    GUI_Alert "Se cargó una copia de seguridad del mapa [" & Chr$(255) & THIS_MAPA.numero & Chr$(255) & "] al estado del " & FileDateTime(app.Path & "\Datos\tmpmap.cache") & "." & vbNewLine & "Por favor no guarde esta versión del mapa a menos que esté seguro de que los cambios que tiene sean correctos.", "Mapa cargado"
    THIS_MAPA.editado = True
End If

LogDebug "Cargando acciones!"
'***********************
' Cargo las Acciones disponibles
Call ME_modAccionEditor.cargarListaAccionTile
Call ME_modAccionEditor.refrescarListaDisponibles(frmMain.listTipoAccionesDisponibles)

' Lista de presets
Call ME_presets.cargarListaPresets

' Lista de entidades
Call Me_Tools_Entidades.cargarListaEntidades

LogDebug "################# INICIANDO ! #################"

Start
'********************************************************'
' Se se cirra bien eliminamos el archivo temporal
If FileExist(app.Path & "\Datos\tmpmap.cache") Then
    Kill app.Path & "\Datos\tmpmap.cache"
End If

End ' FIN

End Sub


Public Sub Iniciar_Motor_Grafico()
    
    LogDebug "Iniciando motor gráfico {"
    
    
    Randomize timeGetTime _
                            Xor 4 _
                            Xor 8 _
                            Xor 15 _
                            Xor 16 _
                            Xor 23 _
                            Xor 42 Xor 48151623 'LOST

    
    'modZLib.Bin_Load_Headers Clientpath & "Graficos\"
    LogDebug "  Archivo de recursos cargado."

    If CargarGraficosIni = False Then
        'Si no pudo cargar el archivo de graficos indexados, termina.
        End
    End If
    
    LogDebug "  Indices de graficos cargados."
    
    'Cargos cuerpos
    Call Me_indexar_Armas.cargarDesdeIni
    
    'Cargos escudos
    Call Me_indexar_Escudos.cargarDesdeIni
    
    LogDebug "  Animaciones cargadas."
    
    'Cuerpos
    Call Me_indexar_Cuerpos.cargarCuerpoEnIni
        
    'Cabezas
    Call Me_indexar_Cabezas.cargarDesdeIni
    
    'Cascos
    Call Me_indexar_Cascos.cargarDesdeIni
    
    LogDebug "  Personajes cargados.."
    
    'Efectos
    Call Me_indexar_Efectos.cargarDesdeIni
    
    'Sonidos
    Call Me_indexar_Sonidos.cargarDesdeIni
    
    'Entidades
    Call Me_indexar_Entidades.cargarDesdeIni
    
    LogDebug "  Efectos cargados."
    
    Call Engine_Init(frmMain.renderer.hwnd, CLng(modPantalla.TilesPantalla.x), CLng(modPantalla.TilesPantalla.y))
    
    Engine_Sangre.Initialize_Sangre
    
    LogDebug "}"
    
    prgRun = True

    re_render_inventario = True
    
    
End Sub

Public Sub Terminar_Motor_Grafico()

Dim LiberarMemoria As String

    prgRun = False
    
    Call Engine.Engine_Deinit
    
    Call DeInit_TextureDB
    
    Call ResetResolution
    
    LiberarMemoria = Space(80000000) 'cerca de 70 mb liberados
    LiberarMemoria = vbNullString
    
    LogDebug "[FIN!]"

End Sub



Public Sub Render()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
'On Error GoTo jojo:
    'particletimer = timerElapsedTime * 0.01
    RENDERCRC = GetTickCount * Rnd
    
#If medir = 1 Then
    If mostrarTiempos Then
        Dim buffer As String
        GetElapsedTimeME
    End If
#End If

    timerElapsedTime = FRAME_TIMER.Time
    
    If timerElapsedTime > 200 Then timerElapsedTime = 1
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed


#If medir = 1 Then
    If mostrarTiempos Then
        TiempoLuces = GetElapsedTimeME
        buffer = buffer & "Luces:       [" & Int(100 * (TiempoLuces / timerElapsedTime)) & "%]" & Round(TiempoLuces) & vbCrLf
    End If
#End If


    cron_tiempo
    'cron_fxs
    Entidades_Actualizar
    map_render_light
    
#If medir = 1 Then
    If mostrarTiempos Then
        TiempoLuces = GetElapsedTimeME
        buffer = buffer & "MapMov:      [" & Int(100 * (TiempoLuces / timerElapsedTime)) & "%]" & Round(TiempoLuces) & vbCrLf
    End If
#End If

    Engine_Calc_Screen_Moviment
    AnimarTilesets
        
    GetElapsedTimeME
    
    Engine_Gfx_BeginScene
    Engine_Gfx_Clear

TiempoPresent = GetElapsedTimeME

    part_totales = 0
        
    GetElapsedTimeME
    Map_Render
    
    ' ¿Necesito capturar la pantalla?
    If Me_Exportar.capturandoPantalla Then Call Me_Exportar.generarFraccionPantalla(D3DDevice, D3DX)
    
    'Engine_LightsTexture_Render
    Engine_LightsTexture_RenderBackbuffer
    
#If medir = 1 Then
    If mostrarTiempos Then
        TiempoLuces = GetElapsedTimeME - TiempoAguatierra
        buffer = buffer & "MapREN:      [" & Int(100 * ((TiempoLuces) / timerElapsedTime)) & "%]" & Round(TiempoLuces) & vbCrLf
        buffer = buffer & "LightREN:      [" & Int(100 * (TiempoLucesLightmaps / timerElapsedTime)) & "%]" & Round(TiempoLucesLightmaps) & vbCrLf
    
    End If
#End If

saltar_render:
    
    If areaSeleccionada.arriba = areaSeleccionada.abajo And areaSeleccionada.derecha = areaSeleccionada.izquierda Then
        Call text_render_graphic("FPS: " & Chr$(255) & format$(FPSTIMER(), "####.##") & " (" & Engine.FPS & ")" & Chr$(255) & vbCrLf & "Mouse " & Chr$(255) & "[" & areaSeleccionada.izquierda & ";" & areaSeleccionada.arriba & "]", 10, 480, &H7FFFFFFF)
    Else
        Call text_render_graphic("FPS: " & Chr$(255) & format$(FPSTIMER(), "####.##") & " (" & Engine.FPS & ")" & Chr$(255) & vbCrLf & "Mouse " & Chr$(255) & "[" & areaSeleccionada.izquierda & ";" & areaSeleccionada.arriba & "] a [" & areaSeleccionada.izquierda & ";" & areaSeleccionada.abajo & "]", 10, 480, &H7FFFFFFF)
    End If
   
    
    If TipoEditorParticulas Then
        Call text_render_graphic("Particulas totales: " & Chr$(255) & part_totales, 10, 42, &HFFEEEEEE)
    End If
    
    BatchSonidos
    
    
    GetElapsedTimeME
    DRAW_TOOL
    
    #If medir = 1 Then
        If mostrarTiempos Then
            TiempoLuces = GetElapsedTimeME
            buffer = buffer & "Tools:       [" & Int(100 * (TiempoLuces / timerElapsedTime)) & "%]" & Round(TiempoLuces) & vbCrLf
            Call text_render_graphic(buffer, 100, 100, &H60FFFFFF)
        End If
    #End If
GetElapsedTimeME

'miniMap_Render 10, 10

GUI_Render

'Engine.Grh_Render_Rotated 1178, 300, 300, -1, 90

Engine_Gfx_EndScene
TiempoPresent = TiempoPresent + GetElapsedTimeME
    
Exit Sub
jojo:
LogError "Render: " & D3DX.GetErrorString(Err.Number) & " Desc: " & Err.description & " #: " & Err.Number
End Sub

' Captura la parte del Render que se esta visualizando
Public Function capturarPantalla(Direct3DDevice As Direct3DDevice8, D3DX As D3DX8, ByVal ScreenHeight As Long, ByVal ScreenWidth As Long, Optional ByVal FilePath As String) As Boolean

    Dim RECT As RECT
    Dim PAL As PALETTEENTRY
    Dim desc As D3DSURFACE_DESC
    Dim srfBackBuffer As Direct3DSurface8
    
    PAL.blue = 255
    PAL.green = 255
    PAL.red = 255
    
   
    Set srfBackBuffer = D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    srfBackBuffer.GetDesc desc
    
    RECT.right = desc.Width
    RECT.bottom = desc.Height
    
    D3DX.SaveSurfaceToFile OPath & "Imagenes\" & FilePath, D3DXIFF_BMP, srfBackBuffer, PAL, RECT
    
End Function


Public Sub Start()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
DoEvents

Do While prgRun
    Dim cut_fps_ud As Long
    Rem Limitar FPS
    Static lFrameTimer As Long

    If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
        CheckKeys
        Render
        
        cut_fps_ud = GetTimer
        FramesPerSecCounter = FramesPerSecCounter + 1
        If (cut_fps_ud - lFrameTimer) >= 1000 Then
            Engine.FPS = FramesPerSecCounter
            FramesPerSecCounter = 0
            lFrameTimer = cut_fps_ud

        End If
    Else
        Sleep 10&
    End If


    'Audio.Music_GetLoop
'Call CargarPresets
    DoEvents
Loop
Engine.Engine_Deinit

End Sub

