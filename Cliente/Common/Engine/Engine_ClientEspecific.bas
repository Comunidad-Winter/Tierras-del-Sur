Attribute VB_Name = "Engine_ClientEspecific"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Consola As vWControlChat

Public Sub Iniciar_Motor_Grafico()
    LogDebug "Iniciando motor gráfico {"
        
    ' Controlador de Caminata
    Timer_Caminar.Time
    
    LogDebug "Iniciando inventario... "
    

    
    LogDebug "Iniciando formulario prinicipal... "
        
    ' Inicializo el formulario
    Engine_Init frmMain.Renderer.hWnd, TILES_WIDTH, TILES_HEIGHT
    
    LogDebug "Iniciando armas prinicipal... "
    
    ' TODO. Esto esta ok?
    Engine_Extend.Init_weapons
    
    LogDebug "Iniciando sangre... "
    
    ' Sangre
    Engine_Sangre.Initialize_Sangre
    
    LogDebug "Iniciando sonido ambiental... "
    
    Sonido_Ambiental_Iniciar 10
    
    LogDebug "Iniciando consola... "
    
    ' Configuro la Consola
    Set Consola = New vWControlChat
    Set Consola_Clan = New vWControlChat
        
    Consola.CantidadDialogos = 8
    Consola_Clan.CantidadDialogos = 5
    
    LogDebug "Render iniciado."
    
    prgRun = True
    
    ' Iniciado
    LogDebug "}"
End Sub

Public Sub Terminar_Motor_Grafico()

    Dim LiberarMemoria As String
    
    Call Sonido_Ambiental_ReIniciar
    
    Call Engine.Engine_Deinit
    
    Call DeInit_TextureDB
    
    Call Cli_CacheMapas.finalizar
    
    Call ResetResolution
    
    LiberarMemoria = Space(80000000) 'cerca de 70 mb liberados
    LiberarMemoria = vbNullString
    
    LogDebug "[FIN!]"

End Sub

Public Sub RenderInterface()

DibujarInterface

End Sub
Public Sub Render()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
'On Error GoTo jojo:
'On Error GoTo 0
Static ultimo As Long

    'particletimer = timerElapsedTime * 0.01
    'FIXME Protocol.aim_pj = 105
    'FIXME If Protocol.aim_pj <> 105 Then End
    If Not Device_Test_Cooperative_Level Then Exit Sub
    
    'PRECALC

    
    If SuperWater Then
        kWATER = (kWATER + timerTicksPerFrame * 16) Mod 360
        map_render_kwateR
    End If
    
    ultimo = timeGetTime
    
    timerElapsedTime = FRAME_TIMER.Time
    
    If timerElapsedTime > 200 Then timerElapsedTime = 200
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    
    cron_tiempo
    
    AnimarTilesets
    
    map_render_light
    
    Entidades_Actualizar

    Engine_Calc_Screen_Moviment
    
    Engine_Gfx_BeginScene
    Engine_Gfx_Clear
    
    'MAPA:
    

    Map_Render

    
    If IScombate Then Call text_render_graphic("Modo Combate", MainViewHeight - 100, 5, mzRed)
    If IsEnvenenado Then Call text_render_graphic("(Envenenado)", MainViewHeight - 100, 20, mzGreen)
    If Istrabajando Then Call text_render_graphic("Trabajando", MainViewHeight - 100, 5, mzWhite)

    If TiempoDrogaInicio > ultimo Then
    
        If MostrarTiempoDrogas Then
            Call Grh_Render(GRH_TIEMPO_DOPA, MainViewWidth - 96, MainViewHeight - 64, -1)

            If TiempoDrogaInicio - ultimo > 1000 Then
                Call text_render_graphic(Str(Fix((TiempoDrogaInicio - ultimo) / 1000)), MainViewWidth - 84, MainViewHeight - 26, COLOR_DOPATIMEOUT)
            Else
                Call text_render_graphic(FormatNumber((TiempoDrogaInicio - ultimo) / 1000, 2, vbFalse), MainViewWidth - 84, MainViewHeight - 26, COLOR_DOPATIMEOUTFINAL)
            End If
        Else
            ' Muestro cuanto tiene en cada stat
            Call Grh_Render(GRH_FUERZA, MainViewWidth - 130, MainViewHeight - 36, -1)  ' Fuerza
            Call Grh_Render(GRH_AGILIDAD, MainViewWidth - 70, MainViewHeight - 36, -1)  ' Agilidad
        
            Call text_render_graphic(Str(UserStats(SlotStats).UserAgilidad), MainViewWidth - 47, MainViewHeight - 28, COLOR_AGILIDAD)
            Call text_render_graphic(Str(UserStats(SlotStats).UserFuerza), MainViewWidth - 108, MainViewHeight - 29, COLOR_FUERZA)
        End If
    End If
    
    If TiempoAnguloNPC > ultimo Then
        Grh_Render_Rotated GRH_FLECHA_CRIATURA, PosAngleFlechaX, PosAngleFlechaY, base_light, AnguloProximoNPC
    End If
    
    If Cartel Then Call DibujarCartel

    Dialogos.Render
    
    If Consola_Clan.CantidadDialogos > 0 And Consola_Clan.Activo = True Then
        Consola_Clan.Draw 5, MainViewHeight - 80
    End If
    
    Consola.Draw 5, 0
    
    GUI_Render
    
    If MousePress = 1 Then
        Call Grh_Render(UserInventory(itemElegido).GrhIndex, MousePressX - 16, MousePressY - 16, mzWhite)
        
        If MousePressPosX > 0 Then
            Dim pixelPressX As Integer
            Dim pixelPressY As Integer

            pixelPressX = (MousePressPosX + minXOffset) * 32 + offset_map.X
            pixelPressY = (MousePressPosY + minYOffset) * 32 + offset_map.Y
            
            Grh_Render_Blocked &H33FFFFFF, pixelPressX, pixelPressY, MousePressPosX, MousePressPosY
        End If
        
    End If

    Engine_Gfx_EndScene
    
    If frmMain.picInv.Visible Then DrawInv
    
    If Comerciando Then
        If frmComerciar.Visible Then
            Call frmComerciar.refrescar
        End If
    End If
    
    If Bovedeando Then
        If frmBancoObj.Visible Then
            Call frmBancoObj.refrescar
        End If
    End If
        
Exit Sub
jojo:
LogError "Render: " & D3DX.GetErrorString(Err.Number) & " Desc: " & _
    Err.Description & " #: " & Err.Number
LogDebug "Render: " & D3DX.GetErrorString(Err.Number) & " Desc: " & _
    Err.Description & " #: " & Err.Number
    
End Sub

