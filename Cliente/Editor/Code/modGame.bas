Attribute VB_Name = "ME_Global"
Option Explicit
Public Const MSGDang As String = "CUIDADO! Este comando puede arruinar el mapa." & vbCrLf & "¿Estas seguro que querés continuar?"


Public MapPath As String
Public DBPath As String


Public Function EditWarning() As Boolean
If MsgBox(MSGDang, vbExclamation Or vbYesNo) = vbNo Then
    EditWarning = True
Else
    EditWarning = False
End If
End Function


Public Function HayCharHeading(ByVal X As Byte, ByVal Y As Byte, ByVal heading As E_Heading) As Boolean
    '*****************************************************************
    'Menduz
    '*****************************************************************

    Dim tTrigger As Long, nuevoX As Byte, nuevoY As Byte
    Dim bloqueoInteresado As eTriggers
    
    HayCharHeading = True
    
    'Me fijo la viabilidad del tile HACIA donde voy.
    'Sino me queda a la mitad, algunas cosas chequeandolas donde estoy y otras hacia donde voy
    'Y la idea es que mantenga una homogeneanidad logica para que sea mas facil de pensar
    
    nuevoX = X
    nuevoY = Y
        
    Select Case heading
        
        Case E_Heading.EAST
            
            nuevoX = X + 1
            bloqueoInteresado = eTriggers.BloqueoOeste
                
        Case E_Heading.WEST
            
            nuevoX = X - 1
            bloqueoInteresado = eTriggers.BloqueoEste
                
        Case E_Heading.NORTH
                
            nuevoY = Y - 1
            bloqueoInteresado = eTriggers.BloqueoSur
                
        Case E_Heading.SOUTH
                
            nuevoY = Y + 1
            bloqueoInteresado = eTriggers.BloqueoNorte
  
        Case E_Heading.NONE
            
            HayCharHeading = True
            
    End Select
    

    If CLI_PosicionesLegales.esPosicionJugable(nuevoX, nuevoY) Then
        HayCharHeading = Not CBool(mapdata(nuevoX, nuevoY).NpcIndex)
    End If
End Function

Public Sub MoveTo(ByVal direccion As E_Heading)
    Dim LegalOk As Boolean
    If WalkMode = True Then
        LegalOk = PuedoCaminar(UserPos.X, UserPos.Y, direccion, False, False) And HayCharHeading(UserPos.X, UserPos.Y, direccion)
        
        If UserCharIndex = 0 Then CrearCharWalkMode
        
        If LegalOk Then
            Char_Move_by_Head UserCharIndex, direccion
            Engine_MoveScreen direccion
        End If
    
    Else
        Select Case direccion
            Case E_Heading.NORTH
                LegalOk = InMapBounds(UserPos.X, UserPos.Y - 1)
            Case E_Heading.EAST
                LegalOk = InMapBounds(UserPos.X + 1, UserPos.Y)
            Case E_Heading.SOUTH
                LegalOk = InMapBounds(UserPos.X, UserPos.Y + 1)
            Case E_Heading.WEST
                LegalOk = InMapBounds(UserPos.X - 1, UserPos.Y)
        End Select
        Engine_MoveScreen direccion
    End If
    
    ' Update 3D sounds!
    'Call Audio.MoveListener(UserPos.X, UserPos.y)
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static lastMovement As Long
    
    ' -------------- Toools --------------------
   ' If GetAsyncKeyState(vbKeyControl) = 1 Then
        
       ' If MOSTRAR_TILESET Then
       '     MOSTRAR_TILESET = False
       '     Me_Tools_TileSet.EsconderVentanaTilesets
      '  End If
        '    If (ME_Tools.TOOL_SELECC And Tools.tool_tileset) Then
        '        If GetTickCount - TiempoBotonTileSetApretado > 200 Then
        '            MOSTRAR_TILESET = True
        '            Call Me_Tools_TileSet.MostrarVentanaTilesets(tileset_actual)
        '        End If
        '    End If
        'End If
  '  End If
    '------------------------------------------
    
    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTimer - lastMovement > (31) Then
        'TODO El GetTimer no se podia guardar en una variable y evitar el llamado a la funcion de la dll?
        lastMovement = GetTimer
    Else
        Exit Sub
    End If
    
    If Not IsAppActive() Then Exit Sub
    


    

            If GetAsyncKeyState(vbKeyNumpad0) Then       'In
                ZooMlevel = ZooMlevel + (timerElapsedTime * 0.003)
                If ZooMlevel > 2 Then ZooMlevel = 2
            ElseIf GetAsyncKeyState(vbKeyNumpad1) Then  'Out
                ZooMlevel = ZooMlevel - (timerElapsedTime * 0.003)
                If ZooMlevel < 0.25 Then ZooMlevel = 0.25
            End If
    'UserDirection = 0
    Dim kp As Boolean
            kp = (GetKeyState(vbKeyUp) < 0 Or _
                GetKeyState(vbKeyRight) < 0 Or _
                GetKeyState(vbKeyDown) < 0 Or _
                GetKeyState(vbKeyLeft) < 0 Or _
                GetKeyState(vbKeyS) < 0 Or _
                GetKeyState(vbKeyD) < 0 Or _
                GetKeyState(vbKeyA) < 0 Or _
                GetKeyState(vbKeyW) < 0) And frmMain.focoEnElRender

            If Not kp Then UserDirection = 0
            
    If UserDirection = 0 And UserMoving = 0 And frmMain.focoEnElRender = True Then

            
            If GetAsyncKeyState(vbKeyUp) < 0 Or GetAsyncKeyState(vbKeyW) < 0 Then
                UserDirection = NORTH
                Exit Sub
            End If
            
            'Move Right
            If GetAsyncKeyState(vbKeyRight) < 0 Or GetAsyncKeyState(vbKeyD) < 0 Then
                UserDirection = EAST
                Exit Sub
            End If
        
            'Move down
            If GetAsyncKeyState(vbKeyDown) < 0 Or GetAsyncKeyState(vbKeyS) < 0 Then
                UserDirection = SOUTH
                Exit Sub
            End If
        
            'Move left
            If GetAsyncKeyState(vbKeyLeft) < 0 Or GetAsyncKeyState(vbKeyA) < 0 Then
                UserDirection = WEST
                Exit Sub
            End If

    End If

End Sub


