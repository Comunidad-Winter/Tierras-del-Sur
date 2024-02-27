Attribute VB_Name = "Engine_Map"
'ESPEC�FICO DEL CLIENTE.E :E
Option Explicit

Enum layer_type
    layer_2 = 0
    layer_3 = 1
    layer_obj = 3
    layer_char = 4
    layer_roof = 5
    layer_particle = 6
    layer_agua = 7
    layer_terreno = 8
    layer_costa = 9
    layer_flare = 10
    layer_piso_capa3 = 11
    layer_entidades = 12
    layer_adornos_paredes = 13
End Enum

Public Const MAP_LAYER_ADORNOS_VERTICALES = 5 ' MAP_LAYER_OBJ_VERTICALES_TileLayer
Public Const MAP_LAYER_COSTAS = 1
Public Const MAP_LAYER_DECORADOS_PISO = 2
Public Const MAP_LAYER_OBJ_VERTICALES = 3
Public Const MAP_LAYER_TECHOS = 4

Public Const MAP_LAYER_COSTAS_TileLayer = 4
Public Const MAP_LAYER_DECORADOS_PISO_TileLayer = 1
Public Const MAP_LAYER_OBJ_VERTICALES_TileLayer = 2
Public Const MAP_LAYER_TECHOS_TileLayer = 3
Public Const MAP_TILESET_TileLayer = 0


Public Type tile
    tileX As Long
    tileY As Long
    PixelPosX As Single
    PixelPosY As Single
    type As Long
    id As Long
    alpha As Byte
End Type

Public Type TileLayer
    tile(2300) As tile
    NumTiles As Long
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 5) As Grh      ' OK
    ObjGrh As Grh               ' OK
    OBJInfo As obj              ' OK
    NpcIndex As Integer         ' OK
    
    trigger As Long             ' OK
        
    
    Particles_groups(0 To 2) As Engine_Particle_Group
    
    light_value(0 To 3) As Long 'Color de luz con el que esta siendo renderizado.

    tile_orientation As Byte
    tile As Box_Vertex
    tile_texture As Integer
    tile_render As Long
    tile_number As Integer

    flare As Byte
    is_water As Byte
    
    luz As Integer
    
    
    
    
    NpcZona As Byte ' Zona donde puede nacer esta criatura
    
    
    
    'Sonido al pisar
    EfectoPisada As Integer
#If esMe = 1 Then
    accion As iAccionEditor
#End If
End Type


Public Type ZonaNacimientoCriatura
     Superior As position
     Inferior As position
     #If esMe = 1 Then
        Nombre As String
     #End If
End Type

'Info de cada mapa
Public Type mapinfo
    Music As Integer
    Name As String
    startPos As WorldPos
    MapVersion As Integer

    ZonasNacCriaturas() As ZonaNacimientoCriatura

    flags As Long
    MaxGrhSizeXInTiles As Integer
    MaxGrhSizeYInTiles As Integer
    
    BaseColor As RGBCOLOR
    ColorPropio As Boolean
    
    agua_tileset As Integer
    agua_rect As RECT
    agua_profundidad As Integer
 
    puede_nieve As Boolean
    puede_lluvia As Boolean
    puede_neblina As Boolean
    puede_niebla As Boolean
    puede_sandstorm As Boolean
    puede_nublado As Boolean

    UsaAguatierra As Boolean

    SonidoLoop As Integer
    
    numero As Integer
    
    MapaNorte As Integer
    MapaSur As Integer
    MapaEste As Integer
    MapaOeste As Integer
    
    Frio As Boolean
        
    MinNivel As Byte
    MaxNivel As Byte
    
    UsuariosMaximo As Integer
    
    SeCaenItems As Boolean
    PermiteRoboNpc As Boolean
    PermiteHechizosPetes As Boolean
    MagiaSinEfecto As Boolean
    
    MapaPK As Boolean
End Type

Public Enum map_flags
    ColorPropio = 1
    
    RadioDeLuz = 2
    
    NoLlueve = 4
    NoNieva = 8
    NoNiebla = 16
    
    DUNGEON = 32
    CIUDAD = 64
    BOSQUE = 128
        
    Tiene_Agua = 256
End Enum

Public TileLayer(0 To 7) As TileLayer

Private techo_alpha As Byte

Public mapdata()                As MapBlock ' Mapa
Public mapinfo                  As mapinfo ' Info acerca del mapa en uso

Public screenminY               As Integer
Public screenmaxY               As Integer
Public screenminX               As Integer
Public screenmaxX               As Integer
Public MinY                     As Integer
Public MaxY                     As Integer
Public MinX                     As Integer
Public MaxX                     As Integer
Public minXOffset               As Integer
Public minYOffset               As Integer
Public tileX                    As Integer
Public tileY                    As Integer
Public add_to_mapY              As Integer

Public ScrollPixelsPerFrameX    As Integer
Public ScrollPixelsPerFrameY    As Integer

Public TileBufferPixelOffsetX   As Integer
Public TileBufferPixelOffsetY   As Integer

Public MinXBorder%, MaxXBorder%, MinYBorder%, MaxYBorder%

Public alpha_racio_luz As Single
Public alpha_neblina_llegar As Single

Public tileset_tex_w!, tileset_tex_h!

Private adya As Single

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numbytes As Long)

Public TileBufferSizeX As Integer
Public TileBufferSizeY As Integer

Public Timer_RerenderMap As New clsPerformanceTimer

Public UltimoParallaxCRC As Long

Public Screen_Desnivel_Offset As Single

Public AlphaTecho As New clsAlpha

Public bCameraCanged As Boolean

Public Const GRILLA_TEXTURA = 1102
Public Const GRILLA_ANCHO = 16
Public Const GRILLA_ALTO = 16
Public Const GRILLA_OFFSET_X = GRILLA_ANCHO - SV_Constantes.X_MINIMO_USABLE + 1
Public Const GRILLA_OFFSET_Y = GRILLA_ALTO - SV_Constantes.Y_MINIMO_USABLE + 1

Public Sub Engine_Calc_Screen_Moviment()
    Static rPixelOffsetX    As Single
    Static rPixelOffsetY    As Single

    Static DirY             As Integer

    Static tmp_offset           As D3DVECTOR2

    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim ady As Integer
    
    bCameraCanged = False
    
    If UserMoving Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            rPixelOffsetX = rPixelOffsetX - CharList(UserCharIndex).Velocidad.X * AddtoUserPos.X * timerTicksPerFrame
            If Abs(rPixelOffsetX) >= Abs(32 * AddtoUserPos.X) Then
                rPixelOffsetX = 0
                AddtoUserPos.X = 0
                UserMoving = 0
            End If
            bCameraCanged = True
        End If
            
        '****** Move screen Up and Down if needed ******
        
        If AddtoUserPos.Y <> 0 Then
            rPixelOffsetY = rPixelOffsetY - CharList(UserCharIndex).Velocidad.Y * AddtoUserPos.Y * timerTicksPerFrame
            If Abs(rPixelOffsetY) >= Abs(32 * AddtoUserPos.Y) Then
                rPixelOffsetY = 0
                AddtoUserPos.Y = 0
                UserMoving = 0
            End If
            bCameraCanged = True
        End If
        Cachear_Tiles = Cachear_Tiles Or bCameraCanged
        
        If AddtoUserPos.X = 0 And AddtoUserPos.Y = 0 Then
            UserMoving = 0
        End If
    End If
    
    'If copy_tile_now = 128 And UserMoving = 0 Then
    '    copy_tile_now = 255
    'Else
    '    copy_tile_now = 0
    'End If
    If UserPos.X = 0 Then
        UserPos.X = MinXBorder
        UserPos.Y = MinYBorder
    End If
    If AlturaPie(UserPos.X, UserPos.Y) > Screen_Desnivel_Offset Then
        DirY = 1
    ElseIf AlturaPie(UserPos.X, UserPos.Y) < Screen_Desnivel_Offset Then
        DirY = -1
    End If

    If DirY <> 0 Then
        Screen_Desnivel_Offset = Screen_Desnivel_Offset + 2 * Sgn(DirY) * timerTicksPerFrame
        If (Sgn(DirY) = 1 And Screen_Desnivel_Offset >= AlturaPie(UserPos.X, UserPos.Y)) Or (Sgn(DirY) = -1 And Screen_Desnivel_Offset <= AlturaPie(UserPos.X, UserPos.Y)) Then
            Screen_Desnivel_Offset = AlturaPie(UserPos.X, UserPos.Y)
            DirY = 0
        End If
        bCameraCanged = True
        'Cachear_Tiles = True
        'copy_tile_now = 255
    End If


    'add_to_mapY = (hMapData(UserPos.x, UserPos.Y).alt / 32) + 1
    
    'offset_mapO = offset_map
    
    If tileY <> UserPos.Y Or tileX <> UserPos.X Then
        'copy_tile_now = 128
        
        tileY = UserPos.Y - AddtoUserPos.Y
        tileX = UserPos.X - AddtoUserPos.X
        ady = Screen_Desnivel_Offset \ 32
        'Figure out Ends and Starts of screen
        screenminY = tileY - HalfWindowTileHeight ' - ady
        screenmaxY = tileY + HalfWindowTileHeight
        screenminX = tileX - HalfWindowTileWidth
        screenmaxX = tileX + HalfWindowTileWidth
        
                MinY = screenminY - TileBufferSizeY
        MaxY = screenmaxY + TileBufferSizeY
        MinX = screenminX - TileBufferSizeX
        MaxX = screenmaxX + TileBufferSizeX
        
        'Make sure mins and maxs are allways in map bounds
        If MinY < Y_MINIMO_VISIBLE Then
            minYOffset = Y_MINIMO_VISIBLE - MinY
            MinY = Y_MINIMO_VISIBLE
        Else
            minYOffset = 0
        End If
        
        If MaxY > Y_MAXIMO_VISIBLE Then MaxY = Y_MAXIMO_VISIBLE
        
        If MinX < X_MINIMO_VISIBLE Then
            minXOffset = X_MINIMO_VISIBLE - MinX
            MinX = X_MINIMO_VISIBLE
        Else
            minXOffset = 0
        End If
        
        If MaxX > Y_MAXIMO_VISIBLE Then MaxX = Y_MAXIMO_VISIBLE
        
        'If we can, we render around the view area to make it smoother
        If screenminY > Y_MINIMO_VISIBLE Then
            screenminY = screenminY - 1
        Else
            screenminY = 1
            ScreenY = 1
        End If
        
        If screenmaxY < Y_MAXIMO_VISIBLE Then screenmaxY = screenmaxY + 1
        
        If screenminX > X_MINIMO_VISIBLE Then
            screenminX = screenminX - 1
        Else
            screenminX = 1
            ScreenX = 1
        End If
        
        If screenmaxX < X_MAXIMO_VISIBLE Then screenmaxX = screenmaxX + 1
        'If minYOffset = 0 Then minYOffset = -ady
        'Cachear_Tiles = True
        adya = ady
        Map_render_2array adya
    End If
    
    
    offset_screen.Y = CInt(rPixelOffsetY + Screen_Desnivel_Offset) '+ minYOffset * 32)
    offset_screen.X = CInt(rPixelOffsetX) '+ minXOffset * 32)

    offset_mapO = offset_map
    offset_map.X = ((-MinX - 1) * 32) + offset_screen.X - TileBufferPixelOffsetX
    offset_map.Y = ((-MinY - 1) * 32) + offset_screen.Y - TileBufferPixelOffsetY

    offset_map_part.X = offset_map.X + minXOffset * 32
    offset_map_part.Y = offset_map.Y + minYOffset * 32
    
    SetCameraPixelPos offset_map_part.X, offset_map_part.Y
    
    Engine_LightsTexture_Render

    'EngineInterface.SetOffsets offset_map.X, offset_map.Y
End Sub

Public Sub clear_map_chars()
Dim X&, Y&
'For Y = 1 To 100
'For X = 1 To 100
'If MapData(X, Y).CharIndex Then EraseChar MapData(X, Y).CharIndex
'MapData(X, Y).CharIndex = 0
'Next X
'Next Y
'For X = 1 To 100
'ResetCharInfo X
'Next X
End Sub

Public Sub rm2a()
        Map_render_2array adya
End Sub

Public Sub Engine_Set_TileBuffer_Size(ByVal sizex As Integer, ByVal sizey As Integer)
    TileBufferSizeX = sizex
    TileBufferSizeY = sizey
    TileBufferPixelOffsetX = (TileBufferSizeX - 1) * TilePixelWidth
    TileBufferPixelOffsetY = (TileBufferSizeY - 1) * TilePixelWidth
End Sub

Public Sub act_charmap()
''Marce On error resume next
'    Dim i As Integer
'    ZeroMemory CharMap(1, 1), 10000 * 2
'    For i = 1 To LastChar
'    With CharList(i).Pos
'        If CharList(i).active Then _
'            If .x > UserPos.x - ARangoX And .x < UserPos.x + ARangoX And .Y > UserPos.Y - ARangoY And .Y < UserPos.Y + ARangoY Then _
'                CharMap(.x, .Y) = i
'    End With
'    Next i
End Sub

Private Sub map_add_tolayer(ByVal Layer As Byte, ByVal tileX As Byte, ByVal tileY As Byte, ByVal PixelOffsetX As Single, ByVal PixelOffsetY As Single, ByVal tipo As layer_type, ByVal id As Integer)
    TileLayer(Layer).NumTiles = TileLayer(Layer).NumTiles + 1
    With TileLayer(Layer).tile(TileLayer(Layer).NumTiles)
        .type = tipo
        .tileX = tileX
        .tileY = tileY
        .PixelPosX = PixelOffsetX
        .PixelPosY = PixelOffsetY
        .id = id
    End With
End Sub


Public Sub Map_render_2array(Optional ByVal offset_y As Single)

    Static LastRenderCRC As Long
    If LastRenderCRC = RENDERCRC Then
        'Debug.Print "-Redibujado del mapa cancelado por CRC"
        Exit Sub
    End If
    LastRenderCRC = RENDERCRC
    Dim Y                       As Long
    Dim X                       As Long
    
    Dim Layer                   As Long
    Dim PixelOffsetXTemp        As Single
    Dim PixelOffsetYTemp        As Single
    Dim ScreenX                 As Single
    Dim ScreenY                 As Single
    Dim tmphe%
    Dim tempha%

    'Timer_RerenderMap.Time
    Dim MinYloop As Long
    Dim MaxYloop As Long

    tmphe = WindowTileHeight + TileBufferSizeY - IIf(AlturaPie(UserPos.X, UserPos.Y) < 0, AlturaPie(UserPos.X, UserPos.Y) - 32, 0) \ 32
    
    If MinX = 0 Then Exit Sub
    
    For Layer = 0 To 6
        TileLayer(Layer).NumTiles = 0
        'ReDim TileLayer(Layer).tile(1 To ((maxY - minY + 1) * (maxX - minX + 1)))
    Next Layer

    hay_fogata_viewport = False

    tempha = -2 - Abs(offset_y * 32)
    ScreenY = minYOffset - TileBufferSizeY - offset_y
    
    MinYloop = maxl(MinY - offset_y, Y_MINIMO_VISIBLE)
    MaxYloop = minl(MaxY - offset_y, Y_MAXIMO_VISIBLE)
    
    Engine_LightsTexture_Clear
    Engine_Shadows_Clear
    
    For Y = MinYloop To MaxYloop

        ScreenX = minXOffset - TileBufferSizeX

        For X = MinX To MaxX

            PixelOffsetXTemp = ScreenX * 32
            PixelOffsetYTemp = ScreenY * 32
            
               
            With mapdata(X, Y)
                If Not .Particles_groups(0) Is Nothing Then
                    map_add_tolayer Engine_Map.MAP_LAYER_OBJ_VERTICALES_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_particle, 0 ' MapData(X, Y).Particles_groups(0)
                End If

                If .ObjGrh.GrhIndex Then
                    map_add_tolayer Engine_Map.MAP_LAYER_OBJ_VERTICALES_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_obj, 0
                End If
 
                If Not .Particles_groups(1) Is Nothing Then
                    map_add_tolayer Engine_Map.MAP_LAYER_OBJ_VERTICALES_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_particle, 1 'MapData(X, Y).Particles_groups(1)
                End If

                If .Graphic(Engine_Map.MAP_LAYER_OBJ_VERTICALES).GrhIndex Then
                    
                    map_add_tolayer Engine_Map.MAP_LAYER_OBJ_VERTICALES_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_3, 0
                 
                    If GrhData(.Graphic(Engine_Map.MAP_LAYER_OBJ_VERTICALES).GrhIndex).SombrasSize > 0 Then
                        Engine_Shadows_Add X * 32 + 16, Y * 32 - 16, GrhData(.Graphic(Engine_Map.MAP_LAYER_OBJ_VERTICALES).GrhIndex).SombrasSize
                    End If
                End If
                
                If .Graphic(Engine_Map.MAP_LAYER_ADORNOS_VERTICALES).GrhIndex Then
                    map_add_tolayer Engine_Map.MAP_LAYER_OBJ_VERTICALES_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_adornos_paredes, 0
                End If
                
                If EntidadesMap(X, Y) <> 0 Then
                    map_add_tolayer Engine_Map.MAP_LAYER_OBJ_VERTICALES_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_entidades, 0
                End If
                
                If Not .Particles_groups(2) Is Nothing Then
                    map_add_tolayer Engine_Map.MAP_LAYER_OBJ_VERTICALES_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_particle, 2 'MapData(X, Y).Particles_groups(2)
                End If
                
                If .Graphic(MAP_LAYER_TECHOS).GrhIndex Then
                    map_add_tolayer MAP_LAYER_TECHOS_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_roof, 0
                End If
                
                ' Me fijo si hay luces en el rango, la agrego a la pila de dibujado
                ' de las luces para normales
                If .luz Then Engine_LightsTexture_Push .luz
                
                If CharMap(X, Y) > 0 Then
                    If CharList(CharMap(X, Y)).luz > 0 Then
                        Engine_LightsTexture_Push CharList(CharMap(X, Y)).luz
                    End If
                    
                    Engine_Shadows_Add CharList(CharMap(X, Y)).Pos.X * 32 + CharList(CharMap(X, Y)).MoveOffsetX + 16, CharList(CharMap(X, Y)).Pos.Y * 32 + CharList(CharMap(X, Y)).MoveOffsetY - 8, 16
                End If
                
                
                'If X > UserPos.X - MinXBorder And X < UserPos.X + MinXBorder And y > UserPos.y - MinYBorder - 1 And y < UserPos.y + MinYBorder + 1 Then
                
                If ScreenY > tempha And ScreenX > -2 Then
                    If ScreenY <= tmphe And ScreenX <= WindowTileWidth + TileBufferSizeX Then


                        If .Graphic(Engine_Map.MAP_LAYER_COSTAS).GrhIndex > 1 Then
                            map_add_tolayer Engine_Map.MAP_LAYER_COSTAS_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_costa, 0
                        End If
                        
                        If .Graphic(Engine_Map.MAP_LAYER_DECORADOS_PISO).GrhIndex Then
                            map_add_tolayer Engine_Map.MAP_LAYER_DECORADOS_PISO_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_2, 0
                        End If
                        
                        If CharMap(X, Y) <> 0 Then
                            If CharList(CharMap(X, Y)).active Then
                                map_add_tolayer Engine_Map.MAP_LAYER_OBJ_VERTICALES_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_char, CharMap(X, Y)
                            End If
                        End If

                        If .tile_texture <> 0 Then
                            map_add_tolayer Engine_Map.MAP_TILESET_TileLayer, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_terreno, .tile_texture
                        End If
                        
                        If ScreenY = tmphe Then
                            If hMapData(X, Y).h <> 0 Then
                                If PixelOffsetYTemp - hMapData(X, Y).hs(1) - Abs(offset_y * 32) < MainViewHeight Then
                                    tmphe = tmphe + 1
                                    MaxYloop = minl(MaxYloop + 1, Y_MAXIMO_VISIBLE)
                                End If
                            End If
                        End If
                    End If
                End If
                
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    

Exit Sub

End Sub



Public Sub Render_Map_Array()
    Dim Y                   As Integer
    Dim X                   As Integer
    
    Dim ScreenX             As Integer
    Dim ScreenY             As Integer
    
    Dim CurrentGrhIndex     As Long
    Dim tx!, ty!
   
    Dim MX!, MY!
    Dim MMX%, MMY%
    
    Dim i                   As Integer
    Dim TempStr             As String
    Dim Color               As Long
    
    Dim offsetX             As Single
    Dim offsetY             As Single
    
    PS_SetearColoresAmbiente
    
    'If frmVertexShader.Check1.value = vbUnchecked Then
        If mapinfo.UsaAguatierra Then
            For i = 1 To TileLayer(MAP_TILESET_TileLayer).NumTiles
                With TileLayer(MAP_TILESET_TileLayer).tile(i)
                    Grh_Render_Tileset .id, .tileX, .tileY, hMapData(.tileX, .tileY), MapBoxes(.tileX, .tileY)
                    'If MapData(.tilex, .tiley).is_water Then
                    If AguaVisiblePosicion(.tileX, .tileY) Then
                        Grh_Render_Water .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tileX, .tileY
                    End If
                End With
            Next i
        Else
            For i = 1 To TileLayer(MAP_TILESET_TileLayer).NumTiles
                With TileLayer(MAP_TILESET_TileLayer).tile(i)
                    Grh_Render_Tileset .id, .tileX, .tileY, hMapData(.tileX, .tileY), MapBoxes(.tileX, .tileY)
                End With
            Next i
        End If
    'Else
    '    MapBox_Draw
    'End If
    Cachear_Tiles = False

    For i = 1 To TileLayer(MAP_LAYER_COSTAS_TileLayer).NumTiles
        With TileLayer(MAP_LAYER_COSTAS_TileLayer).tile(i)
                With mapdata(.tileX, .tileY).Graphic(MAP_LAYER_COSTAS)
                    If .Started = 1 Then
                        .FrameCounter = .FrameCounter + (timerElapsedTime * GrhData(.GrhIndex).NumFrames / .Speed)
                        If .FrameCounter > GrhData(.GrhIndex).NumFrames Then
                            .FrameCounter = (.FrameCounter Mod GrhData(.GrhIndex).NumFrames) + 1
                            If .Loops <> -1 Then
                                If .Loops > 0 Then
                                    .Loops = .Loops - 1
                                Else
                                    .Started = 0
                                End If
                            End If
                        End If
                        
                       ' CurrentGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
                    'Else
                        'CurrentGrhIndex = 0
                        
                    End If
                    CurrentGrhIndex = GrhData(.GrhIndex).frames(Fix(.FrameCounter))
                End With
                If CurrentGrhIndex > 1 Then
                        Grh_Render_relieve CurrentGrhIndex, _
                            .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, _
                            .tileX, .tileY
                End If
        End With
    Next i
    

    
    For i = 1 To TileLayer(MAP_LAYER_DECORADOS_PISO_TileLayer).NumTiles
        With TileLayer(MAP_LAYER_DECORADOS_PISO_TileLayer).tile(i)
                    Draw_Grh mapdata(.tileX, .tileY).Graphic(MAP_LAYER_DECORADOS_PISO), _
                            .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, 1, _
                            .tileX, .tileY, 1

        End With
    Next i

    Sangre_Render

    For i = 1 To TileLayer(MAP_LAYER_OBJ_VERTICALES_TileLayer).NumTiles
        With TileLayer(MAP_LAYER_OBJ_VERTICALES_TileLayer).tile(i)
            Select Case .type
                Case layer_type.layer_3
                    If (mapdata(.tileX, .tileY).trigger And eTriggers.Transparentar) And CharList(UserCharIndex).Pos.X >= .tileX - 3 And CharList(UserCharIndex).Pos.X <= .tileX + 3 Then
                        Call Draw_Grh_Techo(mapdata(.tileX, .tileY).Graphic(Engine_Map.MAP_LAYER_OBJ_VERTICALES), _
                                    .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, 0, 125)
                    Else
                        Call Draw_Grh(mapdata(.tileX, .tileY).Graphic(Engine_Map.MAP_LAYER_OBJ_VERTICALES), _
                                .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, 1, .tileX, .tileY)
                    End If
                Case layer_type.layer_adornos_paredes
                    Call Draw_Grh(mapdata(.tileX, .tileY).Graphic(Engine_Map.MAP_LAYER_ADORNOS_VERTICALES), _
                                .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, 1, .tileX, .tileY)
                'Case layer_type.layer_piso_capa3
                '    Engine.Grh_Render_Tileset .id, .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tilex, .tiley, mapdata(.tilex, .tiley).tile_orientation
                '    If mapdata(.tilex, .tiley).is_water Then
                '        Grh_Render_Water .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tilex, .tiley
                '    End If
                Case layer_type.layer_obj
    
                    If mapdata(.tileX, .tileY).ObjGrh.GrhIndex <> 0 Then
                            Call Draw_Grh(mapdata(.tileX, .tileY).ObjGrh, _
                                        .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, 1, .tileX, .tileY)
                    End If
                Case layer_type.layer_entidades
                    If EntidadesMap(.tileX, .tileY) <> 0 Then _
                        Call Entidades_Render_Recursivo(.PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tileX, .tileY)
                Case layer_type.layer_char
                    Call Char_Render(.id)
                Case layer_type.layer_particle
                    If Not mapdata(.tileX, .tileY).Particles_groups(.id) Is Nothing Then
                        If mapdata(.tileX, .tileY).Particles_groups(.id).Render = False Then
                            Set mapdata(.tileX, .tileY).Particles_groups(.id) = Nothing
                        End If
                    End If
            End Select
        End With
    Next i
    
    FX_Hit_Render

    'render_blood offset_screen.x, offset_screen.Y
    'Sangre.Render

    Projectile_Render

    For i = 1 To TileLayer(MAP_LAYER_TECHOS_TileLayer).NumTiles
        With TileLayer(MAP_LAYER_TECHOS_TileLayer).tile(i)
            If .type = layer_type.layer_particle Then
                If Not mapdata(.tileX, .tileY).Particles_groups(.id) Is Nothing Then
                    If mapdata(.tileX, .tileY).Particles_groups(.id).Render = False Then
                        Set mapdata(.tileX, .tileY).Particles_groups(.id) = Nothing
                    End If
                End If
            ElseIf .type = layer_type.layer_roof Then
                If techo_alpha Then
                    Call Draw_Grh_Techo(mapdata(.tileX, .tileY).Graphic(MAP_LAYER_TECHOS), _
                        .PixelPosX + offset_screen.X, _
                        .PixelPosY + offset_screen.Y, 1, techo_alpha)
                End If
            End If

        End With
    Next i

    'Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_DISABLE)

    'WeatherDoFog = frmMain.VScroll1.value \ 15
    
    Engine_Weather_UpdateFog
    If (estado_time And Tipos_Clima.ClimaTormenta_de_arena) And mapinfo.puede_sandstorm Then Engine_Weather_SandStorm

    If alpha_racio_luz > 0 Then
        Dim cc As Long
        cc = D3DColorARGB(alpha_racio_luz, 0, 0, 0)
        Grh_Render_Simple_box 1131, 16!, -50!, cc, 512!
        Engine.Draw_FilledBox 0, 0, 16, frmMain.Renderer.Height, cc, 0, 0
        Engine.Draw_FilledBox 528, 0, 17, frmMain.Renderer.Height, cc, 0, 0
    End If

#If esMe = 1 Then
    Dim c As Long
    
    If Render_Radio_Luz Then
        c = D3DColorARGB(255, 0, 0, 0)
        Grh_Render_Simple_box 1131, 244!, 0, c, 512!
        Engine.Draw_FilledBox 0, 0, 244, frmMain.Renderer.Height, c, 0, 0
        Engine.Draw_FilledBox 756, 0, 244, frmMain.Renderer.Height, c, 0, 0
    End If
    
    If DRAWBLOQUEOS > 0 Or DRAWTRIGGERS > 0 Or dibujarZonaNacimientoCriatura > 0 Or dibujarAccionTile > 0 Or DRAWGRILLA > 0 Or dibujarCantidadObjetos Or mostrarTileDondeHayLuz > 0 Or mostrarTileNumber > 0 Or mostrarTileEfectoSonidoPasos > 0 Then
        MMX = minl(X_MAXIMO_VISIBLE, screenmaxX)
        MMY = minl(Y_MAXIMO_VISIBLE, screenmaxY)
        For Y = screenminY To MMY
            For X = screenminX To MMX
                    tx = (X + minXOffset) * 32 + offset_map.X
                    ty = (Y + minYOffset) * 32 + offset_map.Y
                    
                    ' Dibujo los bloqueos
                    If DRAWBLOQUEOS > 0 And (mapdata(X, Y).trigger And &H1F) > 0 Then
                        Grh_Render_Bloqueos tx, ty, X, Y
                    End If
                    
                    ' Dibujo letras que identifican a los triggers
                    If mapdata(X, Y).trigger > 0 And DRAWTRIGGERS > 0 Then
                            Grh_Render_Blocked &H7F000000, _
                                tx, ty, X, Y
                    
                            Engine.Text_Render_ext CStr(ME_Tools_Triggers.obtenerDescripcionAbreviatura(mapdata(X, Y).trigger)), ty + 8, tx + 12, 0, 0, &HFFFFFFFF
                    End If
                    
                    '�Dibujo la informaci�n de las acciones?
                    If dibujarAccionTile Then
                        If Not mapdata(X, Y).accion Is Nothing Then
                                If MouseTileX = X And MouseTileY = Y Then
                                    Grh_Render_Blocked mzRed And &HEEFFFFFF, tx, ty, X, Y
                                    Engine.Text_Render_ext mapdata(X, Y).accion.GetNombre, ty, tx + 5, 0, 0, &HFFFFFFFF
                                Else
                                    Grh_Render_Blocked mzRed And &H30FFFFFF, tx, ty, X, Y
                                    Engine.Text_Render_ext mapdata(X, Y).accion.GetId, ty + 7, tx + 5, 0, 0, &HFFFFFFFF
                                End If
                        End If
                    End If
                    
                    ' Dibujo la grilla? Tres colores distintos...
                    If DRAWGRILLA > 0 Then
                        Dim puedeModificarComportamiento As Boolean
                        Dim puedeModificarVisual As Boolean
                        
                        'Permisos que se tienen sobre esta posicion
                        puedeModificarComportamiento = ME_Mundo.puedeModificarComporamientoTile(X, Y)
                        puedeModificarVisual = ME_Mundo.puedeModificarAspectoTile(X, Y)
                        
                        If puedeModificarComportamiento Then
                            If ME_Mundo.esVisibleEnOtroMapa(X, Y) Then
                                Color = &HFF646464
                            Else
                                Color = &HFFFFFFFF
                            End If
                        ElseIf puedeModificarVisual Then
                            Color = &HFF009900
                        Else
                            Color = &HFF990000
                        End If
                        
                        offsetX = (GRILLA_OFFSET_X + X - 1) Mod 16 + 16 * ((GRILLA_OFFSET_Y + Y - 1) Mod 16)
                         
                        Grh_Render_Relieve_Tileset_HCD GRILLA_TEXTURA, offsetX, tx, ty, X, Y, Color
                    End If
                    
                    '�Dibujo la cantidad de objetos?
                    If dibujarCantidadObjetos Then
                        If mapdata(X, Y).OBJInfo.OBJIndex > 0 Then
                             Engine.Text_Render_ext CStr(mapdata(X, Y).OBJInfo.Amount), ty + 7, tx + 5, 100, 10, &HFFFFFFFF
                        End If
                    End If
                    
                    '�Dibujo laz ZONAS
                    If dibujarZonaNacimientoCriaturas Then
                        TempStr = Me_Tools_Npc.obtenerDescripcionAbreviaturaTile(X, Y, mapinfo.ZonasNacCriaturas)
                        If Not TempStr = "" Then
                            Engine.Text_Render_ext CStr(TempStr), ty + 24, tx + 0, 0, 0, &HFFFF0080
                        End If
                    End If
                    
                    
                    '�Dibujo arriba del NPC las zonas donde nace?
                    If dibujarZonaNacimientoCriatura Then
                        If mapdata(X, Y).NpcIndex > 0 Then
                            If mapdata(X, Y).NpcZona > 0 Then
                                Engine.Text_Render_ext CStr(Me_Tools_Npc.obtenerDescripcionAbreviatura(mapdata(X, Y).NpcZona, mapinfo.ZonasNacCriaturas)), ty + 24, tx + 12, 0, 0, &HFFFF8000
                            End If
                        End If
                    End If
                    
                    '�Dibujo donde hay luz?
                    If mostrarTileDondeHayLuz Then
                        If mapdata(X, Y).luz > 0 Then
                             Engine.Text_Render_ext "L", ty + 20, tx + 20, 100, 10, &HFFFFFFFF
                        End If
                    End If
                    
                    'Dibujo el numero de tile?
                    If mostrarTileNumber Then
                        If mapdata(X, Y).tile_texture > 0 Then
                            Engine.Text_Render_ext CStr(mapdata(X, Y).tile_number), ty, tx, 100, 10, &HFFFFFFFF
                        End If
                    End If
                    
                    ' Dibujo el efecto de sonido que se ejecuta al pisar?
                    If mostrarTileEfectoSonidoPasos Then
                        If mapdata(X, Y).EfectoPisada > 0 Then
                            Engine.Text_Render_ext CStr(mapdata(X, Y).EfectoPisada), ty, tx, 100, 10, &HFF00FFFF
                        End If
                    End If
                    
                    If mostrarTileDondeHayGraficos Then
                        If mapdata(X, Y).Graphic(1).GrhIndex Or mapdata(X, Y).Graphic(2).GrhIndex Or mapdata(X, Y).Graphic(3).GrhIndex Or mapdata(X, Y).Graphic(4).GrhIndex Or mapdata(X, Y).Graphic(5).GrhIndex Then
                            Engine.Text_Render_ext "*", ty, tx, 100, 10, &HFF000FFF
                        End If
                    End If
            Next X
        Next Y
    End If

    ' Simular estar en el cliente. Voy a poner cuadrados oscuros en la zona no visible
    
    If DRAWCLIENTAREA Then
        ' Verticales para acortar a lo ancho
        ScreenX = (Engine.MainViewWidth - ClientWindowWidth) \ 2
        
        Engine.Draw_FilledBox 0, 0, ScreenX, Engine.MainViewHeight, mzBlack, 0, 0 'Derecha
        Engine.Draw_FilledBox Engine.MainViewWidth - ScreenX, 0, ScreenX, Engine.MainViewHeight, mzBlack, 0, 0 'Izquierda
        
        ' Horizontales para acortar a lo largo
        ScreenY = (Engine.MainViewHeight - ClientWindowHeight) \ 2
        
        Engine.Draw_FilledBox ScreenX, 0, Engine.MainViewWidth - ScreenX, ScreenY, mzBlack, 0, 0 ' Arriba
        Engine.Draw_FilledBox ScreenX, Engine.MainViewHeight - ScreenY, Engine.MainViewWidth - ScreenX, ScreenY, mzBlack, 0, 0    ' Abajo
    End If
#End If
End Sub

Public Property Let bTecho(T As Boolean)
    Static EstabaBajoTecho As Boolean
    Static SeteadoAlgunaVez As Boolean
    
    #If esMe = 1 Then
        If dibujarTechosTransparentes Then
            AlphaTecho.value = 128
            SeteadoAlgunaVez = False
            Exit Property
        End If
    #End If
    
    If T <> EstabaBajoTecho Or SeteadoAlgunaVez = False Then
        SeteadoAlgunaVez = True
        pbTecho = T
        EstabaBajoTecho = T
        
        If pbTecho Then
            AlphaTecho.value = 0
        Else
            AlphaTecho.value = 255
        End If
    End If
End Property

Public Property Get bTecho() As Boolean
    bTecho = pbTecho
End Property

Public Sub Map_Render()
    Static UltimoAguaTierra As Long
    
    'If mapinfo.UsaAguatierra And AnimarAguatierra And (GetTimer - UltimoAguaTierra) > 64 Then
        'kWATER = UltimoAguaTierra Mod 360
        'map_render_kwateR
        'UltimoAguaTierra = GetTimer
    'End If
    
'    If bTecho Then
'        AlphaTecho.Value = 0
'    Else
'        AlphaTecho.Value = 255
'    End If
    
    techo_alpha = AlphaTecho.value And &HFF
    

    'Debug.Print Hex(techo_alpha)
    'Engine.SetVertexShader FVF
    
    Render_Map_Array
    
    If Lightbeam_do Then Engine_Weather_lightbeam

    If Not meteo_particle Is Nothing Then
        'meteo_particle.SetPixelPos -offset_map.X + 250, -offset_map.y
        If meteo_particle.Render = False Then
            Set meteo_particle = Nothing
        End If
    End If

End Sub

#If esMe = 1 Then
    Public Sub AgregarZona(area As tAreaSeleccionada, ByVal Nombre As String)
        Dim loopZona As Byte
        Dim indice As Byte
        
        indice = 255
        
        For loopZona = LBound(mapinfo.ZonasNacCriaturas) To UBound(mapinfo.ZonasNacCriaturas)
            If mapinfo.ZonasNacCriaturas(loopZona).Nombre = "" Or LCase$(mapinfo.ZonasNacCriaturas(loopZona).Nombre) = LCase$(Nombre) Then
                indice = loopZona
                Exit For
            End If
        Next loopZona
        
        If indice = 255 Then
            indice = UBound(mapinfo.ZonasNacCriaturas) + 1
            ReDim Preserve mapinfo.ZonasNacCriaturas(indice) As ZonaNacimientoCriatura
        End If
        
        mapinfo.ZonasNacCriaturas(indice).Nombre = Nombre
                
        mapinfo.ZonasNacCriaturas(indice).Superior.Y = area.arriba
        mapinfo.ZonasNacCriaturas(indice).Superior.X = area.izquierda
        mapinfo.ZonasNacCriaturas(indice).Inferior.X = area.derecha
        mapinfo.ZonasNacCriaturas(indice).Inferior.Y = area.abajo
    End Sub
    
    Public Function zonaExiste(ByVal Nombre As String) As Boolean
        Dim loopZona As Byte
        
        Nombre = LCase$(Nombre)
        
        For loopZona = LBound(mapinfo.ZonasNacCriaturas) To UBound(mapinfo.ZonasNacCriaturas)
            If LCase$(mapinfo.ZonasNacCriaturas(loopZona).Nombre) = Nombre Then
                zonaExiste = True
                Exit Function
            End If
        Next
        
        zonaExiste = False
    
    End Function
#End If
