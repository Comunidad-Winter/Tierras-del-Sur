Attribute VB_Name = "ME_Engine_Map"


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
End Enum

Public Type tile
    tilex As Byte
    tiley As Byte
    PixelPosX As Single
    PixelPosY As Single
    type As Integer
    ID As Integer
End Type

Public Type TileLayer
    tile(1700) As tile
    NumTiles As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    ObjGrh As Grh
    CharIndex As Integer
    Particles_groups(0 To 2) As Integer
    Particles_groups_original(0 To 2) As Integer
    light_value(0 To 3) As Long 'Color de luz con el que esta siendo renderizado.

    tile_orientation As Byte
    tile As Box_Vertex
    tile_texture As Integer
    tile_render As Byte
    tile_number As Integer

    flare As Byte
    is_water As Byte
    
    luz As Integer
    
    OBJInfo As obj
    
    NpcIndex As Integer
    
    Trigger As Long
#If esME = 1 Then
    accion As iAccionEditor
#End If
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer

    Flags As Long
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

Public TileLayer(1 To 6) As TileLayer

Private techo_alpha As Byte

Public MapData()                As MapBlock ' Mapa
Public MapInfo                  As MapInfo ' Info acerca del mapa en uso

Public screenminY               As Integer
Public screenmaxY               As Integer
Public screenminX               As Integer
Public screenmaxX               As Integer
Public minY                     As Integer
Public maxY                     As Integer
Public minX                     As Integer
Public maxX                     As Integer
Public minXOffset               As Integer
Public minYOffset               As Integer
Public tilex                    As Integer
Public tiley                    As Integer
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

Public Sub Engine_Calc_Screen_Moviment()
    Static rPixelOffsetX    As Single
    Static rPixelOffsetY    As Single

    Static DirY             As Integer

    Static tmp_offset 		As D3DVECTOR2

    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim ady As Integer
    
    If UserMoving Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            rPixelOffsetX = rPixelOffsetX - CharList(UserCharIndex).Velocidad.X * AddtoUserPos.X * timerTicksPerFrame
            If Abs(rPixelOffsetX) >= Abs(32 * AddtoUserPos.X) Then
                rPixelOffsetX = 0
                AddtoUserPos.X = 0
                UserMoving = 0
            End If
        End If
            
        '****** Move screen Up and Down if needed ******
        
        If AddtoUserPos.Y <> 0 Then
            rPixelOffsetY = rPixelOffsetY - CharList(UserCharIndex).Velocidad.Y * AddtoUserPos.Y * timerTicksPerFrame
            If Abs(rPixelOffsetY) >= Abs(32 * AddtoUserPos.Y) Then
                rPixelOffsetY = 0
                AddtoUserPos.Y = 0
                UserMoving = 0
            End If
        End If
        Cachear_Tiles = True
        If AddtoUserPos.X = 0 And AddtoUserPos.Y = 0 Then
            UserMoving = 0
        End If
        'copy_tile_now = 128
    Else
        If UserDirection Then MoveTo UserDirection
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
                Cachear_Tiles = True
        'copy_tile_now = 255
    End If
    offset_screen_old = offset_screen

    add_to_mapY = (AlturaPie(UserPos.X, UserPos.Y) / 32) + 1
    
    'offset_mapO = offset_map
    
    #If EnableParallax = 1 Then
    Map_render_parallax rPixelOffsetX, rPixelOffsetY
    #End If
    
    
    If tiley <> UserPos.Y Or tilex <> UserPos.X Then
        'copy_tile_now = 128
        user_moved = True
        tiley = UserPos.Y - AddtoUserPos.Y
        tilex = UserPos.X - AddtoUserPos.X
        ady = Screen_Desnivel_Offset \ 32
        'Figure out Ends and Starts of screen
        screenminY = tiley - HalfWindowTileHeight ' - ady
        screenmaxY = tiley + HalfWindowTileHeight
        screenminX = tilex - HalfWindowTileWidth
        screenmaxX = tilex + HalfWindowTileWidth
        
		minY = screenminY - TileBufferSizeY
        maxY = screenmaxY + TileBufferSizeY
        minX = screenminX - TileBufferSizeX
        maxX = screenmaxX + TileBufferSizeX
        
        'Make sure mins and maxs are allways in map bounds
        If minY < Y_MINIMO_VISIBLE Then
            minYOffset = Y_MINIMO_VISIBLE - minY
            minY = Y_MINIMO_VISIBLE
        Else
            minYOffset = 0
        End If
        
        If maxY > Y_MAXIMO_VISIBLE Then maxY = Y_MAXIMO_VISIBLE
        
        If minX < X_MINIMO_VISIBLE Then
            minXOffset = X_MINIMO_VISIBLE - minX
            minX = X_MINIMO_VISIBLE
        Else
            minXOffset = 0
        End If
        
        If maxX > Y_MAXIMO_VISIBLE Then maxX = Y_MAXIMO_VISIBLE
        
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
        Cachear_Tiles = True
        adya = ady
        Map_render_2array adya
    End If
    
    
    offset_screen.Y = CInt(rPixelOffsetY + Screen_Desnivel_Offset) '+ minYOffset * 32)
    offset_screen.X = CInt(rPixelOffsetX) '+ minXOffset * 32)

    offset_mapO = offset_map
    offset_map.X = ((-minX - 1) * 32) + offset_screen.X - TileBufferPixelOffsetX
    offset_map.Y = ((-minY - 1) * 32) + offset_screen.Y - TileBufferPixelOffsetY


    EngineInterface.SetOffsets offset_map.X, offset_map.Y
End Sub




Public Sub Engine_Set_TileBuffer_Size(ByVal sizex As Integer, ByVal sizey As Integer)
    TileBufferSizeX = sizex
    TileBufferSizeY = sizey
    TileBufferPixelOffsetX = (TileBufferSizeX - 1) * TilePixelWidth
    TileBufferPixelOffsetY = (TileBufferSizeY - 1) * TilePixelWidth
End Sub


























Private Sub map_add_tolayer(ByVal Layer As Byte, ByVal tilex As Byte, ByVal tiley As Byte, ByVal PixelOffsetX As Single, ByVal PixelOffsetY As Single, ByVal tipo As layer_type, ByVal ID As Integer)
    TileLayer(Layer).NumTiles = TileLayer(Layer).NumTiles + 1
    With TileLayer(Layer).tile(TileLayer(Layer).NumTiles)
        .type = tipo
        .tilex = tilex
        .tiley = tiley
        .PixelPosX = PixelOffsetX
        .PixelPosY = PixelOffsetY
        .ID = ID
    End With
End Sub

#If EnableParallax = 1 Then
Public Sub Map_render_parallax(Optional ByVal offset_x As Single, Optional ByVal offset_y As Single)

On Error GoTo enda:

    If UltimoParallaxCRC = RENDERCRC Then
        'Debug.Print "-Redibujado del mapa cancelado por CRC"
        Exit Sub
    End If
    UltimoParallaxCRC = RENDERCRC
    Dim Y                       As Integer
    Dim X                       As Integer
    Dim lY                      As Byte
    Dim Layer                   As Byte
    Dim PixelOffsetXTemp        As Single
    Dim PixelOffsetYTemp        As Single
    Dim ScreenX                 As Single
    Dim ScreenY                 As Single
    Dim tmphe%
    Dim tempha%
    


    If minX = 0 Then Exit Sub

    ScreenY = minYOffset - TileBufferSizeY
        
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSizeX
        For X = minX To maxX

            PixelOffsetXTemp = ScreenX * 32
            PixelOffsetYTemp = ScreenY * 32

                If hMapData(X, Y).h > 0 Then
                    With ParallaxOffsets(X, Y)
'                        .X = ((PixelOffsetXTemp + D3DWindow.BackBufferWidth / 2) / D3DWindow.BackBufferWidth) * hMapData(X, Y).hs(0) * 1 / 5
'                        .Y = ((PixelOffsetYTemp + D3DWindow.BackBufferHeight / 2) / D3DWindow.BackBufferHeight) * hMapData(X, Y).hs(0) * 1 / 5
                        '.X = ((PixelOffsetXTemp + offset_map.X / 2) / D3DWindow.BackBufferWidth) * hMapData(X, Y).hs(0) * 1
                        '.Y = ((PixelOffsetYTemp + offset_map.Y / 2) / D3DWindow.BackBufferWidth) * hMapData(X, Y).hs(0) * 1 - hMapData(X, Y).hs(0)
GetParalllaxOffset PixelOffsetXTemp, PixelOffsetYTemp, hMapData(X, Y).hs(0), .X, .Y
.Y = .Y - hMapData(X, Y).hs(0)

                    End With

                Else
                    ParallaxOffsets(X, Y).X = 0
                    ParallaxOffsets(X, Y).Y = 0
                End If
                
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    'Cachear_Tiles = True
Exit Sub
enda:
If Err.number = 10 Then
LogError "ERROR 10 EN m2a"
End If
Debug.Assert 0
End Sub
#End If

Public Sub Map_render_2array(Optional ByVal offset_y As Single)
On Error GoTo enda:
    Static LastRenderCRC As Long
    If LastRenderCRC = RENDERCRC Then
        'Debug.Print "-Redibujado del mapa cancelado por CRC"
        Exit Sub
    End If
    LastRenderCRC = RENDERCRC
    Dim Y                       As Integer
    Dim X                       As Integer
    Dim lY                      As Byte
    Dim Layer                   As Byte
    Dim PixelOffsetXTemp        As Single
    Dim PixelOffsetYTemp        As Single
    Dim ScreenX                 As Single
    Dim ScreenY                 As Single
    Dim tmphe%
    Dim tempha%

    'Timer_RerenderMap.Time
    Dim MinYloop As Integer
    Dim MaxYloop As Integer

    tmphe = WindowTileHeight
    If minX = 0 Then Exit Sub
    
    For Layer = 1 To 6
        TileLayer(Layer).NumTiles = 0
        'ReDim TileLayer(Layer).tile(1 To ((maxY - minY + 1) * (maxX - minX + 1)))
    Next Layer

    hay_fogata_viewport = False

	tempha = -2 - offset_y * 32
    ScreenY = minYOffset - offset_y - TileBufferSizeY
    
    MinYloop = maxl(minY - offset_y, Y_MINIMO_VISIBLE)
    MaxYloop = minl(maxY - offset_y, Y_MAXIMO_VISIBLE)
    
    For Y = MinYloop To MaxYloop

        ScreenX = minXOffset - TileBufferSizeX

        For X = minX To maxX

            PixelOffsetXTemp = ScreenX * 32
            PixelOffsetYTemp = ScreenY * 32
            With MapData(X, Y)
                ly = 2
                If .Particles_groups(0) Then
                    map_add_tolayer ly, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_particle, MapData(X, Y).Particles_groups(0)
                End If

                If .ObjGrh.GrhIndex Then
                    map_add_tolayer ly, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_obj, 0
                    'If .ObjGrh.GrhIndex = GrhFogata Then
                    '    hay_fogata_viewport = True
                    '    fogata_pos.x = x
                    '    fogata_pos.Y = Y
                    'End If
                End If

                If .ObjGrh.GrhIndex Then map_add_tolayer lY, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_obj, 0
 
                If .Particles_groups(1) Then map_add_tolayer lY, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_particle, MapData(X, Y).Particles_groups(1)

                If EntidadesMap(X, Y) Then map_add_tolayer lY, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_entidades, 0

                If .Graphic(3).GrhIndex Then map_add_tolayer lY, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_3, 0

                If .Particles_groups(2) Then map_add_tolayer lY, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_particle, MapData(X, Y).Particles_groups(2)

                lY = 3

                If .Graphic(4).GrhIndex Then If techo_alpha Then map_add_tolayer lY, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_roof, 0

				'If X > UserPos.X - MinXBorder And X < UserPos.X + MinXBorder And Y > UserPos.Y - MinYBorder - 1 And Y < UserPos.Y + MinYBorder + 1 Then

                If ScreenY > tempha And ScreenX > -2 Then
                    If ScreenY <= tmphe And ScreenX <= WindowTileWidth Then

                        If .Graphic(1).GrhIndex > 1 Then map_add_tolayer 4, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_costa, 0

                        If .Graphic(2).GrhIndex Then map_add_tolayer 1, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_2, 0

                        If CharMap(X, Y) Then 
							If CharList(CharMap(X, Y)).active Then
								map_add_tolayer 2, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_type.layer_char, CharMap(X, Y)
							End If
						End If

                        If .tile_texture > 0 Then map_add_tolayer 6, X, Y, PixelOffsetXTemp, PixelOffsetYTemp, layer_terreno, .tile_texture

                        If ScreenY = tmphe Then
							If AlturaPie(X, Y) <> 0 Then
								If PixelOffsetYTemp - AlturaPie(X, Y) < MainViewHeight Then
									tmphe = tmphe + 1
								end if
							End if	
						End if
					End If
                End If
                
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    act_light_map = True
Exit Sub
enda:
If Err.number = 10 Then
LogError "ERROR 10 EN m2a"
End If
Debug.Assert 0
End Sub



Public Sub Render_Map_Array()
    Dim Y                   As Integer

    Dim X                   As Integer
    Dim ScreenX             As Integer
    Dim ScreenY             As Integer
    Dim CurrentGrhIndex     As Long
    Dim aCurrentGrhIndex    As Long
    Dim OffX                As Integer
    Dim OffY                As Integer
    Dim TmpInt              As Integer
    Dim tX!, tY!



    Dim i                   As Integer
    Dim jojo                As Byte
    Dim LightOffset         As Long
    Dim ChrID()             As Integer
    Dim ChrY()              As Integer
    Dim PixelOffsetY        As Single
    
    Dim MX!, MY!
    
    Dim MMX%, MMY%
    
    Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, lColorMod)
    
    If MapInfo.UsaAguatierra Then
        For i = 1 To TileLayer(6).NumTiles
            With TileLayer(6).tile(i)
                Engine.Grh_Render_Tileset .ID, .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tilex, .tiley, MapData(.tilex, .tiley).tile_orientation
                If MapData(.tilex, .tiley).is_water Then
                    Grh_Render_Water .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tilex, .tiley

                End If
            End With
        Next i
    Else
        For i = 1 To TileLayer(6).NumTiles
            With TileLayer(6).tile(i)
                Engine.Grh_Render_Tileset .ID, .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tilex, .tiley, MapData(.tilex, .tiley).tile_orientation
            End With
        Next i
    End If
    
    Cachear_Tiles = False
    'batch_render


    For i = 1 To TileLayer(4).NumTiles
        With TileLayer(4).tile(i)
                With MapData(.tilex, .tiley).Graphic(1)
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

                        CurrentGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
                    Else
                        CurrentGrhIndex = 0
                    End If
                End With
                If CurrentGrhIndex > 1 Then
                        Grh_Render_relieve CurrentGrhIndex, _
                            .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, _
                            .tilex, .tiley, MapData(.tilex, .tiley).tile_orientation
                End If
        End With
    Next i
    

    
    For i = 1 To TileLayer(1).NumTiles
        With TileLayer(1).tile(i)
                    Draw_Grh MapData(.tilex, .tiley).Graphic(2), _
                            .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, 1, 1, _
                            .tilex, .tiley, 1

        End With
    Next i

    Sangre_Render

    For i = 1 To TileLayer(2).NumTiles
        With TileLayer(2).tile(i)
            If .type = layer_type.layer_piso_capa3 Then
                Engine.Grh_Render_Tileset .ID, .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tilex, .tiley, MapData(.tilex, .tiley).tile_orientation
                If MapData(.tilex, .tiley).is_water Then
                    Grh_Render_Water .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tilex, .tiley
                End If


            ElseIf .type = layer_type.layer_obj Then

                If MapData(.tilex, .tiley).ObjGrh.GrhIndex <> 0 Then
                    #If EnableParallax = 1 Then
                        GetParalllaxOffset .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, frmMain.agua_profundidad.Value, MX, MY
                        Call Draw_Grh(MapData(.tilex, .tiley).ObjGrh, _
                                    MX + .PixelPosX + offset_screen.X, MY + .PixelPosY + offset_screen.Y, 1, 1, .tilex, .tiley)
                    #Else
                        Call Draw_Grh(MapData(.tilex, .tiley).ObjGrh, _
                                    .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, 1, 1, .tilex, .tiley)
                    #End If

                End If
            ElseIf .type = layer_type.layer_entidades Then
                If EntidadesMap(.tilex, .tiley) <> 0 Then _
                    Call Entidades_Render_Recursivo(.PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, .tilex, .tiley)
            ElseIf .type = layer_type.layer_char Then
                Call Char_Render(.ID)
            ElseIf .type = layer_type.layer_particle Then
                Call Particle_Group_Render(.ID)
            Else
                Call Draw_Grh(MapData(.tilex, .tiley).Graphic(3), _
                            .PixelPosX + offset_screen.X, .PixelPosY + offset_screen.Y, 1, 1, .tilex, .tiley)

            End If
        End With
    Next i


    'render_blood offset_screen.x, offset_screen.Y
    'Sangre.Render

    Projectile_Render

    For i = 1 To TileLayer(3).NumTiles
        With TileLayer(3).tile(i)
            If .type = layer_type.layer_particle Then
                Call Particle_Group_Render(.ID)
            ElseIf .type = layer_type.layer_roof Then
                If techo_alpha Then
                    Call Draw_Grh_Techo(MapData(.tilex, .tiley).Graphic(4), _
                        .PixelPosX + offset_screen.X, _
                        .PixelPosY + offset_screen.Y, 1, techo_alpha)
                End If
            End If

        End With
    Next i

    Call D3DDevice.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_DISABLE)

    'WeatherDoFog = frmMain.VScroll1.value \ 15
    If WeatherDoFog Then Engine_Weather_UpdateFog
    If Sandstorm_do Then Engine_Weather_SandStorm

    If alpha_racio_luz > 0 Then
        Dim cc As Long
        cc = D3DColorARGB(alpha_racio_luz, 0, 0, 0)
        Grh_Render_Simple_box 7535, 16!, -50!, cc, 512!
        Engine.Draw_FilledBox 0, 0, 16, frmMain.Renderer.height, cc, 0, 0
        Engine.Draw_FilledBox 528, 0, 17, frmMain.Renderer.height, cc, 0, 0
    End If

#If esME = 1 Then
    Dim c As Long
    
    If Render_Radio_Luz Then
        c = D3DColorARGB(val(frmMain.rdlo.text) Mod 256, 0, 0, 0)
        Grh_Render_Simple_box 7535, 244!, 0, c, 512!
        Engine.Draw_FilledBox 0, 0, 244, frmMain.renderer.height, c, 0, 0
        Engine.Draw_FilledBox 756, 0, 244, frmMain.renderer.height, c, 0, 0
    End If
    
    If DRAWBLOQUEOS > 0 Or DRAWTRIGGERS > 0 Or dibujarAccionTile > 0 Or DRAWGRILLA > 0 Or dibujarCantidadObjetos Then
        MMX = minl(X_MAXIMO_VISIBLE, screenmaxX)
        MMY = minl(Y_MAXIMO_VISIBLE, screenmaxY)
        For Y = screenminY To MMY
            For X = screenminX To MMX
                    tX = (X + minXOffset) * 32 + offset_map.X
                    tY = (Y + minYOffset) * 32 + offset_map.Y
                    
                    If DRAWBLOQUEOS > 0 And MapData(X, Y).Trigger And &HFF > 0 Then
                        Grh_Render_Bloqueos tX, tY, X, Y
                    End If
                    
                    If MapData(X, Y).Trigger > 0 And DRAWTRIGGERS > 0 Then
                            Grh_Render_Blocked &H7F000000, _
                                tX, tY, X, Y
                    
                            Engine.Text_Render_ext CStr(obtenerDescripcionAbreviatura(MapData(X, Y).Trigger)), tY + 8, tX + 12, 0, 0, &HFFFFFFFF
                    End If
                    
                    '¿Dibujo la información de las acciones?
                    If dibujarAccionTile Then
                        If Not MapData(X, Y).accion Is Nothing Then
                                If MouseTileX = X And MouseTileY = Y Then
                                    Grh_Render_Blocked mzRed And &HEEFFFFFF, tX, tY, X, Y
                                    Engine.Text_Render_ext MapData(X, Y).accion.getNombre, tY, tX + 5, 0, 0, &HFFFFFFFF
                                Else
                                    Grh_Render_Blocked mzRed And &H30FFFFFF, tX, tY, X, Y
                                    Engine.Text_Render_ext MapData(X, Y).accion.getID, tY + 7, tX + 5, 0, 0, &HFFFFFFFF
                                End If
                        End If
                    End If
                    
                    If DRAWGRILLA > 0 Then
                        Grh_Render_Relieve_Tileset_HCD 16038, 51 + (X Mod 4) + 16 * (Y Mod 4), tX, tY, X, Y
                        'Grh_Render_Simple_box 16038, 5, 0, &H44FFFFFF, tileset_tex_h, , tileset_tex_w - tileset_tex_h
                    End If
                    
                    If dibujarCantidadObjetos Then
                        If MapData(X, Y).OBJInfo.OBJIndex > 0 Then
                             Engine.Text_Render_ext CStr(MapData(X, Y).OBJInfo.Amount), tY + 7, tX + 5, 100, 10, &HFFFFFFFF
                        End If
                    End If
            Next X
        Next Y
    End If

    If DRAWCLIENTAREA Then
        ScreenX = 512 - ClientWindowWidth / 2
        Engine.Draw_FilledBox 0, 0, ScreenX + 16, 512, &H7F000000, 0, 0
        Engine.Draw_FilledBox 1024 - ScreenX + 16, 0, ScreenX - 16, 512, &H7F000000, 0, 0
        ScreenY = 256 - ClientWindowHeight / 2
        Engine.Draw_FilledBox ScreenX + 16, 0, ClientWindowWidth, ScreenY + 16, &H7F000000, 0, 0
        Engine.Draw_FilledBox ScreenX + 16, 512 - ScreenY + 16, ClientWindowWidth, ScreenY - 16, &H7F000000, 0, 0
    End If
#End If
End Sub

Public Sub Map_Render()
'cfnc = fnc.E_Map_Render
    Dim Y                   As Integer

    Dim X                   As Integer
    Dim ScreenX             As Integer
    Dim ScreenY             As Integer
    Dim CurrentGrhIndex     As Long
    Dim aCurrentGrhIndex    As Long
    Dim OffX                As Integer
    Dim OffY                As Integer

    Dim TmpInt              As Integer


    Dim i                   As Integer
    Dim jojo                As Byte
    Dim LightOffset         As Long
    Dim ChrID()             As Integer
    Dim ChrY()              As Integer
    Dim PixelOffsetY        As Single
    
    If MapInfo.UsaAguatierra And AnimarAguatierra Then
        GetElapsedTimeME

        kWATER = (kWATER + timerTicksPerFrame * 16) Mod 360
        map_render_kwateR
                
        TiempoAguatierra = GetElapsedTimeME
    Else
        TiempoAguatierra = 0
    End If
    
    Render_Map_Array
    
    If Lightbeam_do Then Engine_Weather_lightbeam
    
    If bTecho Then
        If techo_alpha > 3 Then _
            techo_alpha = techo_alpha - 4
    Else
        If techo_alpha < 251 Then _
            techo_alpha = techo_alpha + 4
    End If

    If meteo_particle <> 0 Then
        Call Particle_Group_Set_PPos(meteo_particle, -offset_map.X + 250, -offset_map.Y)
        Call Particle_Group_Render(meteo_particle)
    End If
	#If esCLIENTE = 1 then
	    If UserMoving = 0 Then If UserDirection Then Moverme UserDirection
	#End if
End Sub

