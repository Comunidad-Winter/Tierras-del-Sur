Attribute VB_Name = "ME_FIFO"
Option Explicit


Private Type POS_BYTE
    X As Byte
    Y As Byte
End Type

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numbytes As Long)

Function AbrirMapa(Path As String) As Boolean
'abre .map
THIS_MAPA.Path = Path
'AbrirMapaCompilado path, 0, 0
 Cargar_Mapa_ME Path
End Function

Function GuardarMapa(Path As String) As Boolean
'guarda .map
THIS_MAPA.editado = False
THIS_MAPA.Path = Path
GuardarMapa = Guardar_Mapa_ME(Path)

End Function

Private Function GenerarBAMapa(ByRef TamanioTotal As Long) As Byte()
    'Buffer de escritura
    Dim ba() As Byte
    
    'Puntero de escritura
    Dim Ptr As Long
    
    'Tamanio del mapa
    Dim TamanioX As Long
    Dim TamanioY As Long
    
    Dim ByFlags As Integer
    
    'Buffers de los costados
    Dim PtrBuffersCostados As Long, ResizeBackBufferX As Integer, ResizeBackBufferY As Integer
    
    Dim PtrDatosTiles As Long
    
    Dim PtrColores As Long
    
    Dim PtrLongitudArchivo As Long
    
    Dim loopc As Integer
    
    Dim lColor  As RGBCOLOR
    Dim lRange  As Byte
    Dim lBrillo As Byte
    Dim lTipo   As Integer
    Dim lID     As Long
    
    TamanioX = MapSize
    TamanioY = MapSize
    
    ReDim ba(1024 * TamanioX * TamanioY) '1KB por tile, despues se trunca
    
    Ptr = VarPtr(ba(0))
    
    MStream.msSetPuntero Ptr
    
    'header de los mapas con formato editor sin comprimir
    msWriteInteger HeadMapaSinComprimir
    
    'Version del formato del mapa
    msWriteInteger 1
    
    'Longitud en bytes del archivo
    PtrLongitudArchivo = msGetCursor
    msWriteLong 0
    
    'Dir del primer byte con datos del tile
    PtrDatosTiles = msGetCursor
    msWriteLong 0
    
    'Dir del primer byte con datos de los colores de las tiles
    PtrColores = msGetCursor
    msWriteLong 0
    
    'Numero del mapa
    msWriteLong THIS_MAPA.numero
    
    'Version del mapa
    msWriteLong THIS_MAPA.Version
    
    'Autor
    msWriteLong THIS_MAPA.Autor
    
    'Nombre del mapa
    msWriteLong LenB(THIS_MAPA.nombre)
    CopyMemory ByVal (msGetCursor + Ptr), StrPtr(THIS_MAPA.nombre), LenB(THIS_MAPA.nombre)
    msSetCursor (msGetCursor + LenB(THIS_MAPA.nombre))

    'Tamanio del mapa
    msWriteLong TamanioX 'Tamanio X
    msWriteLong TamanioY 'Tamanio Y
    
    'Buffer de los costados, se escriben al final del proc.
    PtrBuffersCostados = msGetCursor
    msWriteByte ResizeBackBufferX  'X
    msWriteByte ResizeBackBufferY  'Y


    With MapInfo
    
    msWriteByte .BaseColor.r
    msWriteByte .BaseColor.g
    msWriteByte .BaseColor.b
    
    msWriteByte .ColorPropio
    
    msWriteInteger .agua_tileset
    
    msWriteInteger .agua_rect.top
    msWriteInteger .agua_rect.Bottom
    msWriteInteger .agua_rect.Left
    msWriteInteger .agua_rect.Right
    
    msWriteInteger .agua_profundidad
    
    msWriteByte .puede_nieve
    msWriteByte .puede_lluvia
    msWriteByte .puede_neblina
    msWriteByte .puede_niebla
    msWriteByte .puede_sandstorm
    msWriteByte .puede_nublado
    
    msWriteByte .UsaAguatierra
    
    msWriteInteger .SonidoLoop
    
    msWriteByte .MinNivel
    msWriteByte .MaxNivel

    msWriteInteger .UsuariosMaximo
    
    msWriteByte .SeCaenItems
    msWriteByte .PermiteRoboNpc
    msWriteByte .PermiteHechizosPetes
    msWriteByte .MagiaSinEfecto
    msWriteByte .MapaPK
    
    '############################### TILES ####################################
    
    WriteLong PtrDatosTiles + Ptr, msGetCursor
    
    Dim X As Integer, Y As Integer
    
    For Y = 1 To TamanioX
        For X = 1 To TamanioY
            With MapData(X, Y)
                ByFlags = 0
                
                If .Blocked = 1 Then _
                    ByFlags = ByFlags Or bitwisetable(0)
                
                If .Graphic(1).GrhIndex Then ByFlags = ByFlags Or bitwisetable(1)
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or bitwisetable(2)
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or bitwisetable(3)
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or bitwisetable(4)
                
                If .Trigger Then _
                    ByFlags = ByFlags Or bitwisetable(5)
                    
                If .tile_texture Then _
                    ByFlags = ByFlags Or bitwisetable(6)
                
                If .Particles_groups_original(0) Or .Particles_groups_original(1) Or .Particles_groups_original(2) Then _
                    ByFlags = ByFlags Or bitwisetable(7)
                
                If .is_water Then _
                    ByFlags = ByFlags Or bitwisetable(8)

                If Alturas(X, Y) Or AlturaPie(X, Y) Or hMapData(X, Y).hs(0) Then _
                    ByFlags = ByFlags Or bitwisetable(9)
                
                If .TileExit.map Then _
                    ByFlags = ByFlags Or bitwisetable(10)
                If .NPCIndex Then _
                    ByFlags = ByFlags Or bitwisetable(11)
                    
                If .OBJInfo.OBJIndex Then _
                    ByFlags = ByFlags Or bitwisetable(12)
                    
                If .luz Then
                    If DLL_Luces.Check(.luz) Then
                        ByFlags = ByFlags Or bitwisetable(13)
                    End If
                End If
                
                If .tile_orientation Then ByFlags = ByFlags Or bitwisetable(14)
                
                msWriteInteger ByFlags
                
                'Flags (bits):
                '   0 = Bloqueado
                '   1 = Grh1
                '   2 = Grh2
                '   3 = Grh3
                '   4 = Grh4
                '   5 = trigger
                '   6 = Tileset
                '   7 = particulas
                '   8 = EsAgua
                '   9 = TIene altura
                '   10= Tiene translado
                '   11= TIene NPC
                '   12= TIene OBJ
                '   13= Tiene luz
                '   14= Se dibuja al revés la tile (tile_orientation)
                '   15= Tiene Eventos, disparadores, etc
                
                If .tile_texture Then
                    msWriteInteger .tile_texture
                    msWriteInteger .tile_number
                End If
                
                For loopc = 1 To 4
                    If .Graphic(loopc).GrhIndex Then
                        msWriteLong .Graphic(loopc).GrhIndex
                    End If
                Next loopc
                
                If .Graphic(3).GrhIndex Then
                    'On Local Error Resume Next
                        If ResizeBackBufferY < GrhData(.Graphic(3).GrhIndex).pixelHeight Then ResizeBackBufferY = GrhData(.Graphic(3).GrhIndex).pixelHeight
                        If ResizeBackBufferX < GrhData(.Graphic(3).GrhIndex).pixelWidth Then ResizeBackBufferX = GrhData(.Graphic(3).GrhIndex).pixelWidth
                    'On Local Error GoTo ErrorSave
                End If
                
                If .Trigger Then
                    msWriteInteger .Trigger
                End If
                
                If ByFlags And bitwisetable(7) Then
                    msWriteInteger .Particles_groups_original(0)
                    msWriteInteger .Particles_groups_original(1)
                    msWriteInteger .Particles_groups_original(2)
                End If
                
                If ByFlags And bitwisetable(9) Then
                    msWriteFloat NormalData(X, Y).X     'Normal del vértice
                    msWriteFloat NormalData(X, Y).Y     'Normal del vértice
                    msWriteFloat NormalData(X, Y).z     'Normal del vértice
                    
                    msWriteInteger Alturas(X, Y)         'Alturas
                    msWriteInteger AlturaPie(X, Y)       'Altura Pie
                    
                    msWriteInteger Round(hMapData(X, Y).hs(0)) 'Altura del vértice
                End If
                
                If MapData(X, Y).TileExit.map Then
                    msWriteInteger .TileExit.map
                    msWriteByte .TileExit.X
                    msWriteByte .TileExit.Y
                End If
                
                If MapData(X, Y).NPCIndex Then
                    msWriteInteger .NPCIndex
                End If
                
                If .OBJInfo.OBJIndex Then
                    msWriteInteger .OBJInfo.OBJIndex
                    msWriteInteger .OBJInfo.Amount
                End If

                If ByFlags And bitwisetable(13) Then
                    DLL_Luces.Get_Light .luz, 0, 1, lColor.r, lColor.g, lColor.b, lRange, lBrillo, lID, lTipo
                    msWriteByte lColor.r
                    msWriteByte lColor.g
                    msWriteByte lColor.b
                    
                    msWriteByte lRange
                    msWriteLong lID
                    msWriteByte lBrillo
                    msWriteInteger lTipo
                End If
            End With
        Next X
    Next Y
    
    'Guardo en el header los buffers

    
    WriteInteger PtrBuffersCostados + Ptr, ResizeBackBufferX
    WriteInteger PtrBuffersCostados + Ptr + 2, ResizeBackBufferY
    
    WriteLong PtrColores + Ptr, msGetCursor
    
    MsgBox "PUNTERO COLORES (SAVE): " & msGetCursor

    CopyMemory ByVal (Ptr + msGetCursor), OriginalMapColor(1, 1), 4 * MapSizeS 'Este no cambia porque es un array estático que es de acceso más rápido en la ram porque tiene menos direccionamientos
    CopyMemory ByVal (Ptr + msGetCursor + 4 * MapSizeS), Intensidad_Del_Terreno(1, 1), MapSizeS
    
    msSetCursor msGetCursor + 5 * MapSizeS

    
    End With
    
    
'###############################################################################################################################
    
    Dim TamanioArchivo As Long
    TamanioArchivo = msGetCursor ' (Ultimo puntero de escritura - Puntero inicio)
    
    WriteLong PtrLongitudArchivo + Ptr, TamanioArchivo 'Guardo el tamaño en bytes del archivo.
    
    ReDim Preserve ba(TamanioArchivo)
'
'#If Comprimir_Mapas = 1 Then
'    Dim TamanioArchivoComprimido As Long
'    Dim BaOut() As Byte
'
'    Compress_Data ba
'
'    TamanioArchivoComprimido = UBound(ba) + 1
'
'    ReDim BaOut(TamanioArchivoComprimido + 10)
'
'    Ptr = VarPtr(BaOut(0))
'    msWriteInteger HeadMapaComprimido
'    msWritelong TamanioArchivoComprimido
'    msWritelong TamanioArchivo
'
'    CopyMemory ByVal Ptr, ba(0), TamanioArchivoComprimido
'
'    GenerarBAMapa = BaOut
'#Else
    GenerarBAMapa = ba
'#End If

End Function



Function Guardar_Mapa_ME(ByVal SaveAs As String) As Boolean
    On Error GoTo ErrorSave
        
        Dim FreeFileMap As Long
        
        Dim ba() As Byte
        
        ba = GenerarBAMapa(0)
        
        If UBound(ba) > 0 Then
            If FileExist(SaveAs, vbNormal) = True Then
                Kill SaveAs
            End If
        
            FreeFileMap = FreeFile
        
            Open SaveAs For Binary As FreeFileMap
                Seek FreeFileMap, 1
                Put FreeFileMap, , ba
            Close FreeFileMap
            
            Guardar_Mapa_ME = True
        End If
    Exit Function
    
ErrorSave:
        MsgBox "Error en Guardar mapa ME, nro. " & Err.number & " - " & Err.Description
End Function

Public Sub CompilarMapa()

End Sub

Public Sub Cargar_Mapa_ME(ByVal FILE As String, Optional Offset As Long = 1, Optional Tamanio As Long)
    Dim Y As Integer
    Dim X As Integer
    Dim FreeFileMap As Long
    
    Dim Ptr As Long
    Dim ba() As Byte
    
    
    'Tamanio del mapa
    Dim TamanioX As Long
    Dim TamanioY As Long
    
    Dim ByFlags As Integer
    
    'Buffers de los costados
    Dim PtrBuffersCostados As Long, ResizeBackBufferX As Integer, ResizeBackBufferY As Integer
    
    Dim PtrDatosTiles As Long
    
    Dim PtrColores As Long
    
    Dim PtrLongitudArchivo As Long
    
    Dim tmpLong As Long
    
    Dim lColor  As RGBCOLOR
    Dim lRange  As Byte
    Dim lBrillo As Byte
    Dim lTipo   As Integer
    Dim lID     As Long
    
    LIMPIAR_MAPA
    
    'Obtengo el tamanio si no me lo pasaron
    If Tamanio = 0 Then Tamanio = FileLen(FILE)
    ReDim ba(Tamanio - 1)
    
    'Obtengo el mapa
    FreeFileMap = FreeFile
    Open FILE For Binary As FreeFileMap
        Seek FreeFileMap, Offset
        Get FreeFileMap, , ba
    Close FreeFileMap
    
    Ptr = VarPtr(ba(0))
    
    msSetPuntero Ptr
    
    If (msReadInteger <> HeadMapaSinComprimir) Then Exit Sub
    
    msReadInteger 'version del formato del archivo
    
    PtrLongitudArchivo = msReadLong
    PtrDatosTiles = msReadLong
    PtrColores = msReadLong
    
    THIS_MAPA.numero = msReadLong
    THIS_MAPA.Version = msReadLong
    
    THIS_MAPA.Autor = msReadLong
    
    tmpLong = msReadLong
    THIS_MAPA.nombre = Space$(tmpLong)
    CopyMemory StrPtr(THIS_MAPA.nombre), ByVal Ptr, tmpLong
    msSetCursor msGetCursor + tmpLong
    
    TamanioX = msReadLong
    TamanioY = msReadLong
    
    ResizeBackBufferX = msReadByte
    ResizeBackBufferY = msReadByte
    
    With MapInfo
        .BaseColor.r = msReadByte
        .BaseColor.g = msReadByte
        .BaseColor.b = msReadByte
        .ColorPropio = msReadByte <> 0
        
        .agua_tileset = msReadInteger
        
        .agua_rect.top = msReadInteger
        .agua_rect.Bottom = msReadInteger
        .agua_rect.Left = msReadInteger
        .agua_rect.Right = msReadInteger
        
        .agua_profundidad = msReadInteger
        
        .puede_nieve = msReadByte <> 0
        .puede_lluvia = msReadByte <> 0
        .puede_neblina = msReadByte <> 0
        .puede_niebla = msReadByte <> 0
        .puede_sandstorm = msReadByte <> 0
        .puede_nublado = msReadByte <> 0
        
        .UsaAguatierra = msReadByte <> 0
        
        .SonidoLoop = msReadInteger
        
        .MinNivel = msReadByte
        .MaxNivel = msReadByte
        
        .UsuariosMaximo = msReadInteger
        
        .SeCaenItems = msReadByte <> 0
        .PermiteRoboNpc = msReadByte <> 0
        .PermiteHechizosPetes = msReadByte <> 0
        .MagiaSinEfecto = msReadByte <> 0
        .MapaPK = msReadByte <> 0
        
        'Termine con las caracteristicas, ahora vamos a las tiles.
        
        msSetCursor PtrDatosTiles
        
        For X = 1 To TamanioX
            For Y = 1 To TamanioY
                ByFlags = msReadInteger
                
                With MapData(X, Y)
                                
                    .Blocked = (ByFlags And bitwisetable(0))
                    .is_water = (ByFlags And bitwisetable(8))
                    
                    If ByFlags And bitwisetable(6) Then
                        MapData(X, Y).tile_texture = msReadInteger
                        MapData(X, Y).tile_number = msReadInteger
                    Else
                        MapData(X, Y).tile_texture = 0
                    End If
                    
                    If ByFlags And bitwisetable(1) Then
                        InitGrh .Graphic(1), msReadInteger
                    Else
                        .Graphic(1).GrhIndex = 0
                    End If
                        
                    If ByFlags And bitwisetable(2) Then
                        InitGrh .Graphic(2), msReadInteger
                    Else
                        .Graphic(2).GrhIndex = 0
                    End If
                        
                    'Layer 3 used?
                    If ByFlags And bitwisetable(3) Then
                        InitGrh .Graphic(3), msReadInteger
                    Else
                        .Graphic(3).GrhIndex = 0
                    End If
                        
                    'Layer 4 used?
                    If ByFlags And bitwisetable(4) Then
                        InitGrh .Graphic(4), msReadInteger
                    Else
                        .Graphic(4).GrhIndex = 0
                    End If
        
                    'Trigger used?
                    If ByFlags And bitwisetable(5) Then
                        MapData(X, Y).Trigger = msReadInteger
                    Else
                        MapData(X, Y).Trigger = 0
                    End If
                    
                    If ByFlags And bitwisetable(7) Then
                        MapData(X, Y).Particles_groups_original(0) = msReadInteger
                        MapData(X, Y).Particles_groups_original(1) = msReadInteger
                        MapData(X, Y).Particles_groups_original(2) = msReadInteger
                        
                        If MapData(X, Y).Particles_groups_original(0) Then _
                            Engine_Particles.Particle_Group_Make 0, X, Y, MapData(X, Y).Particles_groups_original(0), 0
                        If MapData(X, Y).Particles_groups_original(1) Then _
                            Engine_Particles.Particle_Group_Make 0, X, Y, MapData(X, Y).Particles_groups_original(1), 1
                        If MapData(X, Y).Particles_groups_original(2) Then _
                            Engine_Particles.Particle_Group_Make 0, X, Y, MapData(X, Y).Particles_groups_original(2), 2
                    Else
                        MapData(X, Y).Particles_groups_original(0) = 0
                        MapData(X, Y).Particles_groups_original(1) = 0
                        MapData(X, Y).Particles_groups_original(2) = 0
                    End If
                    
                    If ByFlags And bitwisetable(9) Then
                        MapData(X, Y).tile_orientation = ByFlags And bitwisetable(14)
                        
                        NormalData(X, Y).X = msReadFloat
                        NormalData(X, Y).Y = msReadFloat
                        NormalData(X, Y).z = msReadFloat
                        
                        Alturas(X, Y) = msReadInteger
                        AlturaPie(X, Y) = msReadInteger
                        hMapData(X, Y).h = Alturas(X, Y)
                        
                        hMapData(X, Y).hs(0) = msReadInteger
                    End If
                    
                    If ByFlags And bitwisetable(10) Then
                        .TileExit.map = msReadInteger
                        .TileExit.X = msReadByte
                        .TileExit.Y = msReadByte
                    End If
            
                    If ByFlags And bitwisetable(11) Then
                        'Get and make NPC
                        .NPCIndex = msReadInteger
            
                        If .NPCIndex < 0 Then
                            .NPCIndex = 0
                        Else
                            Call MakeChar(NextOpenChar(), NpcData(.NPCIndex).Body, NpcData(.NPCIndex).Head, NpcData(.NPCIndex).Heading, X, Y, 0, 0, 0)
                        End If
                    End If
            
                    If ByFlags And bitwisetable(12) Then
                        'Get and make Object
                        .OBJInfo.OBJIndex = msReadInteger
                        .OBJInfo.Amount = msReadInteger
                        
                        If .OBJInfo.OBJIndex > 0 Then
                            InitGrh .ObjGrh, ObjData(.OBJInfo.OBJIndex).GrhIndex
                        End If
                    End If
                    
                    If ByFlags And bitwisetable(13) Then
                        lColor.r = msReadByte
                        lColor.g = msReadByte
                        lColor.b = msReadByte
                        
                        lRange = msReadByte
                        lID = msReadLong
                        
                        lBrillo = msReadByte
                        
                        lTipo = msReadInteger
                        
                        .luz = _
                            DLL_Luces.Crear(X, Y, lColor.r, lColor.g, lColor.b, lRange, lBrillo, lID, lTipo)
                    Else
                        .luz = 0
                    End If
                End With
            Next Y
        Next X
        
        
        MsgBox "Sobra:" & (UBound(ba) - PtrColores - 5 * MapSizeS)
        DXCopyMemory OriginalMapColor(1, 1), ByVal (VarPtr(ba(0)) + PtrColores), 4 * MapSizeS 'Este no cambia porque es un array estático que es de acceso más rápido en la ram porque tiene menos direccionamientos
        DXCopyMemory Intensidad_Del_Terreno(1, 1), ByVal (PtrColores + VarPtr(ba(0)) + 4 * MapSizeS), MapSizeS
        
    End With
    
    Dim XX&, YY&
    
    For Y = 2 To TamanioX - 1
        For X = 2 To TamanioY - 1
            'If GrhData(GrhData(MapData(X, Y).Graphic(1).GrhIndex).Frames(1)).FileNum Then
            '    PreLoadTexture GrhData(GrhData(MapData(X, Y).Graphic(1).GrhIndex).Frames(1)).FileNum
            'End If
            hMapData(X, Y + 1).hs(1) = hMapData(X, Y).hs(0)
            hMapData(X - 1, Y + 1).hs(3) = hMapData(X, Y).hs(0)
            hMapData(X - 1, Y).hs(2) = hMapData(X, Y).hs(0)
            
            If Not MapData(X, Y).is_water Then
                ModSuperWaterMM(X, Y + 1).hs(3) = 0
                ModSuperWaterMM(X, Y + 1).hs(1) = 0

                ModSuperWaterMM(X, Y - 1).hs(2) = 0
                ModSuperWaterMM(X, Y - 1).hs(0) = 0

                ModSuperWaterMM(X + 1, Y).hs(0) = 0
                ModSuperWaterMM(X + 1, Y).hs(1) = 0

                ModSuperWaterMM(X - 1, Y).hs(2) = 0
                ModSuperWaterMM(X - 1, Y).hs(3) = 0

                ModSuperWaterMM(X + 1, Y - 1).hs(0) = 0
                ModSuperWaterMM(X + 1, Y + 1).hs(1) = 0
                ModSuperWaterMM(X - 1, Y - 1).hs(2) = 0
                ModSuperWaterMM(X - 1, Y + 1).hs(3) = 0
            End If

            ModSuperWaterMM(X, Y + 1).hs(1) = ModSuperWaterMM(X, Y).hs(0)
            ModSuperWaterMM(X - 1, Y + 1).hs(3) = ModSuperWaterMM(X, Y).hs(0)
            ModSuperWaterMM(X - 1, Y).hs(2) = ModSuperWaterMM(X, Y).hs(0)

            ModSuperWaterDD(X, Y).hs(0) = 0
            ModSuperWaterDD(X, Y + 1).hs(1) = 0
            ModSuperWaterDD(X - 1, Y + 1).hs(3) = 0
            ModSuperWaterDD(X - 1, Y).hs(2) = 0
        Next X
    Next Y

    MapInfo.MaxGrhSizeXInTiles = ResizeBackBufferX \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
    MapInfo.MaxGrhSizeYInTiles = ResizeBackBufferY \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
  
    Engine_Set_TileBuffer_Size MapInfo.MaxGrhSizeXInTiles, MapInfo.MaxGrhSizeYInTiles

    Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), MapSizeS * 4)

    cron_tiempo
    Light_Update_Map = True
    Light_Update_Sombras = True
        
    frmMain.MousePointer = 0
    
    THIS_MAPA.editado = False
End Sub



Public Sub LIMPIAR_MAPA()
    Dim Y As Long
    Dim X As Long

    DLL_Luces.Remove_All
    Engine_Particles.Particle_Group_Remove_All
    FX_Projectile_Erase_All
    FX_Rayo_Erase_All
    ZeroMemory Alturas(1, 1), MapSizeS * 2
    ZeroMemory ModSuperWater(1, 1), MapSizeS
    ZeroMemory ModSuperWaterMM(1, 1), LenB(ModSuperWaterMM(1, 1)) * MapSizeS
    ZeroMemory ModSuperWaterDD(1, 1), LenB(ModSuperWaterDD(1, 1)) * MapSizeS
    ZeroMemory hMapData(1, 1), LenB(hMapData(1, 1)) * MapSizeS
    ZeroMemory AlturaPie(1, 1), 2 * MapSizeS
    
    Clear_Luces_Mapa
    
    'Reiniciamos el mapa.

    For Y = 1 To MapSize
        For X = 1 To MapSize
            With MapData(X, Y)
                .Blocked = 0
                .is_water = 0
                .Graphic(1).GrhIndex = 0
                .Graphic(2).GrhIndex = 0
                .Graphic(3).GrhIndex = 0
                .Graphic(4).GrhIndex = 0
                .ObjGrh.GrhIndex = 0
                .Trigger = 0
                .tile_texture = 0
                .TileExit.map = 0
                .NPCIndex = 0
                .OBJInfo.OBJIndex = 0

                .Particles_groups_original(0) = 0
                .Particles_groups_original(1) = 0
                .Particles_groups_original(2) = 0

            End With
        Next X
    Next Y
End Sub

Sub NuevoMapa()
    LIMPIAR_MAPA
    
    DLL_Luces.Remove_All
    Engine_Particles.Particle_Group_Remove_All
    FX_Projectile_Erase_All
    FX_Rayo_Erase_All
    
    THIS_MAPA.editado = False
    THIS_MAPA.Path = ""
    THIS_MAPA.numero = 0
    THIS_MAPA.nombre = "Mapa sin nombre"
    
    RemakeWaterTilenumbers 0, 0, 3, 3
    
    MapInfo.UsaAguatierra = False
    
    MapInfo.agua_tileset = 19
End Sub
