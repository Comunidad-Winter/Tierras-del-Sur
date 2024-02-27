Attribute VB_Name = "modPersistencia"
Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Sub Cargar_Mapa_CLI2(ByVal archivo As String, Optional Offset As Long = 1)
    Dim loopC As Long
    Dim Y As Long
    Dim X As Long
    Dim TempInt As Integer
    Dim ByFlags As Byte
    
    Dim nombreArchivo As String
    Dim handleArchivo As Integer
    
    Dim trigger As Integer
              
    handleArchivo = FreeFile()
    
    Open archivo For Binary As handleArchivo
    Seek handleArchivo, Offset

    'map Header
    Get handleArchivo, , mapinfo.MapVersion
    Get handleArchivo, , MiCabecera
    Get handleArchivo, , TempInt
    Get handleArchivo, , TempInt
    Get handleArchivo, , TempInt
    Get handleArchivo, , TempInt
    
    
    mapinfo.BaseColor.r = 255
    mapinfo.BaseColor.g = 255
    mapinfo.BaseColor.b = 255
    'Load arrays
    For Y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For X = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        
            Get handleArchivo, , ByFlags

            
            mapdata(X, Y).trigger = 0
              
            If (ByFlags And 1) Then
                 modTriggers.BloquearTile X, Y
            End If

            Get handleArchivo, , mapdata(X, Y).Graphic(1).GrhIndex
            
            InitGrh mapdata(X, Y).Graphic(1), mapdata(X, Y).Graphic(1).GrhIndex

            'Layer 2 used?
            If ByFlags And 2 Then
                Get handleArchivo, , mapdata(X, Y).Graphic(2).GrhIndex
                InitGrh mapdata(X, Y).Graphic(2), mapdata(X, Y).Graphic(2).GrhIndex
            Else
                mapdata(X, Y).Graphic(2).GrhIndex = 0
            End If

            'Layer 3 used?
            If ByFlags And 4 Then
                Get handleArchivo, , mapdata(X, Y).Graphic(3).GrhIndex
                InitGrh mapdata(X, Y).Graphic(3), mapdata(X, Y).Graphic(3).GrhIndex
            Else
                mapdata(X, Y).Graphic(3).GrhIndex = 0
            End If

            'Layer 4 used?
            If ByFlags And 8 Then
                Get handleArchivo, , mapdata(X, Y).Graphic(4).GrhIndex
                InitGrh mapdata(X, Y).Graphic(4), mapdata(X, Y).Graphic(4).GrhIndex
            Else
                mapdata(X, Y).Graphic(4).GrhIndex = 0
            End If
           
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handleArchivo, , trigger
                
                If trigger = 1 Or trigger = 2 Or trigger = 4 Or trigger = 8 Then
                    mapdata(X, Y).trigger = (mapdata(X, Y).trigger Or eTriggers.BajoTecho)
                End If
            End If

            If (mapdata(X, Y).Graphic(3).GrhIndex >= 7000 And mapdata(X, Y).Graphic(3).GrhIndex <= 7008) Or mapdata(X, Y).Graphic(3).GrhIndex = 648 Or mapdata(X, Y).Graphic(3).GrhIndex = 645 Then
                mapdata(X, Y).trigger = (mapdata(X, Y).trigger Or eTriggers.Transparentar)
            End If
            
            If HayAgua(mapdata(X, Y)) Then
                mapdata(X, Y).trigger = (mapdata(X, Y).trigger Or eTriggers.Navegable)
                mapdata(X, Y).trigger = (mapdata(X, Y).trigger Or eTriggers.NoCaminable)
            End If

            'Erase OBJs
            mapdata(X, Y).ObjGrh.GrhIndex = 0

            mapdata(X, Y).EfectoPisada = 1
            
           'Dim i As Integer
           'For loopC = 1 To 4
           
            'Set up GRH
            'If mapdata(x, y).Graphic(loopC).GrhIndex > 0 Then
                ' EL YIND - Si cargamo un mapa cargamo sus graficos asi no se traba al caminar
                ' For i = 1 To GrhData(mapdata(x, y).Graphic(loopC).GrhIndex).NumFrames
                '    Call CargarSurface(GrhData(GrhData(mapdata(x, y).Graphic(loopC).GrhIndex).frames(i)).filenum)
                ' Next i
            'End If
            'Next loopC
        Next X
    Next Y

    Close handleArchivo

    ' Calcuos dinamicos
    Engine_Montana.ActualizarArraysAlturasMapas
    
    mapinfo.MaxGrhSizeXInTiles = 512 \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
    mapinfo.MaxGrhSizeYInTiles = 512 \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
  
    Engine_Set_TileBuffer_Size mapinfo.MaxGrhSizeXInTiles, mapinfo.MaxGrhSizeYInTiles

     Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)
    
    Engine_Montana.Compute_Mountain
    
    cron_tiempo
    Light_Update_Map = True
    Light_Update_Sombras = True
    
    CurMap = 1

End Sub

Public Sub Cargar_Mapa_CLI(ByVal archivo As String, Optional Offset As Long = 1)

    Dim TempByte As Byte
    Dim body As Integer
    Dim Head As Integer
    Dim heading As Byte
    Dim Y As Integer
    Dim X As Integer
    Dim ByFlags As Integer

    Dim MapaNumeroOriginal As Integer
    Dim ResizeBackBufferX As Integer
    Dim ResizeBackBufferY As Integer
    
    Dim plusa As Integer
    
    Dim tempLuz As tLuzPropiedades
    Dim tempintMarce As Integer
    Dim Char As Integer
    
    Dim FreeFileMap As Integer
    
    FreeFileMap = FreeFile
    
    ' Abro el archivo donde esta el mapa y me posiciono donde comienza la informacion
    Open archivo For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    '###################################################################################
    '########### HEADER
    Dim header   As String * 16
    Get FreeFileMap, , header
    Get FreeFileMap, , MapaNumeroOriginal
         
    '###################################################################################
    '########### PROPIEDADES
    
    Get FreeFileMap, , ResizeBackBufferX
    Get FreeFileMap, , ResizeBackBufferY

    '###################################################################################
    '########### COLORES Y LUCES!
    
    Get FreeFileMap, , OriginalMapColor
    Get FreeFileMap, , Intensidad_Del_Terreno

    With mapinfo
        
        Get FreeFileMap, , .BaseColor
        Get FreeFileMap, , .ColorPropio
        
        Get FreeFileMap, , .agua_tileset
        Get FreeFileMap, , .agua_rect
        Get FreeFileMap, , .agua_profundidad
                
        Get FreeFileMap, , .UsaAguatierra
        
    End With

    
    For Y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For X = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
    
            Get FreeFileMap, , ByFlags
            
            ' 6) TileSet
            If ByFlags And bitwisetable(6) Then
                Get FreeFileMap, , mapdata(X, Y).tile_texture
                Get FreeFileMap, , mapdata(X, Y).tile_number
            Else
                mapdata(X, Y).tile_texture = 0
            End If
            
            '1) Capa 1
            If ByFlags And bitwisetable(1) Then
                Get FreeFileMap, , mapdata(X, Y).Graphic(1).GrhIndex
                InitGrh mapdata(X, Y).Graphic(1), mapdata(X, Y).Graphic(1).GrhIndex
            Else
                mapdata(X, Y).Graphic(1).GrhIndex = 0
            End If
            
            '2) Capa 2
            If ByFlags And bitwisetable(2) Then
                Get FreeFileMap, , mapdata(X, Y).Graphic(2).GrhIndex
                InitGrh mapdata(X, Y).Graphic(2), mapdata(X, Y).Graphic(2).GrhIndex
            Else
                mapdata(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            '3) Capa 3
            If ByFlags And bitwisetable(3) Then
                Get FreeFileMap, , mapdata(X, Y).Graphic(3).GrhIndex
                InitGrh mapdata(X, Y).Graphic(3), mapdata(X, Y).Graphic(3).GrhIndex
            Else
                mapdata(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            '4) Capa 4
            If ByFlags And bitwisetable(4) Then
                Get FreeFileMap, , mapdata(X, Y).Graphic(4).GrhIndex
                InitGrh mapdata(X, Y).Graphic(4), mapdata(X, Y).Graphic(4).GrhIndex
            Else
                mapdata(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            '0) Capa 5
            If ByFlags And bitwisetable(0) Then
                Get FreeFileMap, , mapdata(X, Y).Graphic(5).GrhIndex
                InitGrh mapdata(X, Y).Graphic(5), mapdata(X, Y).Graphic(5).GrhIndex
            Else
                mapdata(X, Y).Graphic(5).GrhIndex = 0
            End If

            '5) Trigger
            If ByFlags And bitwisetable(5) Then
                Get FreeFileMap, , mapdata(X, Y).trigger
            Else
                mapdata(X, Y).trigger = 0
            End If
            
            '9) Altura
            If ByFlags And bitwisetable(9) Then
                    Get FreeFileMap, , NormalData(X, Y)
                    Get FreeFileMap, , Alturas(X, Y)
                    Get FreeFileMap, , AlturaPie(X, Y)
                    Get FreeFileMap, , plusa
                    Get FreeFileMap, , mapdata(X, Y).tile_orientation
                    hMapData(X, Y).h = Alturas(X, Y)
                    hMapData(X, Y).hs(0) = plusa
            End If
        
            
            '13) Luz
            If ByFlags And bitwisetable(13) Then

                Get FreeFileMap, , tempLuz.LuzColor.b
                Get FreeFileMap, , tempLuz.LuzColor.g
                Get FreeFileMap, , tempLuz.LuzColor.r
                Get FreeFileMap, , tempLuz.LuzRadio
                Get FreeFileMap, , tempLuz.LuzBrillo
                Get FreeFileMap, , tempLuz.LuzTipo
                Get FreeFileMap, , tempLuz.luzInicio
                Get FreeFileMap, , tempLuz.luzFin
                
                If Engine_Light_Helper.EsLuzValida(tempLuz.LuzRadio, tempLuz.LuzBrillo, tempLuz.LuzTipo) Then
                    mapdata(X, Y).luz = DLL_Luces.crear(X, Y, tempLuz.LuzColor.r, tempLuz.LuzColor.g, tempLuz.LuzColor.b, tempLuz.LuzRadio, tempLuz.LuzBrillo, tempLuz.LuzTipo, tempLuz.luzInicio, tempLuz.luzFin)
                End If
                
            Else
                mapdata(X, Y).luz = 0
            End If
            
            '9)  Particula 0
            If ByFlags And bitwisetable(7) Then
                Set mapdata(X, Y).Particles_groups(0) = New Engine_Particle_Group
                
                If mapdata(X, Y).Particles_groups(0).Cargar(FreeFileMap) Then
                    mapdata(X, Y).Particles_groups(0).SetPos X, Y
                Else
                    Set mapdata(X, Y).Particles_groups(0) = Nothing
                End If
            End If
            
            '10) Particula 1
            If ByFlags And bitwisetable(14) Then
                Set mapdata(X, Y).Particles_groups(1) = New Engine_Particle_Group
                
                If mapdata(X, Y).Particles_groups(1).Cargar(FreeFileMap) Then
                    mapdata(X, Y).Particles_groups(1).SetPos X, Y
                Else
                    Set mapdata(X, Y).Particles_groups(1) = Nothing
                End If
            End If
            
            '8) Particula 2
            If ByFlags And bitwisetable(8) Then
                Set mapdata(X, Y).Particles_groups(2) = New Engine_Particle_Group
                
                If mapdata(X, Y).Particles_groups(2).Cargar(FreeFileMap) Then
                    mapdata(X, Y).Particles_groups(2).SetPos X, Y
                Else
                    Set mapdata(X, Y).Particles_groups(2) = Nothing
                End If
            End If
            
            If ByFlags And bitwisetable(10) Then
                Get FreeFileMap, , mapdata(X, Y).EfectoPisada
            End If
            
           ' If (mapdata(X, Y).trigger And eTriggers.Navegable) Then
                mapdata(X, Y).is_water = 1
            'End If

        Next X
    Next Y
    
    ' Calcuos dinamicos
    Engine_Montana.ActualizarArraysAlturasMapas
    
    mapinfo.MaxGrhSizeXInTiles = ResizeBackBufferX \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
    mapinfo.MaxGrhSizeYInTiles = ResizeBackBufferY \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
  
    Engine_Set_TileBuffer_Size mapinfo.MaxGrhSizeXInTiles, mapinfo.MaxGrhSizeYInTiles

    Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)
    
    Engine_Montana.Compute_Mountain
    
    cron_tiempo
    Light_Update_Map = True
    Light_Update_Sombras = True

End Sub

