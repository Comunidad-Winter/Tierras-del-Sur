Attribute VB_Name = "ME_CopiadoBordes"
Option Explicit

Private Type LightDescriptor
    tieneLuz    As Boolean
    lColor      As BGRACOLOR_DLL
    LRange      As Byte
    LBrillo     As Byte
    LTipo       As Integer
    luzInicio   As Byte
    luzFin      As Byte
End Type

Private mapaCopiaTile(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)           As MapBlock
Private mapaCopiaLuninosidad(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)    As Byte         'Guarda la intensidad de la luz de un vertice del mapa
Private mapaCopiaColor(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)          As BGRACOLOR_DLL 'Colores precalculados en el mapeditor
Private mapaCopiahMapData(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)       As AUDT
Private mapaCopiaAlturaPie(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)      As Integer
Private mapaCopiaAlturas(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)        As Integer
Private mapaCopiaNormalData(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)     As D3DVECTOR

Private mapaCopiaLuces(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE)          As LightDescriptor


    
Public Function CopiarBordesMapaActual() As Boolean
    Dim MapaN As Integer, MapaS As Integer, MapaE As Integer, MapaO As Integer
    Dim MapaNE As Integer, MapaNO As Integer, MapaSE As Integer, MapaSO As Integer

    Dim numeroMapa As Integer

    'Guardo el mapa actual y guardo el numero en variable
    If Not frmMain.GuardarMapaActual() Then
        MsgBox "No se pudieron copiar los bordes", vbCritical
        Exit Function
    End If
    
    numeroMapa = THIS_MAPA.Numero

    'Buscar mapas perifericos
    MapaN = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, ePuntoCardinal.NORTE)
    MapaS = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, ePuntoCardinal.SUR)
    MapaE = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, ePuntoCardinal.ESTE)
    MapaO = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, ePuntoCardinal.OESTE)
    
    MapaNE = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, ePuntoCardinal.NORESTE)
    MapaNO = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, ePuntoCardinal.NOROESTE)
    MapaSE = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, ePuntoCardinal.SURESTE)
    MapaSO = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, ePuntoCardinal.SUROESTE)
    
    Debug.Print MapaNO; MapaN; MapaNE
    Debug.Print MapaO; numeroMapa; MapaE
    Debug.Print MapaSO; MapaS; MapaSE
    
    'Abro mapas y copio sus respectivos pedazos a una estructura temporal

    'NORTE
    If MapaN > 0 Then
        If pakMapasME.Cabezal_GetFileSize(MapaN) Then
            SwitchMap MapaN
            PegarAreaDesdeHasta X_MINIMO_USABLE, Y_MAXIMO_NO_VISIBLE_OTRO_MAPA + 1, X_MAXIMO_USABLE, Y_MAXIMO_USABLE, X_MINIMO_USABLE, Y_MINIMO_VISIBLE
        Else
            MapaN = 0
        End If
   End If
    
    'SUR
    ' Tengo que copiar el NORTE del mapa SUR y pegarselo en la parte inferior del mapa actual
    If MapaS > 0 Then
        If pakMapasME.Cabezal_GetFileSize(MapaS) Then
            SwitchMap MapaS
            PegarAreaDesdeHasta X_MINIMO_USABLE, Y_MINIMO_USABLE, X_MAXIMO_USABLE, Y_MINIMO_NO_VISIBLE_OTRO_MAPA - 1, X_MINIMO_USABLE, Y_MAXIMO_USABLE + 1
        Else
            MapaS = 0
        End If
    End If
    
    'ESTE
    If MapaE > 0 Then
        If pakMapasME.Cabezal_GetFileSize(MapaE) Then
            SwitchMap MapaE
            PegarAreaDesdeHasta X_MINIMO_USABLE, Y_MINIMO_USABLE, X_MINIMO_NO_VISIBLE_OTRO_MAPA - 1, Y_MAXIMO_USABLE, X_MAXIMO_USABLE + 1, Y_MINIMO_USABLE
        Else
            MapaE = 0
        End If
    End If
    
    'OESTE
    If MapaO > 0 Then
        If pakMapasME.Cabezal_GetFileSize(MapaO) Then
            SwitchMap MapaO
            PegarAreaDesdeHasta X_MAXIMO_NO_VISIBLE_OTRO_MAPA + 1, Y_MINIMO_USABLE, X_MAXIMO_USABLE, Y_MAXIMO_USABLE, X_MINIMO_VISIBLE, Y_MINIMO_USABLE
        Else
            MapaO = 0
        End If
    End If
    
    'NOROESTE
    If MapaNO > 0 Then
        If pakMapasME.Cabezal_GetFileSize(MapaNO) Then
            SwitchMap MapaNO
            PegarAreaDesdeHasta X_MAXIMO_NO_VISIBLE_OTRO_MAPA + 1, Y_MAXIMO_NO_VISIBLE_OTRO_MAPA + 1, X_MAXIMO_USABLE, Y_MAXIMO_USABLE, X_MINIMO_VISIBLE, Y_MINIMO_VISIBLE
        Else
            MapaNO = 0
        End If
    End If
    
    'NORESTE
    If MapaNE > 0 Then
        If pakMapasME.Cabezal_GetFileSize(MapaNE) Then
           SwitchMap MapaNE
            PegarAreaDesdeHasta X_MINIMO_USABLE, Y_MAXIMO_NO_VISIBLE_OTRO_MAPA + 1, X_MINIMO_NO_VISIBLE_OTRO_MAPA - 1, Y_MAXIMO_USABLE, X_MAXIMO_USABLE + 1, Y_MINIMO_VISIBLE
        Else
            MapaNE = 0
        End If
    End If
    
    'SUROESTE
    If MapaSO > 0 Then
       If pakMapasME.Cabezal_GetFileSize(MapaSO) Then
            SwitchMap MapaSO
            PegarAreaDesdeHasta X_MAXIMO_NO_VISIBLE_OTRO_MAPA + 1, Y_MINIMO_USABLE, X_MAXIMO_NO_VISIBLE_OTRO_MAPA + 1 + (BORDE_TILES_INUTILIZABLE - 1), Y_MINIMO_USABLE + (BORDE_TILES_INUTILIZABLE - 1), X_MINIMO_VISIBLE, Y_MAXIMO_USABLE + 1
        Else
            MapaSO = 0
        End If
    End If
    
    'SURESTE
    If MapaSE > 0 Then
        If pakMapasME.Cabezal_GetFileSize(MapaSE) Then
            SwitchMap MapaSE
           PegarAreaDesdeHasta X_MINIMO_USABLE, Y_MINIMO_USABLE, X_MINIMO_USABLE + (BORDE_TILES_INUTILIZABLE - 1), Y_MINIMO_USABLE + (BORDE_TILES_INUTILIZABLE - 1), X_MAXIMO_USABLE + 1, Y_MAXIMO_USABLE + 1
        Else
            MapaSE = 0
        End If
    End If

    'Abro el mapa donde tengo que pegar los pedazos
    Call ME_FIFO.prepararWorkEspace
            
    SwitchMap numeroMapa
    'Luego de cargar el mapa refrescamos la lista de acciones usadas en este mapa
    Call ME_modAccionEditor.refrescarListaUsando(frmMain.listTileAccionActuales)
    

'Pego los pedazos
    If MapaN > 0 Then PegarArea X_MINIMO_USABLE, Y_MINIMO_VISIBLE, X_MAXIMO_USABLE, Y_MINIMO_USABLE - 1
    If MapaS > 0 Then PegarArea X_MINIMO_USABLE, Y_MAXIMO_USABLE + 1, X_MAXIMO_USABLE, Y_MAXIMO_VISIBLE
    If MapaE > 0 Then PegarArea X_MAXIMO_USABLE + 1, Y_MINIMO_USABLE, X_MAXIMO_VISIBLE, Y_MAXIMO_USABLE
    If MapaO > 0 Then PegarArea X_MINIMO_VISIBLE, Y_MINIMO_USABLE, X_MINIMO_JUGABLE, Y_MAXIMO_USABLE
    If MapaNO > 0 Then PegarArea X_MINIMO_VISIBLE, Y_MINIMO_VISIBLE, X_MINIMO_USABLE - 1, Y_MINIMO_USABLE - 1
    If MapaNE > 0 Then PegarArea X_MAXIMO_USABLE + 1, Y_MINIMO_VISIBLE, X_MAXIMO_USABLE + 1 + BORDE_TILES_INUTILIZABLE - 1, Y_MINIMO_VISIBLE + BORDE_TILES_INUTILIZABLE - 1
    If MapaSO > 0 Then PegarArea X_MINIMO_VISIBLE, Y_MAXIMO_USABLE + 1, X_MINIMO_VISIBLE + (BORDE_TILES_INUTILIZABLE - 1), Y_MAXIMO_USABLE + 1 + (BORDE_TILES_INUTILIZABLE - 1)
    If MapaSE > 0 Then PegarArea X_MAXIMO_USABLE + 1, Y_MAXIMO_USABLE + 1, X_MAXIMO_USABLE + 1 + (BORDE_TILES_INUTILIZABLE - 1), Y_MAXIMO_USABLE + 1 + (BORDE_TILES_INUTILIZABLE - 1)
    
    'Actualizo el mapa
    ActualizarArraysAlturasMapas
    
    Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)
    
    'Heightmap_Calculate -63, -63, 0
    Compute_Mountain
    'If cron_tiempo = False Then
    cron_tiempo
    Light_Update_Map = True
    Light_Update_Sombras = True
        
End Function

Private Sub PegarArea(ByVal x1 As Integer, ByVal y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
    Dim X As Integer, Y As Integer, i As Integer
    
    Debug.Print "Copio a ACTUAL desde (" & x1 & ";" & y1; ") Hasta (" & X2 & ";" & Y2 & ")  Tamaño " & (Abs(X2 - x1) + 1) & ";" & (Abs(Y2 - y1) + 1)
    
    For X = x1 To X2
        For Y = y1 To Y2
            For i = 1 To CANTIDAD_CAPAS
                mapdata(X, Y).Graphic(i) = mapaCopiaTile(X, Y).Graphic(i)
            Next
            
            If mapdata(X, Y).luz Then
                If DLL_Luces.Check(mapdata(X, Y).luz) Then DLL_Luces.Quitar mapdata(X, Y).luz
                mapdata(X, Y).luz = 0
            End If
            
            For i = 0 To 2
                Set mapdata(X, Y).Particles_groups(i) = mapaCopiaTile(X, Y).Particles_groups(i)
                If Not mapdata(X, Y).Particles_groups(i) Is Nothing Then
                    Call mapdata(X, Y).Particles_groups(i).SetPos(X, Y)
                End If
            Next
            
            mapdata(X, Y).Trigger = mapaCopiaTile(X, Y).Trigger
             
            mapdata(X, Y).tile_number = mapaCopiaTile(X, Y).tile_number
            mapdata(X, Y).tile_orientation = mapaCopiaTile(X, Y).tile_orientation
            mapdata(X, Y).tile_render = mapaCopiaTile(X, Y).tile_render
            mapdata(X, Y).tile_texture = mapaCopiaTile(X, Y).tile_texture
            
            hMapData(X, Y) = mapaCopiahMapData(X, Y)
            Alturas(X, Y) = mapaCopiaAlturas(X, Y)
            AlturaPie(X, Y) = mapaCopiaAlturaPie(X, Y)
            OriginalMapColor(X, Y) = mapaCopiaColor(X, Y)
            NormalData(X, Y) = mapaCopiaNormalData(X, Y)
            Intensidad_Del_Terreno(X, Y) = mapaCopiaLuninosidad(X, Y)
            
            If mapaCopiaLuces(X, Y).tieneLuz Then
                If EsLuzValida(mapaCopiaLuces(X, Y).LRange, mapaCopiaLuces(X, Y).LBrillo, mapaCopiaLuces(X, Y).LTipo) Then
                    mapdata(X, Y).luz = DLL_Luces.crear(X, Y, mapaCopiaLuces(X, Y).lColor.r, mapaCopiaLuces(X, Y).lColor.g, mapaCopiaLuces(X, Y).lColor.b, mapaCopiaLuces(X, Y).LRange, mapaCopiaLuces(X, Y).LBrillo, mapaCopiaLuces(X, Y).LTipo, mapaCopiaLuces(X, Y).luzInicio, mapaCopiaLuces(X, Y).luzFin)
                End If
            End If
        Next
    Next
End Sub

Private Sub PegarAreaDesdeHasta( _
    ByVal x1 As Integer, ByVal y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, _
    ByVal Hasta_X As Integer, ByVal Hasta_Y As Integer _
)
    
    Dim X As Integer, Y As Integer, i As Integer
    
    Dim X_ As Integer
    Dim Y_ As Integer
    
    Debug.Print "Copio a memoria desde " & x1 & ";" & y1; " Hasta " & X2 & ";" & Y2; " Pego en (" & Hasta_X & ";" & Hasta_Y & ")"; " Tamaño " & (Abs(X2 - x1) + 1) & ";" & (Abs(Y2 - y1) + 1)
    
    For X = x1 To X2
        
        ' Posicion en X en donde se pegara
        X_ = X - x1 + Hasta_X
        
        
        For Y = y1 To Y2
            
            ' Posicion en Y en donde se pegara
            Y_ = Y - y1 + Hasta_Y
       
            For i = 1 To CANTIDAD_CAPAS
                mapaCopiaTile(X_, Y_).Graphic(i) = mapdata(X, Y).Graphic(i)
            Next
            
            For i = 0 To 2
                Set mapaCopiaTile(X_, Y_).Particles_groups(i) = mapdata(X, Y).Particles_groups(i)
            Next
            
            mapaCopiaTile(X_, Y_).Trigger = mapdata(X, Y).Trigger
            
            
            mapaCopiaTile(X_, Y_).tile_number = mapdata(X, Y).tile_number
            mapaCopiaTile(X_, Y_).tile_orientation = mapdata(X, Y).tile_orientation
            mapaCopiaTile(X_, Y_).tile_render = mapdata(X, Y).tile_render
            mapaCopiaTile(X_, Y_).tile_texture = mapdata(X, Y).tile_texture
            
            mapaCopiahMapData(X_, Y_) = hMapData(X, Y)
            mapaCopiaAlturas(X_, Y_) = Alturas(X, Y)
            mapaCopiaAlturaPie(X_, Y_) = AlturaPie(X, Y)
            mapaCopiaColor(X_, Y) = OriginalMapColor(X, Y)
            mapaCopiaNormalData(X_, Y_) = NormalData(X, Y)
            mapaCopiaLuninosidad(X_, Y_) = Intensidad_Del_Terreno(X, Y)
            

            mapaCopiaLuces(X_, Y_).tieneLuz = False
            
            If mapdata(X, Y).luz Then
                If DLL_Luces.Check(mapdata(X, Y).luz) Then
                    Dim tmpByte1 As Byte
                    Dim tmpByte2 As Byte
                    DLL_Luces.Get_Light mapdata(X, Y).luz, tmpByte1, tmpByte2, mapaCopiaLuces(X_, Y_).lColor.r, mapaCopiaLuces(X_, Y_).lColor.g, mapaCopiaLuces(X_, Y_).lColor.b, mapaCopiaLuces(X_, Y_).LRange, mapaCopiaLuces(X_, Y_).LBrillo, mapaCopiaLuces(X_, Y_).LTipo, mapaCopiaLuces(X_, Y_).luzInicio, mapaCopiaLuces(X_, Y_).luzFin
                    If mapaCopiaLuces(X_, Y_).LRange > 0 Then
                        mapaCopiaLuces(X_, Y_).tieneLuz = True
                    End If
                End If
            End If
        Next
    Next
End Sub


