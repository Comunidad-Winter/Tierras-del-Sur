Attribute VB_Name = "ME_FIFO"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Any, ByVal numbytes As Long)

Public AbriendoMapa As Boolean


Function AbrirMapa(Path As String) As Boolean

'abre .map
THIS_MAPA.Path = Path

'AbrirMapaCompilado path, 0, 0
Cargar_Mapa_ME Path
 
' Actualizamos el MiniMapa
Call miniMap_Redraw
 
'Luego de cargar el mapa refrescamos la lista de acciones usadas en este mapa
Call ME_modAccionEditor.refrescarListaUsando(frmMain.listTileAccionActuales)

 
End Function

Function GuardarMapa(Path As String) As Boolean
'guarda .map
THIS_MAPA.editado = False
THIS_MAPA.Path = Path
GuardarMapa = Guardar_Mapa_ME(Path)

End Function

Function Guardar_Mapa_ME(ByVal SaveAs As String) As Boolean

Dim freeFileMap As Integer
Dim ByFlags As Integer

Dim loopC As Long
Dim y As Long
Dim x As Long
Dim loopZona As Byte
Dim tit As Integer

Dim offsetHeader As Long
Dim ResizeBackBufferY As Integer
Dim ResizeBackBufferX As Integer

Dim tempLuz As tLuzPropiedades
Dim posLuzX As Byte
Dim posLuzY As Byte
                        
Dim checkSum As String * 10
                
If FileExist(SaveAs, vbNormal) = True Then
    Kill SaveAs
End If
        
freeFileMap = FreeFile

Open SaveAs For Binary As freeFileMap
Seek freeFileMap, 1
    
    '###################################################################################
    '########### HEADER

    Put freeFileMap, , header_m_3
    Put freeFileMap, , THIS_MAPA.numero
    
    '###################################################################################
    '########### PROPIEDADES
    offsetHeader = seek(freeFileMap)
    
    Put freeFileMap, , ResizeBackBufferX
    Put freeFileMap, , ResizeBackBufferY

    '###################################################################################
    '########### COLORES Y LUCES!
    
    'TODO Precalcular luces estáticas.
    Put freeFileMap, , OriginalMapColor
    Put freeFileMap, , Intensidad_Del_Terreno
    
    With mapinfo
        Put freeFileMap, , .BaseColor
        Put freeFileMap, , .ColorPropio
        
        Put freeFileMap, , .agua_tileset
        Put freeFileMap, , .agua_rect
        Put freeFileMap, , .agua_profundidad
        
        Put freeFileMap, , .UsaAguatierra
    End With
    
    '*************************************************************************
    'Formularios de Acciones
    Call ME_modAccionEditor.persistirListaAccionTileEditorUsando(freeFileMap)
    '*************************************************************************
    ' Zonas de Criaturas creadas
    ' Guardo la cantidad
    Put freeFileMap, , CByte(UBound(mapinfo.ZonasNacCriaturas) + 1)
    
    For loopZona = 0 To UBound(mapinfo.ZonasNacCriaturas)
        Put freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).Superior
        Put freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).Inferior
        Put freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).nombre & Space$(15 - Len(mapinfo.ZonasNacCriaturas(loopZona).nombre))
    Next
    '*************************************************************************
    
    ' Flags. Disponibles: 15
    
    ' Capa 5 -> 0
    ' Capa 1 -> 1
    ' Capa 2 -> 2
    ' Capa 3 -> 3
    ' Capa 4 -> 4
    ' Trigger -> 5
    ' TileSet -> 6
    ' Particula 0 -> 7
    ' Particula 2 -> 8
    ' Montaña -> 9
    ' Acciones -> 10
    ' Criaturas -> 11
    ' Objetos -> 12
    ' Luz -> 13
    ' Particula 1 -> 14
    For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
            With mapdata(x, y)
            
                ByFlags = 0
                                
                '0 a 4) Capas
                If .Graphic(1).GrhIndex Then ByFlags = ByFlags Or bitwisetable(1)
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or bitwisetable(2)
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or bitwisetable(3)
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or bitwisetable(4)
                If .Graphic(5).GrhIndex Then ByFlags = ByFlags Or bitwisetable(0)
                
                '5) Trigger
                If .Trigger Then _
                    ByFlags = ByFlags Or bitwisetable(5)
                
                '6) Piso
                If .tile_texture Then _
                    ByFlags = ByFlags Or bitwisetable(6)
                
                '9) Alturas
                If Alturas(x, y) <> 0 Or AlturaPie(x, y) <> 0 Or hMapData(x, y).hs(0) <> 0 Then _
                    ByFlags = ByFlags Or bitwisetable(9)
                
                '10) Acciones
                If Not mapdata(x, y).accion Is Nothing Then
                    ByFlags = ByFlags Or bitwisetable(10)
                End If
                
                '11) Criaturas
                If .NpcIndex Then _
                    ByFlags = ByFlags Or bitwisetable(11)
                   
                '12) Objetos
                If .OBJInfo.objIndex Then _
                    ByFlags = ByFlags Or bitwisetable(12)
                    
                    
                '13) Luz
                If .luz Then
                    If DLL_Luces.Check(.luz) Then
                    
                        ' Tomamos las Propiedades
                        DLL_Luces.Get_Light .luz, posLuzX, posLuzY, tempLuz.LuzColor.r, tempLuz.LuzColor.g, tempLuz.LuzColor.b, tempLuz.LuzRadio, tempLuz.LuzBrillo, tempLuz.LuzTipo, tempLuz.luzInicio, tempLuz.luzInicio
                        
                        ' Ultimo chequeo antes de guardar la luz
                        If tempLuz.LuzRadio > 0 Then ByFlags = ByFlags Or bitwisetable(13)
                    End If
                End If
                
                '7)  Particula 0
                If Not .Particles_groups(0) Is Nothing Then ByFlags = ByFlags Or bitwisetable(7)
                
                '14) Particula 1
                If Not .Particles_groups(1) Is Nothing Then ByFlags = ByFlags Or bitwisetable(14)
                
                '8) Particula 2
                If Not .Particles_groups(2) Is Nothing Then ByFlags = ByFlags Or bitwisetable(8)
                
                ' Guardamos el Flag
                Put freeFileMap, , ByFlags
                
                ' ------------------------------------------------------------ '
                '           Guardado de Datos                                   '
                
                ' 6) Piso
                If ByFlags And bitwisetable(6) Then
                    Put freeFileMap, , .tile_texture
                    Put freeFileMap, , .tile_number
                End If
                
                ' 0 .. 4) Capas
                For loopC = 1 To CANTIDAD_CAPAS
                    If .Graphic(loopC).GrhIndex Then Put freeFileMap, , .Graphic(loopC).GrhIndex
                Next loopC
                
                ' Analizo el tamaño para ver si debo agrandar el area que el Juego analiza
                If ByFlags And bitwisetable(3) Then
                    If ResizeBackBufferY < GrhData(.Graphic(3).GrhIndex).pixelHeight Then ResizeBackBufferY = GrhData(.Graphic(3).GrhIndex).pixelHeight
                    If ResizeBackBufferX < GrhData(.Graphic(3).GrhIndex).pixelWidth Then ResizeBackBufferX = GrhData(.Graphic(3).GrhIndex).pixelWidth
                End If
                
                '5) Trigger
                If ByFlags And bitwisetable(5) Then Put freeFileMap, , .Trigger
                    
                '9) Altura
                If ByFlags And bitwisetable(9) Then
                    Put freeFileMap, , NormalData(x, y)     'Normal del vértice
                    Put freeFileMap, , Alturas(x, y)        'Alturas
                    Put freeFileMap, , AlturaPie(x, y)      'Altura Pie
                    tit = Round(hMapData(x, y).hs(0))
                    Put freeFileMap, , tit                  'Altura del vértice
                    Put freeFileMap, , .tile_orientation    'Orientación de la tile
                End If
                
                '10) Acciones
                If ByFlags And bitwisetable(10) Then
                    Put freeFileMap, , mapdata(x, y).accion.getID
                End If
                
                '11) Criaturas
                If ByFlags And bitwisetable(11) Then
                    Put freeFileMap, , .NpcIndex
                End If
                
                '12) Objeto
                If ByFlags And bitwisetable(12) Then
                    Put freeFileMap, , .OBJInfo.objIndex
                    Put freeFileMap, , .OBJInfo.Amount
                End If

                ' 13) Luz
                If ByFlags And bitwisetable(13) Then
                    Put freeFileMap, , tempLuz.LuzColor.b
                    Put freeFileMap, , tempLuz.LuzColor.g
                    Put freeFileMap, , tempLuz.LuzColor.r
                    Put freeFileMap, , tempLuz.LuzRadio
                    Put freeFileMap, , tempLuz.LuzBrillo
                    Put freeFileMap, , tempLuz.LuzTipo
                    Put freeFileMap, , tempLuz.luzInicio
                    Put freeFileMap, , tempLuz.luzFin
                End If
                
                '7)  Particula 0
                If ByFlags And bitwisetable(7) Then .Particles_groups(0).persistir freeFileMap
                '14) Particula 1
                If ByFlags And bitwisetable(14) Then .Particles_groups(1).persistir freeFileMap
                '8) Particula 2
                If ByFlags And bitwisetable(8) Then .Particles_groups(2).persistir freeFileMap
                
                ' CheckSum
                checkSum = "hola"
                Put freeFileMap, , checkSum
                 
            End With
        Next x
    Next y
    
    ' Actualizamos el Offset
    Seek freeFileMap, offsetHeader
    
    Put freeFileMap, , ResizeBackBufferX
    Put freeFileMap, , ResizeBackBufferY
    
    Close freeFileMap
    
    Guardar_Mapa_ME = True
Exit Function

ErrorSave:
    MsgBox "Error en GuardarVTDS, nro. " & Err.Number & " - " & Err.description
End Function

Private Sub Cargar_Mapa_Me_1(freeFileMap As Integer)

    Dim tempInt As Integer
    Dim tempbyte As Byte
    Dim body As Integer
    Dim Head As Integer
    Dim heading As Byte
    Dim y As Integer
    Dim x As Integer
    Dim ByFlags As Integer

    Dim NombreMapa As String * 32
    Dim MapaNumeroOriginal As Integer
    Dim ResizeBackBufferX As Integer
    Dim ResizeBackBufferY As Integer
    
    Dim plusa As Integer
    
    Dim nmap As String
    
    
    Dim LRange      As Byte
    Dim LBrillo     As Byte
    Dim LTipo       As Integer
    Dim luzInicio As Byte
    Dim luzFin As Byte
    
    Dim loopZona As Byte
    
    Dim TamanioMapa As Integer
        
    'LUCES
    ReDim Intensidad_Del_Terrenob(1 To 218, 1 To 218) As Byte               'Guarda la intensidad de la luz de un vertice del mapa
    ReDim OriginalMapColorb(1 To 218, 1 To 218) As BGRACOLOR_DLL            'Colores precalculados en el mapeditor
    ReDim OriginalMapColorSombraB(1 To 218, 1 To 218) As Long              'OriginalMapColor * Sombra
    ReDim OriginalColorArrayB(1 To 218, 1 To 218) As Long                   'BACKUP DE ResultColorArray (OriginalMapColorSombra * AMBIENTE)
    ReDim ResultColorArrayB(1 To 218, 1 To 218) As Long                    'OriginalColorArray * LUCES DINÁMICAS * SOMBRAS
    
    'Altura de cada vertice del mapa
    ReDim hmapdatab(1 To 218, 1 To 218) As AUDT
    ' Altura de donde pisa el pj, o donde flotan las cosas, o donde el árbol vuela. Sirve para ahcer escaleras.
    ReDim AlturaPieb(1 To 218, 1 To 218) As Integer
    'Es >0 si en la tile hay una altura distnta a cero.
    ReDim Alturasb(1 To 218, 1 To 218) As Integer
    ' Almacena el vector normalizado de los triángulos del mapa, para calcular la sombra Intensidad_sombra = DOT(NORMALIZED•SOL_POS)
    ReDim NormalDatab(1 To 218, 1 To 218) As D3DVECTOR
    ReDim Sombra_MontañasB(1 To 218, 1 To 218) As Byte
    ReDim MapBoxesB(1 To 218, 1 To 218) As Box_Vertex
    '/MONTAÑAS
    
    ReDim mapdatab(1 To 218, 1 To 218) As MapBlock
    
    
    Get freeFileMap, , NombreMapa
    Get freeFileMap, , MapaNumeroOriginal
    
    THIS_MAPA.nombre = NombreMapa
    THIS_MAPA.numero = MapaNumeroOriginal
        
    nmap = Trim$(NombreMapa)
    If Not Len(nmap) > 0 Then nmap = "Mapa sin nombre"
    
    '###################################################################################
    '########### PROPIEDADES
    
    Get freeFileMap, , ResizeBackBufferX
    Get freeFileMap, , ResizeBackBufferY

    '###################################################################################
    '########### COLORES Y LUCES!
    
    Get freeFileMap, , OriginalMapColorb
    Get freeFileMap, , Intensidad_Del_Terrenob

    With mapinfo
        
        Get freeFileMap, , .BaseColor
        Get freeFileMap, , .ColorPropio
        
        Get freeFileMap, , .agua_tileset
        Get freeFileMap, , .agua_rect
        Get freeFileMap, , .agua_profundidad
        
        Get freeFileMap, , .puede_nieve
        Get freeFileMap, , .puede_lluvia
        Get freeFileMap, , .puede_neblina
        Get freeFileMap, , .puede_niebla
        Get freeFileMap, , .puede_sandstorm
        Get freeFileMap, , .puede_nublado
        
        Get freeFileMap, , .UsaAguatierra
                        
        Get freeFileMap, , .SonidoLoop
        
        Get freeFileMap, , .MinNivel
        Get freeFileMap, , .MaxNivel
    
        Get freeFileMap, , .UsuariosMaximo
        Get freeFileMap, , .SeCaenItems
        Get freeFileMap, , .PermiteRoboNpc
        Get freeFileMap, , .PermiteHechizosPetes
        Get freeFileMap, , .MagiaSinEfecto
        Get freeFileMap, , .MapaPK
        
        'EXPANSIONES?

        Get freeFileMap, , TamanioMapa
        
        If TamanioMapa < 100 Then
            TamanioMapa = 100
        End If
        
        Get freeFileMap, , tempInt '1
        Get freeFileMap, , tempInt '2
        Get freeFileMap, , tempInt '3
        Get freeFileMap, , tempInt '4
        Get freeFileMap, , tempInt '5
        Get freeFileMap, , tempInt '6
        
        Get freeFileMap, , tempInt '7
        Get freeFileMap, , tempInt '8
        Get freeFileMap, , tempInt '9
        Get freeFileMap, , tempInt '10
        Get freeFileMap, , tempInt '11
        Get freeFileMap, , tempInt '12
        Get freeFileMap, , tempInt '13
        Get freeFileMap, , tempInt '14
        Get freeFileMap, , tempInt '15
        Get freeFileMap, , tempInt '16
        Get freeFileMap, , tempInt '17
        Get freeFileMap, , tempInt '18
        Get freeFileMap, , tempInt '19
        Get freeFileMap, , tempInt '20
        Get freeFileMap, , tempInt '21
        
        Get freeFileMap, , tempInt '22
        Get freeFileMap, , tempInt '23
        Get freeFileMap, , tempInt '24
        Get freeFileMap, , tempInt '25
        Get freeFileMap, , tempInt '26
        Get freeFileMap, , tempInt '27
        Get freeFileMap, , tempInt '28
        Get freeFileMap, , tempInt '29
        Get freeFileMap, , tempInt '30
        Get freeFileMap, , tempInt '31
        Get freeFileMap, , tempInt '32
        Get freeFileMap, , tempInt '33
        Get freeFileMap, , tempInt '34
        Get freeFileMap, , tempInt '35
        Get freeFileMap, , tempInt '36
        
    End With

    '*************************************************************************
    'Formularios de Acciones
    Call ME_modAccionEditor.cargarListaAccionesEditorUsando(freeFileMap)
    '*************************************************************************
    ' Zonas
    Get freeFileMap, , tempbyte
    
    ReDim mapinfo.ZonasNacCriaturas(tempbyte) As ZonaNacimientoCriatura
    
    For loopZona = 0 To tempbyte - 1
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).Superior
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).Inferior
        mapinfo.ZonasNacCriaturas(loopZona).nombre = Space$(15)
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).nombre
        mapinfo.ZonasNacCriaturas(loopZona).nombre = Trim$(mapinfo.ZonasNacCriaturas(loopZona).nombre)
    Next
    '*************************************************************************
    
    
    'Load arrays
    For y = 1 To TamanioMapa
        For x = 1 To TamanioMapa
    
            Get freeFileMap, , ByFlags
            
            'mapdatab(X, Y).Blocked = (ByFlags And bitwisetable(0))
            
            'mapdatab(X, Y).is_water = (ByFlags And bitwisetable(8))
            
            If ByFlags And bitwisetable(6) Then
                Get freeFileMap, , mapdatab(x, y).tile_texture
                Get freeFileMap, , mapdatab(x, y).tile_number
            Else
                mapdatab(x, y).tile_texture = 0
            End If
            
            If ByFlags And bitwisetable(1) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(1).GrhIndex
                InitGrh mapdatab(x, y).Graphic(1), mapdatab(x, y).Graphic(1).GrhIndex
            Else
                mapdatab(x, y).Graphic(1).GrhIndex = 0
            End If
                
            If ByFlags And bitwisetable(2) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(2).GrhIndex
                InitGrh mapdatab(x, y).Graphic(2), mapdatab(x, y).Graphic(2).GrhIndex
            Else
                mapdatab(x, y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And bitwisetable(3) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(3).GrhIndex
                InitGrh mapdatab(x, y).Graphic(3), mapdatab(x, y).Graphic(3).GrhIndex
            Else
                mapdatab(x, y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And bitwisetable(4) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(4).GrhIndex
                InitGrh mapdatab(x, y).Graphic(4), mapdatab(x, y).Graphic(4).GrhIndex
            Else
                mapdatab(x, y).Graphic(4).GrhIndex = 0
            End If

            'Trigger used?
            If ByFlags And bitwisetable(5) Then
                Get freeFileMap, , mapdatab(x, y).Trigger
            Else
                mapdatab(x, y).Trigger = 0
            End If
            
            If ByFlags And bitwisetable(9) Then
                    Get freeFileMap, , NormalDatab(x, y)
                    Get freeFileMap, , Alturasb(x, y)
                    Get freeFileMap, , AlturaPieb(x, y)
                    Get freeFileMap, , plusa
                    Get freeFileMap, , mapdatab(x, y).tile_orientation
                    hmapdatab(x, y).h = Alturasb(x, y)
                    hmapdatab(x, y).hs(0) = plusa
            End If
            
            Dim tempintMarce As Integer
            
            If ByFlags And bitwisetable(10) Then
                Get freeFileMap, , tempintMarce
                Set mapdatab(x, y).accion = ME_modAccionEditor.obtenerAccionID(tempintMarce)
            End If
    
            If ByFlags And bitwisetable(11) Then
                'Get and make NPC
                Get freeFileMap, , mapdatab(x, y).NpcIndex

    
                If mapdatab(x, y).NpcIndex < 0 Then
                    mapdatab(x, y).NpcIndex = 0
                Else
                    body = NpcData(mapdatab(x, y).NpcIndex).body
                    Head = NpcData(mapdatab(x, y).NpcIndex).Head
                    heading = NpcData(mapdatab(x, y).NpcIndex).heading
                    Dim Char As Integer
                    
                    'Creo el NPC
                    Char = SV_Simulador.NextOpenChar(True)
                    
                    Call MakeChar(Char, body, Head, heading, x, y, 0, 0, 0)

                    CharList(Char).active = 1
                    
                    CharMap(x, y) = Char
                                        
                End If
            End If
    
            If ByFlags And bitwisetable(12) Then
                'Get and make Object
                Get freeFileMap, , mapdatab(x, y).OBJInfo.objIndex
                Get freeFileMap, , mapdatab(x, y).OBJInfo.Amount
                If mapdatab(x, y).OBJInfo.objIndex > 0 Then
                    InitGrh mapdatab(x, y).ObjGrh, ObjData(mapdatab(x, y).OBJInfo.objIndex).GrhIndex
                End If
            End If
            
            If ByFlags And bitwisetable(13) Then
                Dim luz_r As Byte, luz_g As Byte, luz_b As Byte
                Get freeFileMap, , luz_b
                Get freeFileMap, , luz_g
                Get freeFileMap, , luz_r
                Get freeFileMap, , LRange
                Get freeFileMap, , LBrillo
                Get freeFileMap, , LTipo
                Get freeFileMap, , luzInicio
                Get freeFileMap, , luzFin
                If EsLuzValida(LRange, LBrillo, LTipo) Then
                    mapdatab(x, y).luz = DLL_Luces.crear(x, y, luz_r, luz_g, luz_b, LRange, LBrillo, LTipo, luzInicio, luzFin)
                End If
            Else
                mapdatab(x, y).luz = 0
            End If
            
            '9)  Particula 0
            If ByFlags And bitwisetable(7) Then
                Set mapdatab(x, y).Particles_groups(0) = New Engine_Particle_Group
                
                If mapdatab(x, y).Particles_groups(0).Cargar(freeFileMap) Then
                    mapdatab(x, y).Particles_groups(0).SetPos x, y
                Else
                    Set mapdatab(x, y).Particles_groups(0) = Nothing
                End If
            End If
            
            '10) Particula 1
            If ByFlags And bitwisetable(14) Then
                Set mapdatab(x, y).Particles_groups(1) = New Engine_Particle_Group
                
                If mapdatab(x, y).Particles_groups(1).Cargar(freeFileMap) Then
                    mapdatab(x, y).Particles_groups(1).SetPos x, y
                Else
                    Set mapdatab(x, y).Particles_groups(1) = Nothing
                End If
            End If
            
            '8) Particula 2
            If ByFlags And bitwisetable(8) Then
                Set mapdatab(x, y).Particles_groups(2) = New Engine_Particle_Group
                
                If mapdatab(x, y).Particles_groups(2).Cargar(freeFileMap) Then
                    mapdatab(x, y).Particles_groups(2).SetPos x, y
                Else
                    Set mapdatab(x, y).Particles_groups(2) = Nothing
                End If
            End If
            
        Next x
    Next y
    
    'Close files
    For x = SV_Constantes.X_MINIMO_VISIBLE To SV_Constantes.X_MAXIMO_VISIBLE
        For y = SV_Constantes.Y_MINIMO_JUGABLE To SV_Constantes.Y_MAXIMO_JUGABLE
            mapdata(x, y) = mapdatab(x, y)
        Next y
    Next x

    
    ActualizarArraysAlturasMapas
    
    frmMain.setValoresAgua

    mapinfo.MaxGrhSizeXInTiles = ResizeBackBufferX \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
    mapinfo.MaxGrhSizeYInTiles = ResizeBackBufferY \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
  
    Engine_Set_TileBuffer_Size mapinfo.MaxGrhSizeXInTiles, mapinfo.MaxGrhSizeYInTiles

    Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)
    
    'Heightmap_Calculate -63, -63, 0
Compute_Mountain
    'If cron_tiempo = False Then
        cron_tiempo
        Light_Update_Map = True
        Light_Update_Sombras = True
        
        'map_render_light
    'End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    THIS_MAPA.editado = False
    
    Call Me_Tools_Npc.cargarZonasDeNacimiento(mapinfo.ZonasNacCriaturas)
    
End Sub

Private Sub Cargar_Mapa_Me_0(freeFileMap As Integer)

    Dim tempInt As Integer
    Dim body As Integer
    Dim Head As Integer
    Dim heading As Byte
    Dim y As Integer
    Dim x As Integer
    Dim ByFlags As Integer

    Dim NombreMapa As String * 32
    Dim MapaNumeroOriginal As Integer
    Dim ResizeBackBufferX As Integer
    Dim ResizeBackBufferY As Integer
    
    Dim plusa As Integer
    
    Dim nmap As String
    
    
    Dim LRange      As Byte
    Dim LBrillo     As Byte
    Dim LTipo       As Integer
    Dim luzInicio As Byte
    Dim luzFin As Byte
    
    
    Dim TamanioMapa As Integer

    Get freeFileMap, , NombreMapa
    Get freeFileMap, , MapaNumeroOriginal
    
    THIS_MAPA.nombre = NombreMapa
    THIS_MAPA.numero = MapaNumeroOriginal
        
    nmap = Trim$(NombreMapa)
    If Not Len(nmap) > 0 Then nmap = "Mapa sin nombre"
    
    '###################################################################################
    '########### PROPIEDADES
    
    Get freeFileMap, , ResizeBackBufferX
    Get freeFileMap, , ResizeBackBufferY

    '###################################################################################
    '########### COLORES Y LUCES!
    
    Get freeFileMap, , OriginalMapColor
    Get freeFileMap, , Intensidad_Del_Terreno

    With mapinfo
        
        Get freeFileMap, , .BaseColor
        Get freeFileMap, , .ColorPropio
        
        Get freeFileMap, , .agua_tileset
        Get freeFileMap, , .agua_rect
        Get freeFileMap, , .agua_profundidad
        
        Get freeFileMap, , .puede_nieve
        Get freeFileMap, , .puede_lluvia
        Get freeFileMap, , .puede_neblina
        Get freeFileMap, , .puede_niebla
        Get freeFileMap, , .puede_sandstorm
        Get freeFileMap, , .puede_nublado
        
        Get freeFileMap, , .UsaAguatierra
                        
        Get freeFileMap, , .SonidoLoop
        
        Get freeFileMap, , .MinNivel
        Get freeFileMap, , .MaxNivel
    
        Get freeFileMap, , .UsuariosMaximo
        Get freeFileMap, , .SeCaenItems
        Get freeFileMap, , .PermiteRoboNpc
        Get freeFileMap, , .PermiteHechizosPetes
        Get freeFileMap, , .MagiaSinEfecto
        Get freeFileMap, , .MapaPK
        
        'EXPANSIONES?

        Get freeFileMap, , TamanioMapa
        
        If TamanioMapa < 100 Then
            TamanioMapa = 100
        End If
        
        Get freeFileMap, , tempInt '1
        Get freeFileMap, , tempInt '2
        Get freeFileMap, , tempInt '3
        Get freeFileMap, , tempInt '4
        Get freeFileMap, , tempInt '5
        Get freeFileMap, , tempInt '6
        
        Get freeFileMap, , tempInt '7
        Get freeFileMap, , tempInt '8
        Get freeFileMap, , tempInt '9
        Get freeFileMap, , tempInt '10
        Get freeFileMap, , tempInt '11
        Get freeFileMap, , tempInt '12
        Get freeFileMap, , tempInt '13
        Get freeFileMap, , tempInt '14
        Get freeFileMap, , tempInt '15
        Get freeFileMap, , tempInt '16
        Get freeFileMap, , tempInt '17
        Get freeFileMap, , tempInt '18
        Get freeFileMap, , tempInt '19
        Get freeFileMap, , tempInt '20
        Get freeFileMap, , tempInt '21
        
        Get freeFileMap, , tempInt '22
        Get freeFileMap, , tempInt '23
        Get freeFileMap, , tempInt '24
        Get freeFileMap, , tempInt '25
        Get freeFileMap, , tempInt '26
        Get freeFileMap, , tempInt '27
        Get freeFileMap, , tempInt '28
        Get freeFileMap, , tempInt '29
        Get freeFileMap, , tempInt '30
        Get freeFileMap, , tempInt '31
        Get freeFileMap, , tempInt '32
        Get freeFileMap, , tempInt '33
        Get freeFileMap, , tempInt '34
        Get freeFileMap, , tempInt '35
        Get freeFileMap, , tempInt '36
        
    End With

    '*************************************************************************
    'Formularios de Acciones
    Call ME_modAccionEditor.cargarListaAccionesEditorUsando(freeFileMap)
    '*************************************************************************
    'Load arrays
    For y = 1 To TamanioMapa
        For x = 1 To TamanioMapa
    
            Get freeFileMap, , ByFlags
            
            'MapData(X, Y).Blocked = (ByFlags And bitwisetable(0))
            
            'MapData(X, Y).is_water = (ByFlags And bitwisetable(8))
            
            If ByFlags And bitwisetable(6) Then
                Get freeFileMap, , mapdata(x, y).tile_texture
                Get freeFileMap, , mapdata(x, y).tile_number
            Else
                mapdata(x, y).tile_texture = 0
            End If
            
            If ByFlags And bitwisetable(1) Then
                Get freeFileMap, , mapdata(x, y).Graphic(1).GrhIndex
                InitGrh mapdata(x, y).Graphic(1), mapdata(x, y).Graphic(1).GrhIndex
            Else
                mapdata(x, y).Graphic(1).GrhIndex = 0
            End If
                
            If ByFlags And bitwisetable(2) Then
                Get freeFileMap, , mapdata(x, y).Graphic(2).GrhIndex
                InitGrh mapdata(x, y).Graphic(2), mapdata(x, y).Graphic(2).GrhIndex
            Else
                mapdata(x, y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And bitwisetable(3) Then
                Get freeFileMap, , mapdata(x, y).Graphic(3).GrhIndex
                InitGrh mapdata(x, y).Graphic(3), mapdata(x, y).Graphic(3).GrhIndex
            Else
                mapdata(x, y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And bitwisetable(4) Then
                Get freeFileMap, , mapdata(x, y).Graphic(4).GrhIndex
                InitGrh mapdata(x, y).Graphic(4), mapdata(x, y).Graphic(4).GrhIndex
            Else
                mapdata(x, y).Graphic(4).GrhIndex = 0
            End If

            'Trigger used?
            If ByFlags And bitwisetable(5) Then
                Get freeFileMap, , mapdata(x, y).Trigger
            Else
                mapdata(x, y).Trigger = 0
            End If
            
            If ByFlags And bitwisetable(9) Then
                    Get freeFileMap, , NormalData(x, y)
                    Get freeFileMap, , Alturas(x, y)
                    Get freeFileMap, , AlturaPie(x, y)
                    Get freeFileMap, , plusa
                    Get freeFileMap, , mapdata(x, y).tile_orientation
                    hMapData(x, y).h = Alturas(x, y)
                    hMapData(x, y).hs(0) = plusa
            End If
            
            Dim tempintMarce As Integer
            
            If ByFlags And bitwisetable(10) Then
                Get freeFileMap, , tempintMarce
                Set mapdata(x, y).accion = ME_modAccionEditor.obtenerAccionID(tempintMarce)
            End If
    
            If ByFlags And bitwisetable(11) Then
                'Get and make NPC
                Get freeFileMap, , mapdata(x, y).NpcIndex

    
                If mapdata(x, y).NpcIndex < 0 Then
                    mapdata(x, y).NpcIndex = 0
                Else
                    body = NpcData(mapdata(x, y).NpcIndex).body
                    Head = NpcData(mapdata(x, y).NpcIndex).Head
                    heading = NpcData(mapdata(x, y).NpcIndex).heading
                    Dim Char As Integer
                    
                    'Creo el NPC
                    Char = SV_Simulador.NextOpenChar(True)
                    
                    Call MakeChar(Char, body, Head, heading, x, y, 0, 0, 0)

                    CharList(Char).active = 1
                    
                    CharMap(x, y) = Char
                                        
                End If
            End If
    
            If ByFlags And bitwisetable(12) Then
                'Get and make Object
                Get freeFileMap, , mapdata(x, y).OBJInfo.objIndex
                Get freeFileMap, , mapdata(x, y).OBJInfo.Amount
                If mapdata(x, y).OBJInfo.objIndex > 0 Then
                    InitGrh mapdata(x, y).ObjGrh, ObjData(mapdata(x, y).OBJInfo.objIndex).GrhIndex
                End If
            End If
            
            If ByFlags And bitwisetable(13) Then
                Dim luz_r As Byte, luz_g As Byte, luz_b As Byte
                Get freeFileMap, , luz_b
                Get freeFileMap, , luz_g
                Get freeFileMap, , luz_r
                Get freeFileMap, , LRange
                Get freeFileMap, , LBrillo
                Get freeFileMap, , LTipo
                Get freeFileMap, , luzInicio
                Get freeFileMap, , luzFin
                If EsLuzValida(LRange, LBrillo, LTipo) Then
                    mapdata(x, y).luz = DLL_Luces.crear(x, y, luz_r, luz_g, luz_b, LRange, LBrillo, LTipo, luzInicio, luzFin)
                End If
            Else
                mapdata(x, y).luz = 0
            End If
            
            '9)  Particula 0
            If ByFlags And bitwisetable(7) Then
                Set mapdata(x, y).Particles_groups(0) = New Engine_Particle_Group
                
                If mapdata(x, y).Particles_groups(0).Cargar(freeFileMap) Then
                    mapdata(x, y).Particles_groups(0).SetPos x, y
                Else
                    Set mapdata(x, y).Particles_groups(0) = Nothing
                End If
            End If
            
            '10) Particula 1
            If ByFlags And bitwisetable(14) Then
                Set mapdata(x, y).Particles_groups(1) = New Engine_Particle_Group
                
                If mapdata(x, y).Particles_groups(1).Cargar(freeFileMap) Then
                    mapdata(x, y).Particles_groups(1).SetPos x, y
                Else
                    Set mapdata(x, y).Particles_groups(1) = Nothing
                End If
            End If
            
            '8) Particula 2
            If ByFlags And bitwisetable(8) Then
                Set mapdata(x, y).Particles_groups(2) = New Engine_Particle_Group
                
                If mapdata(x, y).Particles_groups(2).Cargar(freeFileMap) Then
                    mapdata(x, y).Particles_groups(2).SetPos x, y
                Else
                    Set mapdata(x, y).Particles_groups(2) = Nothing
                End If
            End If
            
        Next x
    Next y
    
    'Close files


    
    ActualizarArraysAlturasMapas
    
    frmMain.setValoresAgua

    mapinfo.MaxGrhSizeXInTiles = ResizeBackBufferX \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
    mapinfo.MaxGrhSizeYInTiles = ResizeBackBufferY \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
  
    Engine_Set_TileBuffer_Size mapinfo.MaxGrhSizeXInTiles, mapinfo.MaxGrhSizeYInTiles

    Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)
    
    'Heightmap_Calculate -63, -63, 0
Compute_Mountain
    'If cron_tiempo = False Then
        cron_tiempo
        Light_Update_Map = True
        Light_Update_Sombras = True
        
        'map_render_light
    'End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    THIS_MAPA.editado = False
    
    ReDim mapinfo.ZonasNacCriaturas(0) As ZonaNacimientoCriatura
    Call Me_Tools_Npc.cargarZonasDeNacimiento(mapinfo.ZonasNacCriaturas)
    
End Sub
Public Sub Cargar_Mapa_ME(ByVal map As String, Optional Offset As Long = 1)
    Dim freeFileMap As Integer
    Dim tmpHeader As String * 16
    
    AbriendoMapa = True

    If Offset = 0 Then Offset = 1
  
    Call LIMPIAR_MAPA

    freeFileMap = FreeFile
    
    Open map For Binary As freeFileMap
        Seek freeFileMap, Offset
        Get freeFileMap, , tmpHeader
            
        ' Dependiendo el header es la funcion que utilizo
        Select Case tmpHeader
            Case header_m_0
                Call Cargar_Mapa_Me_0(freeFileMap)
            Case header_m_1
                Call Cargar_Mapa_Me_1(freeFileMap)
            Case header_m_2
                Call Cargar_Mapa_Me_2(freeFileMap)
            Case header_m_3
                Call Cargar_Mapa_Me_3(freeFileMap)
                
        End Select
       
        
    Close #freeFileMap

    
    AbriendoMapa = False
End Sub

Private Sub Cargar_Mapa_Me_3(freeFileMap As Integer)
    Dim tempbyte As Byte
    Dim body As Integer
    Dim Head As Integer
    Dim heading As Byte
    Dim y As Integer
    Dim x As Integer
    Dim ByFlags As Integer
    

    Dim MapaNumeroOriginal As Integer
    Dim ResizeBackBufferX As Integer
    Dim ResizeBackBufferY As Integer
    
    Dim plusa As Integer
    
    Dim tempLuz As tLuzPropiedades
    Dim tempintMarce As Integer
    Dim Char As Integer
    
    
    Dim checkSum As String * 10
    
    ' Numero de Mapa
    Get freeFileMap, , MapaNumeroOriginal
    
    THIS_MAPA.numero = MapaNumeroOriginal
        
    '###################################################################################
    '########### PROPIEDADES
    
    Get freeFileMap, , ResizeBackBufferX
    Get freeFileMap, , ResizeBackBufferY

    '###################################################################################
    '########### COLORES Y LUCES!
    
    Get freeFileMap, , OriginalMapColor
    Get freeFileMap, , Intensidad_Del_Terreno

    With mapinfo
        
        Get freeFileMap, , .BaseColor
        Get freeFileMap, , .ColorPropio
        
        Get freeFileMap, , .agua_tileset
        Get freeFileMap, , .agua_rect
        Get freeFileMap, , .agua_profundidad
                
        Get freeFileMap, , .UsaAguatierra
        
    End With

    '*************************************************************************
    'Formularios de Acciones
    Call ME_modAccionEditor.cargarListaAccionesEditorUsando(freeFileMap)
    '*************************************************************************
    ' Zonas
    Get freeFileMap, , tempbyte
    
    ReDim mapinfo.ZonasNacCriaturas(tempbyte) As ZonaNacimientoCriatura
    
    Dim loopZona As Byte
        
    For loopZona = 0 To tempbyte - 1
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).Superior
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).Inferior
        mapinfo.ZonasNacCriaturas(loopZona).nombre = Space$(15)
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).nombre
        mapinfo.ZonasNacCriaturas(loopZona).nombre = Trim$(mapinfo.ZonasNacCriaturas(loopZona).nombre)
    Next
    '************************************************************************
    
    For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
    
            Get freeFileMap, , ByFlags
            
            ' 6) TileSet
            If ByFlags And bitwisetable(6) Then
                Get freeFileMap, , mapdata(x, y).tile_texture
                Get freeFileMap, , mapdata(x, y).tile_number
            Else
                mapdata(x, y).tile_texture = 0
            End If
            
            '1) Capa 1
            If ByFlags And bitwisetable(1) Then
                Get freeFileMap, , mapdata(x, y).Graphic(1).GrhIndex
                InitGrh mapdata(x, y).Graphic(1), mapdata(x, y).Graphic(1).GrhIndex
            Else
                mapdata(x, y).Graphic(1).GrhIndex = 0
            End If
            
            '2) Capa 2
            If ByFlags And bitwisetable(2) Then
                Get freeFileMap, , mapdata(x, y).Graphic(2).GrhIndex
                InitGrh mapdata(x, y).Graphic(2), mapdata(x, y).Graphic(2).GrhIndex
            Else
                mapdata(x, y).Graphic(2).GrhIndex = 0
            End If
                
            '3) Capa 3
            If ByFlags And bitwisetable(3) Then
                Get freeFileMap, , mapdata(x, y).Graphic(3).GrhIndex
                InitGrh mapdata(x, y).Graphic(3), mapdata(x, y).Graphic(3).GrhIndex
            Else
                mapdata(x, y).Graphic(3).GrhIndex = 0
            End If
                
            '4) Capa 4
            If ByFlags And bitwisetable(4) Then
                Get freeFileMap, , mapdata(x, y).Graphic(4).GrhIndex
                InitGrh mapdata(x, y).Graphic(4), mapdata(x, y).Graphic(4).GrhIndex
            Else
                mapdata(x, y).Graphic(4).GrhIndex = 0
            End If
            
            '0) Capa 5
            If ByFlags And bitwisetable(0) Then
                Get freeFileMap, , mapdata(x, y).Graphic(5).GrhIndex
                InitGrh mapdata(x, y).Graphic(5), mapdata(x, y).Graphic(5).GrhIndex
            Else
                mapdata(x, y).Graphic(5).GrhIndex = 0
            End If

            '5) Trigger
            If ByFlags And bitwisetable(5) Then
                Get freeFileMap, , mapdata(x, y).Trigger
            Else
                mapdata(x, y).Trigger = 0
            End If
            
            '9) Altura
            If ByFlags And bitwisetable(9) Then
                    Get freeFileMap, , NormalData(x, y)
                    Get freeFileMap, , Alturas(x, y)
                    Get freeFileMap, , AlturaPie(x, y)
                    Get freeFileMap, , plusa
                    Get freeFileMap, , mapdata(x, y).tile_orientation
                    hMapData(x, y).h = Alturas(x, y)
                    hMapData(x, y).hs(0) = plusa
            End If
            
            '10) Acciones
            If ByFlags And bitwisetable(10) Then
                Get freeFileMap, , tempintMarce
                Set mapdata(x, y).accion = ME_modAccionEditor.obtenerAccionID(tempintMarce)
            End If
    
            '11)
            If ByFlags And bitwisetable(11) Then
            
                ' Criatura Index
                Get freeFileMap, , mapdata(x, y).NpcIndex

                If mapdata(x, y).NpcIndex < 0 Then
                    mapdata(x, y).NpcIndex = 0
                Else
                    body = NpcData(mapdata(x, y).NpcIndex).body
                    Head = NpcData(mapdata(x, y).NpcIndex).Head
                    heading = NpcData(mapdata(x, y).NpcIndex).heading
                    
                    
                    'Creo el NPC
                    Char = SV_Simulador.NextOpenChar(True)
                    
                    Call MakeChar(Char, body, Head, heading, x, y, 0, 0, 0)

                    CharList(Char).active = 1
                    
                    CharMap(x, y) = Char
                                        
                End If
            End If
        
            '12) Objetos
            If ByFlags And bitwisetable(12) Then
            
                Get freeFileMap, , mapdata(x, y).OBJInfo.objIndex
                Get freeFileMap, , mapdata(x, y).OBJInfo.Amount
                
                If mapdata(x, y).OBJInfo.objIndex > 0 Then
                    InitGrh mapdata(x, y).ObjGrh, ObjData(mapdata(x, y).OBJInfo.objIndex).GrhIndex
                End If
                
            End If
            
            '13) Luz
            If ByFlags And bitwisetable(13) Then

                Get freeFileMap, , tempLuz.LuzColor.b
                Get freeFileMap, , tempLuz.LuzColor.g
                Get freeFileMap, , tempLuz.LuzColor.r
                Get freeFileMap, , tempLuz.LuzRadio
                Get freeFileMap, , tempLuz.LuzBrillo
                Get freeFileMap, , tempLuz.LuzTipo
                Get freeFileMap, , tempLuz.luzInicio
                Get freeFileMap, , tempLuz.luzFin
                
                If EsLuzValida(tempLuz.LuzRadio, tempLuz.LuzBrillo, tempLuz.LuzTipo) Then
                    mapdata(x, y).luz = DLL_Luces.crear(x, y, tempLuz.LuzColor.r, tempLuz.LuzColor.g, tempLuz.LuzColor.b, tempLuz.LuzRadio, tempLuz.LuzBrillo, tempLuz.LuzTipo, tempLuz.luzInicio, tempLuz.luzFin)
                End If
                
            Else
                mapdata(x, y).luz = 0
            End If
            
            '9)  Particula 0
            If ByFlags And bitwisetable(7) Then
                Set mapdata(x, y).Particles_groups(0) = New Engine_Particle_Group
                
                If mapdata(x, y).Particles_groups(0).Cargar(freeFileMap) Then
                    mapdata(x, y).Particles_groups(0).SetPos x, y
                Else
                    Set mapdata(x, y).Particles_groups(0) = Nothing
                End If
            End If
            
            '10) Particula 1
            If ByFlags And bitwisetable(14) Then
                Set mapdata(x, y).Particles_groups(1) = New Engine_Particle_Group
                
                If mapdata(x, y).Particles_groups(1).Cargar(freeFileMap) Then
                    mapdata(x, y).Particles_groups(1).SetPos x, y
                Else
                    Set mapdata(x, y).Particles_groups(1) = Nothing
                End If
            End If
            
            '8) Particula 2
            If ByFlags And bitwisetable(8) Then
                Set mapdata(x, y).Particles_groups(2) = New Engine_Particle_Group
                
                If mapdata(x, y).Particles_groups(2).Cargar(freeFileMap) Then
                    mapdata(x, y).Particles_groups(2).SetPos x, y
                Else
                    Set mapdata(x, y).Particles_groups(2) = Nothing
                End If
            End If
            
            ' CheckSum (?)
            Get freeFileMap, , checkSum
            
            If Not Trim(checkSum) = "hola" Then
                Call MsgBox("Error al cargar el mapa. El CheckSum no corresponde. Tile " & x & ";" & y, vbCritical)
                Exit Sub
            End If
        Next x
    Next y
    
    'Close files
    ActualizarArraysAlturasMapas
    
    frmMain.setValoresAgua

    mapinfo.MaxGrhSizeXInTiles = ResizeBackBufferX \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
    mapinfo.MaxGrhSizeYInTiles = ResizeBackBufferY \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
  
    Engine_Set_TileBuffer_Size mapinfo.MaxGrhSizeXInTiles, mapinfo.MaxGrhSizeYInTiles

    Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)
    
    'Heightmap_Calculate -63, -63, 0
    Compute_Mountain
    'If cron_tiempo = False Then
        cron_tiempo
        Light_Update_Map = True
        Light_Update_Sombras = True
        
        'map_render_light
    'End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    THIS_MAPA.editado = False
    
    Call Me_Tools_Npc.cargarZonasDeNacimiento(mapinfo.ZonasNacCriaturas)
End Sub

Private Sub Cargar_Mapa_Me_2(freeFileMap As Integer)
    Dim tempbyte As Byte
    Dim body As Integer
    Dim Head As Integer
    Dim heading As Byte
    Dim y As Integer
    Dim x As Integer
    Dim ByFlags As Integer

    Dim NombreMapa As String * 32
    Dim MapaNumeroOriginal As Integer
    Dim ResizeBackBufferX As Integer
    Dim ResizeBackBufferY As Integer
    
    Dim plusa As Integer
    
    Dim LRange      As Byte
    Dim LBrillo     As Byte
    Dim LTipo       As Integer
    Dim luzInicio As Byte
    Dim luzFin As Byte
    
    'LUCES
    ReDim Intensidad_Del_Terrenob(1 To 218, 1 To 218) As Byte               'Guarda la intensidad de la luz de un vertice del mapa
    ReDim OriginalMapColorb(1 To 218, 1 To 218) As BGRACOLOR_DLL            'Colores precalculados en el mapeditor
    ReDim OriginalMapColorSombraB(1 To 218, 1 To 218) As Long              'OriginalMapColor * Sombra
    ReDim OriginalColorArrayB(1 To 218, 1 To 218) As Long                   'BACKUP DE ResultColorArray (OriginalMapColorSombra * AMBIENTE)
    ReDim ResultColorArrayB(1 To 218, 1 To 218) As Long                    'OriginalColorArray * LUCES DINÁMICAS * SOMBRAS
    
    'Altura de cada vertice del mapa
    ReDim hmapdatab(1 To 218, 1 To 218) As AUDT
    ' Altura de donde pisa el pj, o donde flotan las cosas, o donde el árbol vuela. Sirve para ahcer escaleras.
    ReDim AlturaPieb(1 To 218, 1 To 218) As Integer
    'Es >0 si en la tile hay una altura distnta a cero.
    ReDim Alturasb(1 To 218, 1 To 218) As Integer
    ' Almacena el vector normalizado de los triángulos del mapa, para calcular la sombra Intensidad_sombra = DOT(NORMALIZED•SOL_POS)
    ReDim NormalDatab(1 To 218, 1 To 218) As D3DVECTOR
    ReDim Sombra_MontañasB(1 To 218, 1 To 218) As Byte
    ReDim MapBoxesB(1 To 218, 1 To 218) As Box_Vertex
    '/MONTAÑAS
    
    ReDim mapdatab(1 To 218, 1 To 218) As MapBlock
    
    Get freeFileMap, , MapaNumeroOriginal
    
    THIS_MAPA.numero = MapaNumeroOriginal
        
    
    '###################################################################################
    '########### PROPIEDADES
    
    Get freeFileMap, , ResizeBackBufferX
    Get freeFileMap, , ResizeBackBufferY

    '###################################################################################
    '########### COLORES Y LUCES!
    
    Get freeFileMap, , OriginalMapColorb
    Get freeFileMap, , Intensidad_Del_Terrenob

    With mapinfo
        
        Get freeFileMap, , .BaseColor
        Get freeFileMap, , .ColorPropio
        
        Get freeFileMap, , .agua_tileset
        Get freeFileMap, , .agua_rect
        Get freeFileMap, , .agua_profundidad
                
        Get freeFileMap, , .UsaAguatierra
        
    End With

    '*************************************************************************
    'Formularios de Acciones
    Call ME_modAccionEditor.cargarListaAccionesEditorUsando(freeFileMap)
    '*************************************************************************
    ' Zonas
    Get freeFileMap, , tempbyte
    
    ReDim mapinfo.ZonasNacCriaturas(tempbyte) As ZonaNacimientoCriatura
    
    Dim loopZona As Byte
        
    For loopZona = 0 To tempbyte - 1
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).Superior
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).Inferior
        mapinfo.ZonasNacCriaturas(loopZona).nombre = Space$(15)
        Get freeFileMap, , mapinfo.ZonasNacCriaturas(loopZona).nombre
        mapinfo.ZonasNacCriaturas(loopZona).nombre = Trim$(mapinfo.ZonasNacCriaturas(loopZona).nombre)
    Next
    '*************************************************************************
    
    
    'Load arrays
    For y = 1 To 218
        For x = 1 To 218
    
            Get freeFileMap, , ByFlags
            
            'MapData(X, Y).Blocked = (ByFlags And bitwisetable(0))
            
            'MapData(X, Y).is_water = (ByFlags And bitwisetable(8))
            
            If ByFlags And bitwisetable(6) Then
                Get freeFileMap, , mapdatab(x, y).tile_texture
                Get freeFileMap, , mapdatab(x, y).tile_number
            Else
                mapdatab(x, y).tile_texture = 0
            End If
            
            If ByFlags And bitwisetable(1) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(1).GrhIndex
                InitGrh mapdatab(x, y).Graphic(1), mapdatab(x, y).Graphic(1).GrhIndex
            Else
                mapdatab(x, y).Graphic(1).GrhIndex = 0
            End If
                
            If ByFlags And bitwisetable(2) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(2).GrhIndex
                InitGrh mapdatab(x, y).Graphic(2), mapdatab(x, y).Graphic(2).GrhIndex
            Else
                mapdatab(x, y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And bitwisetable(3) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(3).GrhIndex
                InitGrh mapdatab(x, y).Graphic(3), mapdatab(x, y).Graphic(3).GrhIndex
            Else
                mapdatab(x, y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And bitwisetable(4) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(4).GrhIndex
                InitGrh mapdatab(x, y).Graphic(4), mapdatab(x, y).Graphic(4).GrhIndex
            Else
                mapdatab(x, y).Graphic(4).GrhIndex = 0
            End If
            
            'Layer 5 used?
            If ByFlags And bitwisetable(0) Then
                Get freeFileMap, , mapdatab(x, y).Graphic(5).GrhIndex
                InitGrh mapdatab(x, y).Graphic(5), mapdatab(x, y).Graphic(5).GrhIndex
            Else
                mapdatab(x, y).Graphic(5).GrhIndex = 0
            End If

            'Trigger used?
            If ByFlags And bitwisetable(5) Then
                Get freeFileMap, , mapdatab(x, y).Trigger
            Else
                mapdatab(x, y).Trigger = 0
            End If
            
            If ByFlags And bitwisetable(9) Then
                    Get freeFileMap, , NormalDatab(x, y)
                    Get freeFileMap, , Alturasb(x, y)
                    Get freeFileMap, , AlturaPieb(x, y)
                    Get freeFileMap, , plusa
                    Get freeFileMap, , mapdatab(x, y).tile_orientation
                    hmapdatab(x, y).h = Alturas(x, y)
                    hmapdatab(x, y).hs(0) = plusa
            End If
            
            Dim tempintMarce As Integer
            
            If ByFlags And bitwisetable(10) Then
                Get freeFileMap, , tempintMarce
                Set mapdatab(x, y).accion = ME_modAccionEditor.obtenerAccionID(tempintMarce)
            End If
    
            If ByFlags And bitwisetable(11) Then
                'Get and make NPC
                Get freeFileMap, , mapdatab(x, y).NpcIndex

    
                If mapdatab(x, y).NpcIndex < 0 Then
                    mapdatab(x, y).NpcIndex = 0
                Else
                    body = NpcData(mapdatab(x, y).NpcIndex).body
                    Head = NpcData(mapdatab(x, y).NpcIndex).Head
                    heading = NpcData(mapdatab(x, y).NpcIndex).heading
                    Dim Char As Integer
                    
                    'Creo el NPC
                    Char = SV_Simulador.NextOpenChar(True)
                    
                    Call MakeChar(Char, body, Head, heading, x, y, 0, 0, 0)

                    CharList(Char).active = 1
                    
                    CharMap(x, y) = Char
                                        
                End If
            End If
    
            If ByFlags And bitwisetable(12) Then
                'Get and make Object
                Get freeFileMap, , mapdatab(x, y).OBJInfo.objIndex
                Get freeFileMap, , mapdatab(x, y).OBJInfo.Amount
                If mapdatab(x, y).OBJInfo.objIndex > 0 Then
                    InitGrh mapdatab(x, y).ObjGrh, ObjData(mapdatab(x, y).OBJInfo.objIndex).GrhIndex
                End If
            End If
            
            If ByFlags And bitwisetable(13) Then
                Dim luz_r As Byte, luz_g As Byte, luz_b As Byte
                Get freeFileMap, , luz_b
                Get freeFileMap, , luz_g
                Get freeFileMap, , luz_r
                Get freeFileMap, , LRange
                Get freeFileMap, , LBrillo
                Get freeFileMap, , LTipo
                Get freeFileMap, , luzInicio
                Get freeFileMap, , luzFin
                If EsLuzValida(LRange, LBrillo, LTipo) Then
                    mapdatab(x, y).luz = DLL_Luces.crear(x, y, luz_r, luz_g, luz_b, LRange, LBrillo, LTipo, luzInicio, luzFin)
                End If
            Else
                mapdatab(x, y).luz = 0
            End If
            
            '9)  Particula 0
            If ByFlags And bitwisetable(7) Then
                Set mapdatab(x, y).Particles_groups(0) = New Engine_Particle_Group
                
                If mapdatab(x, y).Particles_groups(0).Cargar(freeFileMap) Then
                    mapdatab(x, y).Particles_groups(0).SetPos x, y
                Else
                    Set mapdatab(x, y).Particles_groups(0) = Nothing
                End If
            End If
            
            '10) Particula 1
            If ByFlags And bitwisetable(14) Then
                Set mapdatab(x, y).Particles_groups(1) = New Engine_Particle_Group
                
                If mapdatab(x, y).Particles_groups(1).Cargar(freeFileMap) Then
                    mapdatab(x, y).Particles_groups(1).SetPos x, y
                Else
                    Set mapdatab(x, y).Particles_groups(1) = Nothing
                End If
            End If
            
            '8) Particula 2
            If ByFlags And bitwisetable(8) Then
                Set mapdatab(x, y).Particles_groups(2) = New Engine_Particle_Group
                
                If mapdatab(x, y).Particles_groups(2).Cargar(freeFileMap) Then
                    mapdatab(x, y).Particles_groups(2).SetPos x, y
                Else
                    Set mapdatab(x, y).Particles_groups(2) = Nothing
                End If
            End If
            
            ' CheckSum (?)
            Dim marceCheck As String * 10
            
            Get freeFileMap, , marceCheck
            
            If Not Trim(marceCheck) = "hola" Then
                Call MsgBox("Error al cargar el mapa. El CheckSum no corresponde. Tile " & x & ";" & y, vbCritical)
                Exit Sub
            End If
        Next x
    Next y
    
    'Close files

    For x = SV_Constantes.X_MINIMO_VISIBLE To SV_Constantes.X_MAXIMO_VISIBLE
        For y = SV_Constantes.Y_MINIMO_JUGABLE To SV_Constantes.Y_MAXIMO_JUGABLE
            mapdata(x, y) = mapdatab(x, y)
        Next y
    Next x
   
    ActualizarArraysAlturasMapas
    
    frmMain.setValoresAgua

    mapinfo.MaxGrhSizeXInTiles = ResizeBackBufferX \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
    mapinfo.MaxGrhSizeYInTiles = ResizeBackBufferY \ 32 + 1 'FIXME HARDCODEADO, GUARDAR EN EL ARCHIVO DE MAPA
  
    Engine_Set_TileBuffer_Size mapinfo.MaxGrhSizeXInTiles, mapinfo.MaxGrhSizeYInTiles

    Call DXCopyMemory(OriginalMapColorSombra(1, 1), OriginalMapColor(1, 1), TILES_POR_MAPA * 4)
    
    'Heightmap_Calculate -63, -63, 0
    Compute_Mountain
    'If cron_tiempo = False Then
        cron_tiempo
        Light_Update_Map = True
        Light_Update_Sombras = True
        
        'map_render_light
    'End If
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    THIS_MAPA.editado = False
    
    Call Me_Tools_Npc.cargarZonasDeNacimiento(mapinfo.ZonasNacCriaturas)

End Sub

Public Sub LIMPIAR_MAPA()
    Dim y As Long
    Dim x As Long
    Dim loopCapa As Byte
    
    FX_Projectile_Erase_All

    ZeroMemory Alturas(1, 1), TILES_POR_MAPA * 2
    'ZeroMemory ModSuperWater(1, 1), TILES_POR_MAPA
    ZeroMemory ModSuperWaterMM(1, 1), LenB(ModSuperWaterMM(1, 1)) * TILES_POR_MAPA
    ZeroMemory ModSuperWaterDD(1, 1), LenB(ModSuperWaterDD(1, 1)) * TILES_POR_MAPA
    ZeroMemory hMapData(1, 1), LenB(hMapData(1, 1)) * TILES_POR_MAPA
    ZeroMemory AlturaPie(1, 1), 2 * TILES_POR_MAPA
    
    'Elimino las luces
    Clear_Luces_Mapa
    
    'Reiniciamos el mapa.
    For y = 1 To SV_Constantes.ALTO_MAPA
        For x = 1 To SV_Constantes.ANCHO_MAPA
            With mapdata(x, y)
                '.Blocked = 0
                .is_water = 0
                
                For loopCapa = 1 To CANTIDAD_CAPAS
                    .Graphic(loopCapa).GrhIndex = 0
                Next loopCapa
                
                .ObjGrh.GrhIndex = 0
                .Trigger = 0
                .tile_texture = 0
                .tile_number = 0
                
                'Quito la referencia a la luz
                .luz = 0
                
                Set .accion = Nothing
                
                .NpcIndex = 0
                                
                .OBJInfo.objIndex = 0
                
                Set .Particles_groups(0) = Nothing
                Set .Particles_groups(1) = Nothing
                Set .Particles_groups(2) = Nothing


                If CharMap(x, y) And CharMap(x, y) <> UserCharIndex Then
                    Call SV_Simulador.EraseIndexChar(CharMap(x, y))
                    Call EraseChar(CharMap(x, y))
                    CharMap(x, y) = 0
                    .Charindex = 0
                End If
                
                Do While EntidadesMap(x, y) > 0
                   Call Engine_Entidades.eliminar(EntidadesMap(x, y))
                Loop
                
                
            End With
            
          
        Next x
    Next y
    
    ReDim mapinfo.ZonasNacCriaturas(0) As ZonaNacimientoCriatura
    mapinfo.ZonasNacCriaturas(0).nombre = ""
    
    mapinfo.UsaAguatierra = False
    mapinfo.agua_tileset = 19
    
    mapinfo.ColorPropio = False
    mapinfo.BaseColor.r = 255
    mapinfo.BaseColor.g = 255
    mapinfo.BaseColor.b = 255
    
    Compute_Mountain
    
    'Engine_Map.Map_render_2array
    rm2a
    
End Sub


Function CompilarMapa() As Boolean
'Marce On error resume next
    Dim salida As String

    If THIS_MAPA.numero = 0 Then
        MsgBox "No podes compilar un mapa con Num=0"
        Exit Function
    End If
    
    '************ Cliente
    If pakMapas Is Nothing Then
        MsgBox "EL enpaquetado de mapas del cliente tiene un formato incorrecto. Borrelo para hacer uno nuevo y reinicie el editor de mapas."
        LogDebug "NO Se guardó el mapa de cliente numero " & THIS_MAPA.numero & " PORQUE NO ESTABA CARGADO EL PAK DE MAPAS"
        Exit Function
    End If
    
    salida = OPath & "Mapas\Cliente\" & THIS_MAPA.numero & ".tdsmap"
    
    If Not pakMapas Is Nothing Then
        If Guardar_Mapa_CLI(salida) Then
            pakMapas.Parchear THIS_MAPA.numero, salida
            LogDebug "Se guardó el mapa de cliente numero " & THIS_MAPA.numero
        Else
            LogDebug "NO Se parcheó el mapa de cliente numero " & THIS_MAPA.numero
        End If
    End If
    
    '************ Server
    salida = OPath & "Mapas\Servidor\" & THIS_MAPA.numero & ".servermap"
    
    If Guardar_Mapa_SV(salida) Then
        LogDebug "Se guardó el mapa del server numero " & THIS_MAPA.numero
    Else
        LogDebug "No Se guardó el mapa de server numero " & THIS_MAPA.numero
        MsgBox "Fallo al querer intentar guardar el mapa en formato SERVER.", vbCritical
    End If
End Function

Sub NuevoMapa()
    Call LIMPIAR_MAPA
    
    Call ME_FIFO.prepararWorkEspace
    
    THIS_MAPA.editado = False
    THIS_MAPA.Path = ""
    THIS_MAPA.numero = 0
    THIS_MAPA.nombre = "Mapa sin nombre"
    
    RemakeWaterTilenumbers 0, 0, 3, 3
        
    frmMain.setValoresAgua
    
    Call frmMain.act_titulo
End Sub

Public Sub prepararWorkEspace()

    Call ME_Tools.iniciar
        
    ' Menu del portapapeles
    Dim loopPorta As Byte
    
    If UBound(portapapeles) > frmMain.mnuPortapapeles.count Then
        For loopPorta = frmMain.mnuPortapapeles.count + 1 To UBound(portapapeles)
            load frmMain.mnuPortapapeles(loopPorta)
            frmMain.mnuPortapapeles.item(loopPorta).caption = loopPorta & ": < Vacio >"
        Next
    End If
    
    'Inicio las estructuras del hacer/deshacer
    Call ME_modComandos.iniciarDesHacerReHacer
    
    'Lista de acciones que esoty usando... en 0.
    Call ME_modAccionEditor.iniciar
    Call ME_modAccionEditor.refrescarListaUsando(frmMain.listTileAccionActuales)

    ' Lista de areas
    Call Me_Tools_Npc.cargarZonasDeNacimiento(mapinfo.ZonasNacCriaturas)
        
End Sub
