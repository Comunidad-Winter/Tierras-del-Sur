Attribute VB_Name = "ME_Mundo"
Option Explicit

Public Type MapaS
    numero As Integer
    Color As Long
    Puedo As Byte
    existe As Byte
End Type

Public Enum ePuntoCardinal
    CENTRO = 0
    OESTE = 1
    NORTE = 2
    ESTE = 3
    SUR = 4
    SUROESTE = 5
    NOROESTE = 6
    NORESTE = 7
    SURESTE = 8
End Enum

Public Type coordenadaMundo  'Coordenadas del mapa dentro del mundo
    x As Integer
    y As Integer
End Type

Private Type tZonaDisponible
    nombre As String
    archivo As String
End Type

Private zonaActual As Byte

Public zonas() As tZonaDisponible
Public CantidadZonasCargadas As Byte

Public MapasArray() As MapaS

Public cantidadMapasX As Integer
Public cantidadMapasY As Integer

Public MapasArrayCargado As Boolean

Public Function obtenerNombreZonaActual() As String
    obtenerNombreZonaActual = zonas(zonaActual).nombre
End Function
Private Function obtenerDatosZona(archivoOrigen As String) As String

    Dim archivo As Integer
    Dim nombreMundo As String
    Dim longitudNombre As Byte
    
    archivo = FreeFile
    Open archivoOrigen For Binary As archivo
    
    'Cargamos el nombre
    Get archivo, , longitudNombre
    nombreMundo = Space$(longitudNombre)
    Get archivo, , nombreMundo
    
    Close #archivo
    
    obtenerDatosZona = nombreMundo
End Function

Public Sub cargarZonasPosibles()
    Dim sFileName As String
    Dim cantidad As Byte
    
    CantidadZonasCargadas = 0
    cantidad = 0
    'Recorremos todos los archivos buscando mundoTDS
    sFileName = Dir$(DatosPath & "Bot\", vbArchive)
    
    Do While sFileName > ""
        If right$(sFileName, 8) = "mundoTDS" Then
            
            ReDim Preserve zonas(cantidad) As tZonaDisponible
            
            zonas(cantidad).nombre = obtenerDatosZona(DatosPath & "Bot\" & sFileName)
            zonas(cantidad).archivo = sFileName
            
            cantidad = cantidad + 1
        End If
        
        sFileName = Dir$()
    Loop
    
    CantidadZonasCargadas = cantidad

End Sub

Public Function obtenerIDZona(ByVal Zona As String) As Integer
    Dim loopZona As Byte
    
    Zona = UCase$(Zona)
    
    For loopZona = LBound(zonas) To UBound(zonas)
        If UCase$(zonas(loopZona).nombre) = Zona Then
            obtenerIDZona = loopZona
            Exit Function
        End If
    Next
    obtenerIDZona = -1
End Function
Public Function existePosicion(x As Integer, y As Integer) As Boolean
    
    If x >= X_MINIMO_VISIBLE And x <= SV_Constantes.X_MAXIMO_VISIBLE _
        And y >= SV_Constantes.Y_MINIMO_VISIBLE And y <= SV_Constantes.Y_MAXIMO_VISIBLE Then
        existePosicion = True
    Else
        existePosicion = False
    End If


End Function

Public Function puedeModificarComporamientoTile(coordX As Integer, coordY As Integer) As Boolean

    If (obtenerMapaDuenioCoordenadas(coordX, coordY) = ePuntoCardinal.CENTRO) Then
        puedeModificarComporamientoTile = True
    Else
        puedeModificarComporamientoTile = False
    End If
    
End Function

' ATENCION. Esta funcion se le debe aplicar a un X,Y que sea puedeModificarComporamientoTile(x,y) = verdadero

Public Function esVisibleEnOtroMapa(coordX As Integer, coordY As Integer) As Boolean
    Dim punto As ePuntoCardinal
    
    esVisibleEnOtroMapa = False
    
    If coordY < Y_MINIMO_NO_VISIBLE_OTRO_MAPA Then  'Hacia el Norte
    
        ' Esta parte se va a ver en dos mapas, en el norte y en el oeste, si es que alguno existe
        If mapaTieneLimitrofe(THIS_MAPA.numero, ePuntoCardinal.NORTE) Then
            esVisibleEnOtroMapa = True
            Exit Function
        End If
            
        If coordX < X_MINIMO_NO_VISIBLE_OTRO_MAPA Then
        
            ' Esta parte se va a ver en dos mapas, en el norte y en el oeste
            If mapaTieneLimitrofe(THIS_MAPA.numero, ePuntoCardinal.OESTE) Then
                esVisibleEnOtroMapa = True
            End If
            
        ElseIf coordX > SV_Constantes.X_MAXIMO_NO_VISIBLE_OTRO_MAPA Then
        
            ' Esta parte se va a ver en dos mapas, en el norte y en el oeste, si es que alguno existe
            If mapaTieneLimitrofe(THIS_MAPA.numero, ePuntoCardinal.ESTE) Then
                esVisibleEnOtroMapa = True
            End If
            
        End If
    ElseIf coordY > Y_MAXIMO_NO_VISIBLE_OTRO_MAPA Then 'Hacia el sur
        
        ' Si o si se va a ver en el sur
        If mapaTieneLimitrofe(THIS_MAPA.numero, ePuntoCardinal.SUR) Then
            esVisibleEnOtroMapa = True
            Exit Function
        End If
        
        ' Y ademas puede tambien verse en el oeste o el este
        If coordX < X_MINIMO_NO_VISIBLE_OTRO_MAPA Then
            
            If mapaTieneLimitrofe(THIS_MAPA.numero, ePuntoCardinal.OESTE) Then
                esVisibleEnOtroMapa = True
            End If
            
        ElseIf coordX > X_MAXIMO_NO_VISIBLE_OTRO_MAPA Then
            
            If mapaTieneLimitrofe(THIS_MAPA.numero, ePuntoCardinal.ESTE) Then
                esVisibleEnOtroMapa = True
            End If
            
        End If
    Else 'Estamos en la linea media
        If coordX < X_MINIMO_NO_VISIBLE_OTRO_MAPA Then
            punto = ePuntoCardinal.OESTE
        ElseIf coordX > X_MAXIMO_NO_VISIBLE_OTRO_MAPA Then
            punto = ePuntoCardinal.ESTE
        Else
            punto = ePuntoCardinal.CENTRO
        End If
        
        If Not punto = ePuntoCardinal.CENTRO Then
            If mapaTieneLimitrofe(THIS_MAPA.numero, punto) Then
                esVisibleEnOtroMapa = True
            Else
                esVisibleEnOtroMapa = False
            End If
        End If
    End If
        
End Function
Public Function puedeModificarAspectoTile(coordX As Integer, coordY As Integer) As Boolean
    
    Dim puntoCardinal As ePuntoCardinal
    
    puntoCardinal = obtenerMapaDuenioCoordenadas(coordX, coordY)
    
    If (obtenerMapaDuenioCoordenadas(coordX, coordY) = ePuntoCardinal.CENTRO) Then
        puedeModificarAspectoTile = True
    Else
        'TODO: Esta variable global THIS_MAPA.numero es fea que este en esta funcion
        If mapaTieneLimitrofe(THIS_MAPA.numero, puntoCardinal) Then
            puedeModificarAspectoTile = False
        Else
            puedeModificarAspectoTile = True
        End If
    End If

End Function
'A partir de una coordenada absoluta devuelvo a que mapa pertenece ese coordenada
Public Function obtenerMapaDuenioCoordenadas(coordX As Integer, coordY As Integer) As ePuntoCardinal

        
        If coordY < Y_MINIMO_USABLE Then 'Hacia el Norte
            If coordX < X_MINIMO_USABLE Then
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.NOROESTE
            ElseIf coordX > X_MAXIMO_USABLE Then
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.NORESTE
            Else
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.NORTE
            End If
        ElseIf coordY > Y_MAXIMO_USABLE Then 'Hacia el sur
            If coordX < X_MINIMO_USABLE Then
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.SUROESTE
            ElseIf coordX > X_MAXIMO_USABLE Then
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.SURESTE
            Else
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.SUR
            End If
        Else 'Estamos en la linea media
            If coordX < X_MINIMO_USABLE Then
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.OESTE
            ElseIf coordX > X_MAXIMO_USABLE Then
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.ESTE
            Else
                obtenerMapaDuenioCoordenadas = ePuntoCardinal.CENTRO
            End If
        End If

End Function

Private Function obtenerMapa(coordX As Integer, coordY As Integer) As Integer
    
    If coordX < 1 Or coordX > cantidadMapasX Or coordY < 1 Or coordY > cantidadMapasY Then
        obtenerMapa = 0
    Else
        obtenerMapa = MapasArray(coordX, coordY).numero
    End If
        
End Function
Public Function obtenerCoordenada(numeroMapa As Integer) As coordenadaMundo

    Dim coordX As Integer
    Dim coordY As Integer
    
    For coordX = 1 To cantidadMapasX
        For coordY = 1 To cantidadMapasY
            If MapasArray(coordX, coordY).numero = numeroMapa Then
                obtenerCoordenada.x = coordX
                obtenerCoordenada.y = coordY
                Exit Function
            End If
        Next coordY
    Next coordX
End Function

'Dado un numero de mapa y un punto cardinal devuelve el mapa que se encuentra ahí o 0 si no existe
Public Function obtenerMapaLimitrofeMapa(numero As Integer, puntoCardinal As ePuntoCardinal) As Integer
        Dim coordenadaMapa As coordenadaMundo
        
        coordenadaMapa = obtenerCoordenada(numero)
        obtenerMapaLimitrofeMapa = obtenerMapaLimitrofe(coordenadaMapa.x, coordenadaMapa.y, puntoCardinal)
End Function

'Dada una coordenada de un mapa dentro del mundo devuelve el mapa que se encuentra en el punto cardinal
Public Function obtenerMapaLimitrofe(ByVal coordXMapa As Integer, ByVal coordYMapa As Integer, puntoCardinal As ePuntoCardinal)
    
    Select Case puntoCardinal
        Case ePuntoCardinal.CENTRO
        Case ePuntoCardinal.ESTE
            coordXMapa = coordXMapa + 1
        Case ePuntoCardinal.NORESTE
            coordXMapa = coordXMapa + 1
            coordYMapa = coordYMapa - 1
        Case ePuntoCardinal.NOROESTE
            coordXMapa = coordXMapa - 1
            coordYMapa = coordYMapa - 1
        Case ePuntoCardinal.NORTE
            coordYMapa = coordYMapa - 1
        Case ePuntoCardinal.OESTE
            coordXMapa = coordXMapa - 1
        Case ePuntoCardinal.SUR
            coordYMapa = coordYMapa + 1
        Case ePuntoCardinal.SURESTE
            coordYMapa = coordYMapa + 1
            coordXMapa = coordXMapa + 1
        Case ePuntoCardinal.SUROESTE
            coordYMapa = coordYMapa + 1
            coordXMapa = coordXMapa - 1
    End Select
    
    obtenerMapaLimitrofe = obtenerMapa(coordXMapa, coordYMapa)
End Function

Public Function mapaTieneLimitrofe(numeroMapa As Integer, puntoCardinal As ePuntoCardinal) As Boolean
    mapaTieneLimitrofe = (obtenerMapaLimitrofeMapa(numeroMapa, puntoCardinal) <> 0)
End Function

Public Sub ActualizarPuedoMapas()
Dim x%, y%
   
        For x = 1 To cantidadMapasX
            For y = 1 To cantidadMapasY
                With MapasArray(x, y)
                    If .numero Then
                        'If pakMapasME.Cabezal_GetFilePtr(.numero) Then
                            .existe = 1
                            .Puedo = True 'pakMapasME.Puedo_Editar(.numero, CDM_UserPrivs, CDM_UserID)
                       ' Else
                       '     .existe = 0
                        '    .Puedo = 0
                        'End If
                    End If
                End With
            Next y
        Next x
    
End Sub


Private Function cargarEstructuraMundo(archivoOrigen As String) As Boolean
    Dim Cantidad_Mapas As Integer
    Dim archivo As Integer
    Dim TempColor As Long
    
    Dim nombreMundo As String
    Dim longitudNombre As Byte
    
    'Una variable de tipo Libro de Excel
    Dim col As Integer, fil As Integer
    
    Cantidad_Mapas = 0

    If FileExist(archivoOrigen, vbArchive) Then
        archivo = FreeFile
        Open archivoOrigen For Binary As archivo
    Else
        MsgBox "El archivo " & archivoOrigen & " no existe."
        cargarEstructuraMundo = False
        Exit Function
    End If
    
    'Cargamos el nombre
    Get archivo, , longitudNombre
    nombreMundo = Space$(longitudNombre)
    Get archivo, , nombreMundo
    
    'Cargamos las dimensiones
    Get archivo, , cantidadMapasX
    Get archivo, , cantidadMapasY
    
    ReDim MapasArray(1 To cantidadMapasX, 1 To cantidadMapasY)
    '
    For fil = 1 To cantidadMapasY
        For col = 1 To cantidadMapasX
            With MapasArray(col, fil)

                Get archivo, , .numero
                Get archivo, , TempColor
                
                If .numero > 0 Then
                    .Color = VBCOLOR2DXCOLOR(TempColor)
                    Cantidad_Mapas = Cantidad_Mapas + 1
                End If
            End With
        Next
    Next

    Close #archivo
    'Retorno
    zonaActual = obtenerIDZona(nombreMundo)
    cargarEstructuraMundo = True
End Function
Public Sub CargarArrayMapas(archivoDatos As String)

If cargarEstructuraMundo(DatosPath & "Bot\" & archivoDatos) Then
    ActualizarPuedoMapas
    MapasArrayCargado = True
End If

End Sub

Public Sub aplicar_traslados(ByRef infoMap As mapinfo, ByRef mapdata() As MapBlock)

    Dim numeroMapa As Integer
    Dim borde As ePuntoCardinal
    Dim posInicial As Position
    Dim posFinal As Position
    Dim numeroMapaLimitrofe As Integer
    Dim accion As cAccionCompuestaEditor
    
    numeroMapa = infoMap.numero
        
    For borde = ePuntoCardinal.OESTE To ePuntoCardinal.SUR
    
        'Obtengo el mapa lindero
        numeroMapaLimitrofe = ME_Mundo.obtenerMapaLimitrofeMapa(numeroMapa, borde)
        
        'Obtengo la linea donde voy a aplicar, ya sea los traslados
        'o el bloqueo
        Select Case borde
        
            Case ePuntoCardinal.NORTE
                posInicial.x = SV_Constantes.X_MINIMO_USABLE
                posInicial.y = SV_Constantes.Y_MINIMO_JUGABLE
                posFinal.x = SV_Constantes.X_MAXIMO_USABLE
                posFinal.y = SV_Constantes.Y_MINIMO_JUGABLE
            Case ePuntoCardinal.SUR
                posInicial.x = SV_Constantes.X_MINIMO_USABLE
                posInicial.y = SV_Constantes.Y_MAXIMO_JUGABLE
                posFinal.x = SV_Constantes.X_MAXIMO_USABLE
                posFinal.y = SV_Constantes.Y_MAXIMO_JUGABLE
            Case ePuntoCardinal.OESTE
                posInicial.x = SV_Constantes.X_MINIMO_JUGABLE
                posInicial.y = SV_Constantes.Y_MINIMO_USABLE
                posFinal.x = SV_Constantes.X_MINIMO_JUGABLE
                posFinal.y = SV_Constantes.Y_MAXIMO_USABLE
            Case ePuntoCardinal.ESTE
                posInicial.x = SV_Constantes.X_MAXIMO_JUGABLE
                posInicial.y = SV_Constantes.Y_MINIMO_USABLE
                posFinal.x = SV_Constantes.X_MAXIMO_JUGABLE
                posFinal.y = SV_Constantes.Y_MAXIMO_USABLE
        End Select
        
        If numeroMapaLimitrofe > 0 Then
            Call desbloquearArea(mapdata, posInicial, posFinal)
            
            Set accion = obtenerAccionExit(numeroMapaLimitrofe, borde)
            Call aplicarAccionArea(mapdata, posInicial, posFinal, accion)
        Else
            Call bloquearArea(mapdata, posInicial, posFinal)
            Call aplicarAccionArea(mapdata, posInicial, posFinal, Nothing)
        End If
    Next
    
End Sub

Private Function obtenerAccionExit(numeroMapa As Integer, borde As ePuntoCardinal) As cAccionCompuestaEditor
    
    Set obtenerAccionExit = ME_modAccionEditor.obtenerAccion("Exit al mapa " & numeroMapa)
    Dim accion As iAccion
    
    
    If obtenerAccionExit Is Nothing Then
        'Tengo que crear la accion
        Dim nuevaAccionPadre As New cAccionCompuestaEditor
        Dim nuevaAccionHijo As New cAccionTileEditor
        
        Select Case borde
            Case ePuntoCardinal.NORTE
                Set accion = New cAccionExitNorte
            Case ePuntoCardinal.SUR
                Set accion = New cAccionExitSur
            Case ePuntoCardinal.ESTE
                Set accion = New cAccionExitEste
            Case ePuntoCardinal.OESTE
                Set accion = New cAccionExitOeste
        End Select
        
        Dim parametro As cParamAccionTileEditor

        Set parametro = ME_constructoresAccionEditor.construirCampoMapa
        
        Call parametro.setValor(CStr(numeroMapa))
        
        Call nuevaAccionHijo.crear("", "", accion)
        Call nuevaAccionHijo.agregarParametro(parametro)
        
        Call nuevaAccionPadre.iAccionEditor_crear("Exit al mapa " & numeroMapa, "")
        Call nuevaAccionPadre.agregarHijo(nuevaAccionHijo)
        Call nuevaAccionPadre.setVisible(False) 'No me interesa que se vea al publico
        Call ME_modAccionEditor.agregarNuevaAccion(nuevaAccionPadre)
        
        
        Set obtenerAccionExit = nuevaAccionPadre
    End If
End Function


Public Sub aplicarAccionArea(mapdata() As MapBlock, posInicial As Position, posFinal As Position, accion As cAccionCompuestaEditor)
    
    Dim x As Integer
    Dim y As Integer
    

    For x = posInicial.x To posFinal.x
        For y = posInicial.y To posFinal.y
            Set mapdata(x, y).accion = accion
        Next y
    Next
    
End Sub
Private Sub desbloquearArea(mapdata() As MapBlock, posInicial As Position, posFinal As Position)
    Dim x As Integer
    Dim y As Integer
    

    For x = posInicial.x To posFinal.x
        For y = posInicial.y To posFinal.y
            If mapdata(x, y).Trigger > 0 Then
                mapdata(x, y).Trigger = (mapdata(x, y).Trigger Xor eTriggers.TodosBordesBloqueados)
            End If
        Next y
    Next

End Sub

Private Sub bloquearArea(mapdata() As MapBlock, posInicial As Position, posFinal As Position)
    Dim x As Integer
    Dim y As Integer
    

    For x = posInicial.x To posFinal.x
        For y = posInicial.y To posFinal.y
            mapdata(x, y).Trigger = (mapdata(x, y).Trigger Or eTriggers.TodosBordesBloqueados)
        Next y
    Next

End Sub

Public Sub actualizarEfectoPisada()

'Dim efecto As Integer
'Dim loopX As Integer
'Dim loopY As Integer
'Dim loopXExpandir As Integer
'Dim loopYExpandir As Integer
'
'' Aqui voy a guardar por cada tile, cual es la capa que tome para aplicar el sonido
'' 0 es el piso
'Dim CapaFuente(SV_Constantes.X_MINIMO_JUGABLE To SV_Constantes.X_MAXIMO_JUGABLE, SV_Constantes.Y_MINIMO_JUGABLE To SV_Constantes.Y_MAXIMO_JUGABLE) As Byte
'
'efecto = 0
'
'For loopX = SV_Constantes.X_MINIMO_JUGABLE To SV_Constantes.X_MAXIMO_JUGABLE
'
'    For loopY = SV_Constantes.Y_MINIMO_JUGABLE To SV_Constantes.Y_MAXIMO_JUGABLE
'
'    With mapdata(loopX, loopY)
'
'        ' En orden de prioridad
'
'        ' Piso
'        If .tile_texture > 0 Then
'
'            If CapaFuente(loopX, loopY) = 0 Then
'                mapdata(loopX, loopY).EfectoPisada = Tilesets(.tile_texture).EfectoPisada(.tile_number)
'            End If
'        End If
'
'        ' Capa 1
'        If .Graphic(1).GrhIndex > 0 Then
'
'            ' Expandimos el analisis a todos los tilesets que abarca el gráfico
'            For loopXExpandir = loopX To loopX + GrhData(.Graphic(1).GrhIndex).TileWidth - 1
'                For loopYExpandir = loopY To loopY + GrhData(.Graphic(1).GrhIndex).TileHeight - 1
'
'                    '¿El efecto que ya tiene, es de una capa de igual o menor prioridad?
'                    If CapaFuente(loopXExpandir, loopYExpandir) <= 1 Then
'                        mapdata(loopXExpandir, loopYExpandir).EfectoPisada = GrhData(.Graphic(1).GrhIndex).EfectoPisada
'                        CapaFuente(loopXExpandir, loopYExpandir) = 1
'                    End If
'
'                Next loopYExpandir
'            Next loopXExpandir
'
'        End If
'
'        ' Capa 2
'        If .Graphic(2).GrhIndex > 0 Then
'
'            ' Expandimos el analisis a todos los tilesets que abarca el gráfico
'            For loopXExpandir = loopX To loopX + GrhData(.Graphic(2).GrhIndex).TileWidth - 1
'                For loopYExpandir = loopY - (GrhData(.Graphic(2).GrhIndex).TileHeight - 1) To loopY
'
'                    '¿El efecto que ya tiene, es de una capa de igual o menor prioridad?
'                    If CapaFuente(loopXExpandir, loopYExpandir) <= 2 Then
'                        mapdata(loopXExpandir, loopYExpandir).EfectoPisada = GrhData(.Graphic(2).GrhIndex).EfectoPisada
'                        CapaFuente(loopXExpandir, loopYExpandir) = 2
'                    End If
'
'                Next loopYExpandir
'            Next loopXExpandir
'        End If
'
'    End With
'
'    Next loopY
'
'Next loopX


End Sub

Private Function VBCOLOR2DXCOLOR(Color As Long) As Long
    Dim c(3) As Byte
    DXCopyMemory c(0), Color, 4
    
    VBCOLOR2DXCOLOR = D3DColorXRGB(c(0), c(1), c(2))
End Function

