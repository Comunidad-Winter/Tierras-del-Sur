Attribute VB_Name = "Me_Tools_TileSet"
Option Explicit

Public Enum eHerramientasTileSet
    ninguna = 0
    insertar = 1
    borrar = 2
End Enum

Public herramientaInternaTileSet As eHerramientasTileSet

Public TilesetSeleccionado() As Integer
Public tilesetNumeroSeleccionado() As Integer
Public TileSetSectorAncho As Byte
Public TileSetSectorAlto As Byte
Public TilesetWindow        As vw_Tileset

' Si esto está activado el Editor le da libertad al usuario para insertar el piso seleccionado
' en cualquier parte del mapa
Public NoForzarInserccionCorrecta As Boolean

Public Sub actualizarListaTileSetDeSeleccion()
        If Not TilesetWindow Is Nothing Then
            TilesetWindow.ActualizarLista
        End If
End Sub
Public Sub MostrarVentanaTilesets(numeroTileSet As Integer, numeroTileSetVirtual As Byte)
    If Not TilesetWindow Is Nothing And TOOL_SELECC And Tools.tool_tileset Then
    
        TilesetWindow.TilesetSeleccionado = numeroTileSet
        TilesetWindow.TileSetVirtualSeleccionado = numeroTileSetVirtual
        
        GUI_SetFocus TilesetWindow
        mostrandoVWindows = True
    End If
End Sub

Public Sub EsconderVentanaTilesets()
    If Not TilesetWindow Is Nothing Then
        GUI_Quitar TilesetWindow
        mostrandoVWindows = False
    End If
End Sub

Public Sub iniciarToolTileSets()
    If TilesetWindow Is Nothing Then
        Set TilesetWindow = New vw_Tileset
    End If
    Call establecerTileSet(0, 0)
End Sub

Public Sub establecerTileSet(tileset As Integer, numero As Integer)

    ReDim TilesetSeleccionado(1 To 1, 1 To 1)
    ReDim tilesetNumeroSeleccionado(1 To 1, 1 To 1)
    
    TilesetSeleccionado(1, 1) = tileset
    tilesetNumeroSeleccionado(1, 1) = numero
End Sub

Public Sub establecerAreaTileSet(tileset As Integer, TileSetVirtual As Byte, area As tAreaSeleccionada)

    Dim formato As eFormatoTileSet ' Formato del TileSet seleccionado
    
    ' Tamaño del Lienzo con el cual vamos a trabajar
    Dim xMaximoArea As Byte
    Dim yMaximoArea As Byte
        
    ' De la posicion original del tile en el tileset vamos a tener que posicionarlo
    ' en el lienzo que vamos a insertar respetando la posicion relativa
    Dim offsetX As Integer
    Dim offsetY As Integer
    
    ' La posicion del tile en el lienzo luego de aplicada la tranformacion
    Dim tileX As Integer
    Dim tileY As Integer
        
    Dim xAux As Byte ' Auxiliares para for
    Dim yAux As Byte
        
    formato = Tilesets(tileset).formato
    
    xMaximoArea = 0
    yMaximoArea = 0
    offsetX = 0
    offsetY = 0
    
    ' ¿Debo forzar a que lo seleccionado tenga determinado tamaño?
    If Not NoForzarInserccionCorrecta Then
        If formato = eFormatoTileSet.textura_simple Or formato = eFormatoTileSet.camino_chico Then
            xMaximoArea = 8
            yMaximoArea = 8
            TileSetSectorAncho = 8
            TileSetSectorAlto = 8
        ElseIf formato = camino_grande_parte2 Then
            ' Sino es Transiciones o es la Interseccion de las Transiciones
            If TileSetVirtual < 5 Or (TileSetVirtual = 5 And area.abajo <= 7) Then
                xMaximoArea = 8
                yMaximoArea = 8
                TileSetSectorAncho = 8
                TileSetSectorAlto = 8
            Else
                ' Es el de transiciones. Tengo que hacordear un poco
                If area.izquierda < 8 Then ' Transiciones Verticales (Oeste y Este)
                
                    ' Los obligo a seleccionar correctamente
                    area.izquierda = (area.izquierda \ 2) * 2
                    area.derecha = area.izquierda + 1
                    
                    xMaximoArea = 9
                    yMaximoArea = 8
                    TileSetSectorAncho = 8
                    TileSetSectorAlto = 8
                    offsetX = -7 + area.izquierda
                Else ' Transiciones Horizontales (Sur y Norte)
                
                    ' Los obligo a seleccionar correctamente
                    area.arriba = (area.arriba \ 2) * 2
                    area.abajo = area.arriba + 1
                    
                    xMaximoArea = 8
                    yMaximoArea = 9
                    TileSetSectorAncho = 8
                    TileSetSectorAlto = 8
                    offsetY = area.arriba - (7 + (area.arriba \ yMaximoArea) * yMaximoArea)
                End If
            End If
        ElseIf formato = eFormatoTileSet.textura_agua Or formato = costa_tipo_1_parte2 Or formato = rocas_acuaticas Then
            xMaximoArea = 16
            yMaximoArea = 16
            TileSetSectorAncho = 16
            TileSetSectorAlto = 16
        End If
    End If
    
    ' No se fuerza. Dejo el tamaño a tal cual lo selecciono
    If xMaximoArea = 0 Then
        xMaximoArea = area.derecha - area.izquierda + 1
        yMaximoArea = area.abajo - area.arriba + 1
        offsetX = offsetX + area.izquierda
        offsetY = offsetY + area.arriba
        TileSetSectorAncho = 0
        TileSetSectorAlto = 0
    Else
        offsetX = offsetX + (area.izquierda \ xMaximoArea) * xMaximoArea
        offsetY = offsetY + (area.arriba \ yMaximoArea) * yMaximoArea
    End If

    ' Redimensionamos donde voy a guardar la data del tileset seleccionado
    ReDim Me_Tools_TileSet.tilesetNumeroSeleccionado(0 To xMaximoArea - 1, 0 To yMaximoArea - 1)
    ReDim Me_Tools_TileSet.TilesetSeleccionado(0 To xMaximoArea - 1, 0 To yMaximoArea - 1)
    
    ' Blanqueo
    For xAux = 0 To xMaximoArea - 1
        For yAux = 0 To yMaximoArea - 1
            Me_Tools_TileSet.tilesetNumeroSeleccionado(xAux, yAux) = 0
            Me_Tools_TileSet.TilesetSeleccionado(xAux, yAux) = 0
        Next yAux
    Next xAux
    
    ' Establezco lo seleccionado
    For xAux = area.izquierda To area.derecha
        For yAux = area.arriba To area.abajo
        
            tileX = xAux - offsetX
            tileY = yAux - offsetY
            
            'Esta tile del tileset donde lo tengo que posicionar dentro del area que voy a insertar?
            If tileX < xMaximoArea And tileY < yMaximoArea Then
                If TileSetVirtual = 0 Then
                    Me_Tools_TileSet.tilesetNumeroSeleccionado(tileX, tileY) = xAux + yAux * 16
                    Me_Tools_TileSet.TilesetSeleccionado(tileX, tileY) = tileset_actual
                Else
                    Me_Tools_TileSet.tilesetNumeroSeleccionado(tileX, tileY) = Tilesets(tileset).matriz_transformacion(TileSetVirtual, xAux, yAux).numero
                    Me_Tools_TileSet.TilesetSeleccionado(tileX, tileY) = Tilesets(tileset).matriz_transformacion(TileSetVirtual, xAux, yAux).textura
                End If
            End If
        Next yAux
    Next xAux
    
   
    Me_Tools_TileSet.herramientaInternaTileSet = eHerramientasTileSet.insertar
    
    
    Call ME_Tools.establecerAmpliacionDeArea(CInt(xMaximoArea), CInt(yMaximoArea))

End Sub

Public Function aplicarTexturaTodoMapa(ByVal tileset As Integer, area As tAreaSeleccionada)

    Dim loopX As Integer
    Dim loopY As Integer
    
    Dim xInicial As Integer
    Dim yInicial As Integer
    Dim xFinal As Integer
    Dim yFinal As Integer
    
    Dim tileRelativoX As Integer
    Dim tileRelativoY As Integer

            
    xInicial = X_MINIMO_USABLE
    yInicial = Y_MINIMO_USABLE
        
    xFinal = SV_Constantes.X_MAXIMO_USABLE
    yFinal = SV_Constantes.Y_MAXIMO_USABLE
        
    ' Guardamos el original
    Dim TileSetOriginal() As Integer
    Dim TileSetNumeroOriginal() As Integer
    
    ReDim TileSetOriginal(LBound(TilesetSeleccionado, 1) To UBound(TilesetSeleccionado, 1), LBound(TilesetSeleccionado, 2) To UBound(TilesetSeleccionado, 2)) As Integer
    ReDim TileSetNumeroOriginal(LBound(tilesetNumeroSeleccionado, 1) To UBound(tilesetNumeroSeleccionado, 1), LBound(tilesetNumeroSeleccionado, 2) To UBound(tilesetNumeroSeleccionado, 2)) As Integer
    
    TileSetOriginal = TilesetSeleccionado
    TileSetNumeroOriginal = tilesetNumeroSeleccionado
    
    ReDim TilesetSeleccionado(X_MINIMO_USABLE To X_MAXIMO_USABLE, Y_MINIMO_USABLE To Y_MAXIMO_USABLE) As Integer
    ReDim tilesetNumeroSeleccionado(X_MINIMO_USABLE To X_MAXIMO_USABLE, Y_MINIMO_USABLE To Y_MAXIMO_USABLE) As Integer

    ' Recorremos para cambiar los -1 (marca especial) por 0 (sin tileset)
    For loopX = xInicial To xFinal
        For loopY = yInicial To yFinal
             
            tileRelativoX = (loopX - GRILLA_OFFSET_X + 1) Mod (UBound(TileSetOriginal, 1) + 1)
            tileRelativoY = (loopY - GRILLA_OFFSET_Y + 1) Mod (UBound(TileSetOriginal, 2) + 1)
                          
            TilesetSeleccionado(loopX, loopY) = TileSetOriginal(tileRelativoX, tileRelativoY)
            tilesetNumeroSeleccionado(loopX, loopY) = TileSetNumeroOriginal(tileRelativoX, tileRelativoY)
    
        Next loopY
    Next loopX
    
    'Seleccionamos el area donde vamos a trabajar
    Call modSeleccionArea.puntoArea(ME_Tools.areaSeleccionada, X_MINIMO_USABLE, Y_MINIMO_USABLE)
    Call modSeleccionArea.actualizarArea(ME_Tools.areaSeleccionada, X_MAXIMO_USABLE, Y_MAXIMO_USABLE)

    Me_Tools_TileSet.herramientaInternaTileSet = eHerramientasTileSet.insertar
    Call selectToolMultiple(Tools.tool_tileset, "Insertar piso en todo el mapa")
    
    aplicarTexturaTodoMapa = True
    
End Function

Private Sub insertarPisoExpansivoPos(posicion As cPosition, posiciones As Collection, posicionOriginal As cPosition, TileSetOriginal() As Integer, TileSetNumeroOriginal() As Integer, considerarBloqueos As Boolean)
    Dim posicionNuevaParaAnalizar As cPosition
    Dim loopX As Integer
    Dim loopY As Integer
    
    Dim tileRelativoX As Integer
    Dim tileRelativoY As Integer
    
    Dim tamX As Integer
    Dim tamY As Integer
    
    ' ¿Dentro de los limites?
    If posicion.x < X_MINIMO_USABLE Or posicion.x > X_MAXIMO_USABLE Then Exit Sub
    If posicion.y < Y_MINIMO_USABLE Or posicion.y > Y_MAXIMO_USABLE Then Exit Sub
    
    ' Ya hay un piso aca?. o lo voy a poner?
    If mapdata(posicion.x, posicion.y).tile_texture = 0 And TilesetSeleccionado(posicion.x, posicion.y) = 0 And (considerarBloqueos = False Or ((mapdata(posicion.x, posicion.y).Trigger And eTriggers.TodosBordesBloqueados) = 0) Or mapdata(posicion.x, posicion.y).Graphic(3).GrhIndex > 0) Then
    
        tamX = UBound(TileSetOriginal, 1) - LBound(TileSetOriginal, 1) + 1
        tamY = UBound(TileSetOriginal, 2) - LBound(TileSetOriginal, 2) + 1
    
        If posicion.x - posicionOriginal.x >= 0 Then
            tileRelativoX = (posicion.x - posicionOriginal.x) Mod tamX
        Else
            tileRelativoX = tamX - (posicionOriginal.x - posicion.x) Mod tamX
            If tileRelativoX = tamX Then tileRelativoX = 0
        End If
        
        If posicion.y - posicionOriginal.y >= 0 Then
           tileRelativoY = (posicion.y - posicionOriginal.y) Mod tamY
        Else
            tileRelativoY = tamY - (posicionOriginal.y - posicion.y) Mod tamY
            If tileRelativoY = tamY Then tileRelativoY = 0
        End If
        
        ' Marcamos para bloquear
        TilesetSeleccionado(posicion.x, posicion.y) = TileSetOriginal(tileRelativoX, tileRelativoY)
        tilesetNumeroSeleccionado(posicion.x, posicion.y) = TileSetNumeroOriginal(tileRelativoX, tileRelativoY)
        
        ' Sino tiene tileset, lo marcamos como -1 para no volver a procesarlo
        If TilesetSeleccionado(posicion.x, posicion.y) = 0 Then
            TilesetSeleccionado(posicion.x, posicion.y) = -1
        End If
                    
        For loopX = posicion.x - 1 To posicion.x + 1
        
            For loopY = posicion.y - 1 To posicion.y + 1
                
                ' Evitamos procesar la actual
                If Not (posicion.x = loopX And posicion.y = loopY) Then
                        
                        ' Creamos la nueva posicion a analizar
                        Set posicionNuevaParaAnalizar = New cPosition
                        posicionNuevaParaAnalizar.x = loopX
                        posicionNuevaParaAnalizar.y = loopY
                        Call posiciones.Add(posicionNuevaParaAnalizar)
                        
                End If
            
            Next
        
        Next

    End If

End Sub
' A partir de una posicion hace una insercion expansiva hasta ser encerrado por otros pisos
Public Function generarPisoExpansivo(ByVal x As Byte, ByVal y As Byte, considerarBloqueos As Boolean) As Boolean
    Dim loopX As Integer
    Dim loopY As Integer
    
    Dim posicionesPendientes As Collection 'Posiciones que tengo que analizar
    Dim posicion As cPosition ' Posicion actual a analizar
    Dim posicionOriginal As cPosition
    Set posicionesPendientes = New Collection
    
    ' Chequeamos que parta de una posicion correcta para insertar el piso
    If TileSetSectorAncho > 0 And TileSetSectorAlto > 0 Then
        If Not (((x + GRILLA_OFFSET_X - 1) Mod Me_Tools_TileSet.TileSetSectorAncho = 0) And ((GRILLA_OFFSET_Y + y - 1) Mod Me_Tools_TileSet.TileSetSectorAlto) = 0) Then
            generarPisoExpansivo = False
            Exit Function
        End If
    End If
    
    
    ' Posicion incial
    Set posicion = New cPosition
    posicion.x = x
    posicion.y = y
    
    Set posicionOriginal = posicion
    
    ' Guardamos el original
    Dim TileSetOriginal() As Integer
    Dim TileSetNumeroOriginal() As Integer
    
    ReDim TileSetOriginal(LBound(TilesetSeleccionado, 1) To UBound(TilesetSeleccionado, 1), LBound(TilesetSeleccionado, 2) To UBound(TilesetSeleccionado, 2)) As Integer
    ReDim TileSetNumeroOriginal(LBound(tilesetNumeroSeleccionado, 1) To UBound(tilesetNumeroSeleccionado, 1), LBound(tilesetNumeroSeleccionado, 2) To UBound(tilesetNumeroSeleccionado, 2)) As Integer
    
    TileSetOriginal = TilesetSeleccionado
    TileSetNumeroOriginal = tilesetNumeroSeleccionado
    
    ReDim TilesetSeleccionado(X_MINIMO_USABLE To X_MAXIMO_USABLE, Y_MINIMO_USABLE To Y_MAXIMO_USABLE) As Integer
    ReDim tilesetNumeroSeleccionado(X_MINIMO_USABLE To X_MAXIMO_USABLE, Y_MINIMO_USABLE To Y_MAXIMO_USABLE) As Integer

    ' Expando
    Call insertarPisoExpansivoPos(posicion, posicionesPendientes, posicionOriginal, TileSetOriginal, TileSetNumeroOriginal, considerarBloqueos)
        
    ' Mientras haya posiciones que debo analizar
    Do While posicionesPendientes.count > 0
        
        ' Obtengo la posicion
        Set posicion = posicionesPendientes.item(1)
        posicionesPendientes.Remove (1)
        
        ' Bloqueamos y expandimos
        Call insertarPisoExpansivoPos(posicion, posicionesPendientes, posicionOriginal, TileSetOriginal, TileSetNumeroOriginal, considerarBloqueos)
    Loop
    
    ' Recorremos para cambiar los -1 (marca especial) por 0 (sin tileset)
    For loopX = LBound(TilesetSeleccionado, 1) To UBound(TilesetSeleccionado, 1)
        For loopY = LBound(TilesetSeleccionado, 2) To UBound(TilesetSeleccionado, 2)
            If TilesetSeleccionado(loopX, loopY) = -1 Then TilesetSeleccionado(loopX, loopY) = 0
        Next loopY
    Next loopX
    
    'Seleccionamos el area donde vamos a trabajar
    Call modSeleccionArea.puntoArea(ME_Tools.areaSeleccionada, X_MINIMO_USABLE, Y_MINIMO_USABLE)
    Call modSeleccionArea.actualizarArea(ME_Tools.areaSeleccionada, X_MAXIMO_USABLE, Y_MAXIMO_USABLE)

    Me_Tools_TileSet.herramientaInternaTileSet = eHerramientasTileSet.insertar
    Call selectToolMultiple(Tools.tool_tileset, "Aplicar piso a área")
    
    generarPisoExpansivo = True
End Function
