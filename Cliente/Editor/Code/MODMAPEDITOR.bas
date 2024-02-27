Attribute VB_Name = "ME_Tools"
Option Explicit

Public CONTROL_APRETADO As Boolean

Public TILE_SELECTED As Position

Public MOSTRAR_TILESET As Boolean
Public TiempoBotonTileSetApretado As Long

Public mostrandoVWindows As Boolean

Public TOOL_SELECC As Long
Public TOOL_MOUSEOVER As Integer

Public Const toolcount As Integer = 13
Const toolnames As String = "Ninguna-Bloqueos-Tilesets-Montañas-Modificar Altura de tile-Triggers-Luces-Posicion del sol-Filtros-Partículas-Translados-Npcs-Objetos-Graficos-"

Public Enum Tools
    tool_none = 0
    tool_bloqueo = 1
    tool_tileset = 2
    tool_montaña = 4
    tool_altura_vertex = 8
    tool_triggers = 16
    tool_luces = 32
    tool_sol = 64
    tool_filtros = 128
    tool_particles = 256
    tool_acciones = 512
    tool_npc = 1024
    tool_obj = 2048
    tool_grh = 4096
    tool_seleccion = 8192
    tool_copiar = 16384
    tools_seleccionMinima = 32768
    tool_entidades = 65536
    tool_todas = 131071
End Enum

Public tool_bloqueo_area As Boolean

Public DRAWTRIGGERS As Byte
Public DRAWBLOQUEOS As Byte
Public dibujarZonaNacimientoCriatura As Byte ' Muestro en que zona nace cada criatura
Public dibujarZonaNacimientoCriaturas As Byte ' Dibujo la Zonas donde nacen las criaturas
Public dibujarAccionTile As Byte 'Dibujo o no el indicador de que hay una accion?
Public dibujarCantidadObjetos As Byte 'Dibujo o no la cantidad de objetos que hay en un tile?
Public mostrarTileDondeHayLuz As Byte
Public mostrarTileNumber As Byte
Public mostrarTileEfectoSonidoPasos As Byte
Public dibujarTechosTransparentes As Byte
Public mostrarTileDondeHayGraficos As Byte

Public clickpos As Position
Public clickposp As Position

Public radio_montana As Integer
Public editando_montaña As Boolean

Public montaña_clean As Boolean
Public montaña_meseta As Boolean

Public Enum mtools
    mt_promedio = 0
    mp_suma
    mt_slerp
    mt_clean
    mt_meseta
    mt_blur
    mt_pie
End Enum

Public mt_select As Long

Public Enum mfiltros
    filtro_none = 0
    filtro_linear
    filtro_bilinear
    filtro_trilinear
    filtro_ansitropic
End Enum

Public Const mfiltroscount As Integer = 4
Public mf_select As Long
Const mfiltrosnames As String = "NONE-LINEAR-BILINEAR-TRILINEAR-ANSITROPIC"

Public hMapDataORIGINAL(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As AUDT
Public alturasORIGINAL(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As Integer
Public NormalDataORIGINAL(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As D3DVECTOR

Public AlturaPieORIGINAL(X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As Integer

'simular modAreas.bas
Public Const ARangoX = 12
Public Const ARangoY = 10
'/simular

Public inicial_click_tile As D3DVECTOR2
Public final_click_tile As D3DVECTOR2
Public inicial_click As D3DVECTOR2
Public final_click As D3DVECTOR2

'TILESETS
Public seleccionando_area_tileset As Boolean
Public Area_Tileset As tAreaSeleccionada
Public tileset_actual As Integer
Public tileset_actual_virtual As Byte ' Nuevo formato de pisos
'/TILESETS

Public Type Mapa
    nombre As String * 32
    editado As Boolean
    numero As Integer
    Path As String
    Pak As Boolean
    Version As Long
    Autor As Long 'ID DEL CDM
End Type

'mapa
Public THIS_MAPA As Mapa

'Nombre de la herramienta seleccionada actualmente
Public tool_act_name As String

'/mapa
Type SupData
    Name As String
    Grh As Integer
    width As Byte
    height As Byte
    block As Boolean
    Capa As Byte
End Type
Public MaxSup As Integer
Public SupData() As SupData
''
'************************************
'/ HACER - RE HACER - DES HACER    '*
Private botonUtilizado As CommandButton

Public HerramientaIndiceInterno As Integer
Public HerramientaIndiceInternoMaximo As Integer

'Seleccion
Public areaSeleccionada As tAreaSeleccionada
Public areaSeleccionadaAmpliacionAncho As Integer
Public areaseleccioandaAmpliacionAlto As Integer

Public Sub iniciar()
    radio_montana = 3
    deseleccionarTool
    Call ME_Tools_Triggers.iniciarToolTrigger
    Call Me_Tools_Objetos.iniciarToolObjetos
    Call Me_Tools_Npc.iniciarToolNPC
    Call ME_Tools_Acciones.iniciarToolAcciones
    Call Me_Tools_TileSet.iniciarToolTileSets
    Call ME_Tools_Graficos.iniciarToolGraficos
    Call ME_Tools_Particulas.iniciarToolsParticulas
    Call Me_Tools_Entidades.iniciarToolEntidades
End Sub

Public Sub rotarHerramientaInterna(paraArriba As Boolean)
        
    Select Case ME_Tools.TOOL_SELECC
        Case Tools.tool_bloqueo
            Call ME_Tools_Triggers.rotarHerramientaInternaBloqueo(paraArriba)
        Case Tools.tool_triggers
            Call ME_Tools_Triggers.rotarHerramientaInternaTrigger(paraArriba)
        Case Tools.tool_acciones
            Call ME_Tools_Acciones.rotarHerramientaInterna(paraArriba)
        Case Tools.tool_npc
            Call Me_Tools_Npc.rotarHerramientaInternaNPC(paraArriba)
        Case Tools.tool_obj
            Call Me_Tools_Objetos.rotarHerramientaInternaObjeto(paraArriba)
        Case Tools.tool_grh
            Call ME_Tools_Graficos.rotarHerramientaInterna(paraArriba)
        Case Tools.tool_luces
            Call Me_Tools_Luces.rotarHerramientaInterna(paraArriba)
        Case Tools.tool_montaña
            If paraArriba Then
                If radio_montana < 10 Then radio_montana = radio_montana + 1
            Else
                If radio_montana > 1 Then radio_montana = radio_montana - 1
            End If
            
            frmMain.radio_montaña.value = radio_montana
            frmMain.radio_montaña_lbl.caption = "Radio: " & radio_montana
            Restore_HM
        Case Tools.tool_entidades
            Call Me_Tools_Entidades.rotarHerramientaInternaEntidad(paraArriba)
    End Select

End Sub

Public Sub selectToolMultiple(tool As Tools, nombreMultipleTool As String, Optional botonPresionado As CommandButton = Nothing)
    tool_act_name = nombreMultipleTool
    TOOL_SELECC = tool
    
    If Not botonPresionado Is Nothing Then
        Set botonUtilizado = botonPresionado
        botonUtilizado.font.bold = True
    End If
End Sub

Public Function isToolSeleccionada(tool As Tools) As Boolean
    isToolSeleccionada = (TOOL_SELECC And tool)
End Function


Public Sub seleccionarTool(botonPresionado As CommandButton, tool As Tools)

    Call deseleccionarTool
    
    If Not botonPresionado Is Nothing Then
        Set botonUtilizado = botonPresionado
        botonUtilizado.font.bold = True
    End If
    
    TOOL_SELECC = tool
End Sub

Public Sub deseleccionarTool()

    If Me_Tools_Seleccion.copiando Then
        Me_Tools_Seleccion.restablecerBackupHerramientas
    End If
    
    TOOL_SELECC = Tools.tool_none
    tool_act_name = "Ninguna"
    
    If Not botonUtilizado Is Nothing Then
        botonUtilizado.font.bold = False
    End If

End Sub

Public Sub ejecutarComando(comando As iComando)
    Call comando.hacer
    Call ME_modComandos.agregarComandoADesHacer(comando)
End Sub
'**************************************

Public Sub DRAW_TOOLS(ByVal toolnamesa As String, ByVal toolcounta As Long, ByVal TOOL_SELECCa As Long, Optional ByVal left As Single)
    Dim i As Integer
    Dim j$
    Dim ja() As String
    ja = Split(toolnamesa, "-")
    Dim ji$
    Dim Jo$
    For i = 0 To toolcounta
        If TOOL_SELECCa = i Then
            Jo = Chr$(255) & i & Chr$(255)
        Else
            Jo = i
        End If
        If TOOL_MOUSEOVER = i Then
            j = j & Jo & " - " & Chr$(255) & ja(i) & Chr$(255) & vbCrLf
        Else
            j = j & Jo & " - " & ja(i) & vbCrLf
        End If
        If Len(ja(i)) > Len(ji) Then ji = ja(i)
        
    Next i

    Engine.Draw_FilledBox left, 15, Engine.Engine_GetTextWidth(ji) + 45, (UBound(ja) + 1) * 16, &HFF000000, &H7F115511, 5
    Engine.Text_Render_ext j, 22, 10 + left, 0, 0, &H7FCCCCCC
End Sub


Public Sub DRAW_TOOL()
    Dim x%, y%, Color&, MX%, MY%, CurrentGrhIndex(1 To CANTIDAD_CAPAS) As Integer
    Dim tempLong As Long
    Dim tempInt As Long
    Dim tempInt2 As Integer
    Dim grosor As Byte
  
    Dim puedeModificarComportamiento As Boolean 'Permisos
    Dim puedeModificarVisual As Boolean
    
    Dim xArea As Integer 'Tile que estoy procesando
    Dim yArea As Integer

    Dim xRelativa As Integer 'Posicion relativa que estoy procesando dentro del area
    Dim yRelativa As Integer 'De 1 al ancho(x), largo (y)
    
    Dim xAreaInsertar As Integer
    Dim yAreaInsertar As Integer
    
    Dim loopX As Integer ' Auxiliares para Fors de cada herramienta
    Dim loopY As Integer
    
    Dim luz As tLuzPropiedades
    
    Engine.Text_Render_ext "Herramienta seleccionada: " & Chr$(255) & tool_act_name, 10, 10, 0, 0, &H7FCCCCCC
    
    If (TOOL_SELECC And Tools.tools_seleccionMinima) Then
        If areaSeleccionada.derecha = areaSeleccionada.izquierda And areaSeleccionada.arriba = areaSeleccionada.abajo Then
            areaSeleccionada.derecha = areaSeleccionada.izquierda + ME_Tools.areaSeleccionadaAmpliacionAncho - 1
            areaSeleccionada.abajo = areaSeleccionada.arriba + ME_Tools.areaseleccioandaAmpliacionAlto - 1
        End If
    End If

    'Proceso cada tile que tengo seleccionado
    For xArea = areaSeleccionada.izquierda To areaSeleccionada.derecha
        For yArea = areaSeleccionada.arriba To areaSeleccionada.abajo
            
            If ME_Mundo.existePosicion(xArea, yArea) Then
                'Permisos que se tienen sobre esta posicion
                puedeModificarComportamiento = ME_Mundo.puedeModificarComporamientoTile(xArea, yArea)
                puedeModificarVisual = ME_Mundo.puedeModificarAspectoTile(xArea, yArea)
                
                'La posicion relativa del area. desde 1 hasta el ancho/largo del area
                xRelativa = xArea - areaSeleccionada.izquierda
                yRelativa = yArea - areaSeleccionada.arriba
                
                'Calculos el pixel?
                x = (xArea + minXOffset) * 32 + offset_map.x
                y = (yArea + minYOffset) * 32 + offset_map.y
                
                'Marcamos este tile como seleccionado.
                Grh_Render_Blocked &H33FFFFFF, x, y, xArea, yArea
                                      
                'HERRAMIENTAS QUE MODIFICAN LA VISUAL
                If puedeModificarVisual Then
                
                    'Tileset
                    If TOOL_SELECC And Tools.tool_tileset Then

                        If Not (MOSTRAR_TILESET) Then
 
                                'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 1) + xRelativa Mod (UBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 1) - LBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 1) + 1)
                                yAreaInsertar = LBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 2) + yRelativa Mod (UBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 2) - LBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 2) + 1)
                                                                
                                tempInt = Me_Tools_TileSet.TilesetSeleccionado(xAreaInsertar, yAreaInsertar)
                                
                                If tempInt > 0 Then
                                    tempInt = Tilesets(tempInt).filenum
                                    tempInt2 = Me_Tools_TileSet.tilesetNumeroSeleccionado(xAreaInsertar, yAreaInsertar)
                                    
                                    If tempInt > 0 And tempInt2 >= 0 Then
                
                                        Grh_Render_Relieve_Tileset_HC tempInt, tempInt2, x, y, xArea, yArea, 50
                                            
                                        If mostrarTileNumber = 1 Then
                                            Engine.Text_Render_ext CStr(tempInt2), y, x, 100, 10, &HFF0FF00F
                                        End If
                                            
                                        '¿Tiene restriccion en donde se puede insertar?
                                        If Me_Tools_TileSet.TileSetSectorAncho > 0 And Me_Tools_TileSet.TileSetSectorAlto > 0 Then
  
                                            ' ¿Acá esta bien?
                                            If Not (xRelativa Mod Me_Tools_TileSet.TileSetSectorAncho = (xArea + GRILLA_OFFSET_X - 1) Mod Me_Tools_TileSet.TileSetSectorAncho And yRelativa Mod Me_Tools_TileSet.TileSetSectorAlto = (GRILLA_OFFSET_Y + yArea - 1) Mod Me_Tools_TileSet.TileSetSectorAlto) Then
                                               ' Sino lo esta marcamos el error
                                               Grh_Render_Blocked &H33FFFFCC, x, y, xArea, yArea
                                            End If
                                            
                                        End If
                                        
                                    End If
                                End If
                                
                        End If
                    End If

                    'Graficos
                    If TOOL_SELECC And Tools.tool_grh Then
                    
                        If ME_Tools_Graficos.herramientaInternaGraficos = eHerramientaGraficos.insertar Then
                            
                             'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) + xRelativa Mod (UBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) - LBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) + 1)
                            yAreaInsertar = LBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) + yRelativa Mod (UBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) - LBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) + 1)
                                
                            'Grh_Render_Blocked mzGreen, (MouseTileX + minXOffset) * 32 + offset_map.x, (MouseTileY + minYOffset) * 32 + offset_map.y, MouseTileX, MouseTileY
                            
                            For MX = 1 To CANTIDAD_CAPAS
                                If GrhData(ME_Tools_Graficos.grhInfoSeleccionada(xAreaInsertar, yAreaInsertar).grhInfoPosicion(MX).GrhIndex).NumFrames > 1 Then
                                    CurrentGrhIndex(MX) = GrhData(grhInfoSeleccionada(xAreaInsertar, yAreaInsertar).grhInfoPosicion(MX).GrhIndex).Frames(1)
                                Else
                                    CurrentGrhIndex(MX) = ME_Tools_Graficos.grhInfoSeleccionada(xAreaInsertar, yAreaInsertar).grhInfoPosicion(MX).GrhIndex
                                End If
                            Next MX
            
            
                            If CurrentGrhIndex(1) Then
                                Grh_Render_relieve CurrentGrhIndex(1), _
                                    x, y, _
                                    xArea, yArea
                            End If
            
                            For MX = 2 To CANTIDAD_CAPAS
                                Grh_Render_new CurrentGrhIndex(MX), x + GrhData(CurrentGrhIndex(MX)).offsetX, y + GrhData(CurrentGrhIndex(MX)).offsetY, xArea, yArea
                            Next MX
                        ElseIf ME_Tools_Graficos.herramientaInternaGraficos = eHerramientaGraficos.borrar Then
                            Grh_Render_Blocked mzGreen And &H7FFFFFFF, x, y, xArea, yArea
                        End If
                    End If
                    
                    'TO-DO Faltan las particulas

                    'Luces
                    If TOOL_SELECC And Tools.tool_luces Then

                        If Me_Tools_Luces.herramientaInternaLuces = eHerramientasLuces.insertar Then
                            'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(Me_Tools_Luces.infoLuzSeleccionada, 1) + xRelativa Mod (UBound(Me_Tools_Luces.infoLuzSeleccionada, 1) - LBound(Me_Tools_Luces.infoLuzSeleccionada, 1) + 1)
                            yAreaInsertar = LBound(Me_Tools_Luces.infoLuzSeleccionada, 2) + yRelativa Mod (UBound(Me_Tools_Luces.infoLuzSeleccionada, 2) - LBound(Me_Tools_Luces.infoLuzSeleccionada, 2) + 1)

                            luz = Me_Tools_Luces.infoLuzSeleccionada(xAreaInsertar, yAreaInsertar)

                            If luz.LuzRadio > 0 Then
                                For loopX = xArea - luz.LuzRadio To xArea + luz.LuzRadio
                                   For loopY = yArea - luz.LuzRadio To yArea + luz.LuzRadio
                                        If InMapBounds(loopX, loopY) Then
                                            If ME_Mundo.puedeModificarAspectoTile(loopX, loopY) Then
                                            
                                                If (luz.LuzTipo And TipoLuces.Luz_Cuadrada) Then
                                                    Color = 1
                                                Else
                                                    Color = Sqr(((yArea - loopY) * (yArea - loopY) + (xArea - loopX) * (xArea - loopX)))
                                                End If
                                                
                                                If Color < luz.LuzRadio Then
                                                    Color = D3DColorARGB((1 - Color / luz.LuzRadio) * 255, luz.LuzColor.r, luz.LuzColor.g, luz.LuzColor.b)
        
                                                    Grh_Render_Blocked Color, (loopX + minXOffset) * 32 + offset_map.x, (loopY + minYOffset) * 32 + offset_map.y, loopX, loopY
                                                End If
                                                
                                            End If
                                       End If
                                    Next loopY
                               Next loopX
                           End If
                       End If
                    End If
                
                    'Montañas
                    'HERRAMIENTAS QUE FALTAN ESTANDARIZAR
                    If TOOL_SELECC And Tools.tool_montaña Then
                        If mt_select = mt_pie Then
                            D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
                            MX = clickpos.x
                            MY = clickpos.y + 1
                            For x = MX - radio_montana To MX + radio_montana
                                For y = MY - radio_montana To MY + radio_montana
                                    If InMapBounds(x, y) And InMapBounds(x + 1, y - 1) Then
                                        Color = Sqr(((MY - y) * (MY - y) + (MX - x) * (MX - x)))
                                        If Color < radio_montana Then
                                            Color = D3DColorMake(1, 1, 1, 1 - Color / radio_montana)
                    
                                            Grh_Render_PIE Color, (x + minXOffset) * 32 + offset_map.x, (y + minYOffset - 1) * 32 + offset_map.y, x, y
                                        End If
                                    End If
                                Next y
                            Next x
                            D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
                            If (InMapBounds(clickpos.x, MY - 1) And InMapBounds(clickpos.x + 1, MY - 2)) Then
                                Grh_Render_PIE_Tool (MX + minXOffset) * 32 + offset_map.x, (MY + minYOffset - 1) * 32 + offset_map.y, clickpos.x, MY - 1
                            End If
                        Else
                        
                            For x = clickpos.x - radio_montana To clickpos.x + radio_montana
                                For y = clickpos.y - radio_montana To clickpos.y + radio_montana
                                If InMapBounds(x, y) Then
                                    Color = Sqr(((clickpos.y - y) * (clickpos.y - y) + (clickpos.x - x) * (clickpos.x - x)))
                                    If Color < radio_montana Then
                                        Color = D3DColorMake(0, 1, 0, 1 - Color / radio_montana)
                
                                        Grh_Render_Blocked Color, (x + minXOffset) * 32 + offset_map.x, (y + minYOffset) * 32 + offset_map.y, x, y
                                    End If
                                End If
                                Next y
                            Next x
                        End If
                        If TOOL_MOUSEOVER = 1 Then
                            Engine.Text_Render_ext "Radio: " & Chr$(255) & radio_montana & "   Se modifica con las teclas + y -", 32, 10, 0, 0, &HFFCCCCCC
                        Else
                            Engine.Text_Render_ext "Radio: " & Chr$(255) & radio_montana, 32, 10, 0, 0, &H7FCCCCCC
                        End If
                    End If
                End If

                '******************************************************************
                'HERRAMIENTAS QUE MODIFICAN EL COMPORTAMIENTO
                If puedeModificarComportamiento Then
                
                    'Triggers
                    If TOOL_SELECC And Tools.tool_triggers Then
                                       
                        If ME_Tools_Triggers.herramientaInternaTrigger = herramientasTriggers.insertar Then
                            
                            'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(ME_Tools_Triggers.triggerSeleccionado, 1) + xRelativa Mod (UBound(ME_Tools_Triggers.triggerSeleccionado, 1) - LBound(ME_Tools_Triggers.triggerSeleccionado, 1) + 1)
                            yAreaInsertar = LBound(ME_Tools_Triggers.triggerSeleccionado, 2) + yRelativa Mod (UBound(ME_Tools_Triggers.triggerSeleccionado, 2) - LBound(ME_Tools_Triggers.triggerSeleccionado, 2) + 1)
                                    
                          '  Grh_Render_Blocked mzYellow And &H7FFFFFFF, x, y, xArea, yArea
                            Engine.Text_Render_ext ME_Tools_Triggers.obtenerDescripcionAbreviatura(ME_Tools_Triggers.triggerSeleccionado(xAreaInsertar, yAreaInsertar)), y + 8, x, 0, 0, &HFFFFFFFF
                        
                            'Dibujo los bordes bloqueados que no tienen descripcion
                            If (ME_Tools_Triggers.triggerSeleccionado(xAreaInsertar, yAreaInsertar) And eTriggers.BloqueoEste) Then
                                Draw_FilledBox x + 30, y, 2, 32, &HFFFC0000, 0, 0
                            End If
                            
                            If (ME_Tools_Triggers.triggerSeleccionado(xAreaInsertar, yAreaInsertar) And eTriggers.BloqueoOeste) Then
                                Draw_FilledBox x, y, 2, 32, &HFFFC0000, 0, 0
                            End If
                            
                            If (ME_Tools_Triggers.triggerSeleccionado(xAreaInsertar, yAreaInsertar) And eTriggers.BloqueoNorte) Then
                                Draw_FilledBox x, y, 32, 2, &HFFFC0000, 0, 0
                            End If
                            
                            If (ME_Tools_Triggers.triggerSeleccionado(xAreaInsertar, yAreaInsertar) And eTriggers.BloqueoSur) Then
                                Draw_FilledBox x, y + 30, 32, 2, &HFFFC0000, 0, 0
                            End If
                        
                        ElseIf ME_Tools_Triggers.herramientaInternaTrigger = herramientasTriggers.borrar Then
                            Grh_Render_Blocked mzYellow And &H4FFFFFFF, x, y, xArea, yArea
                            Engine.Text_Render_ext "Borrar", y + 8, x, 0, 0, &HFFFFFFFF
                        End If
                    End If

                    
                    'Bloqueos
                    If TOOL_SELECC And Tools.tool_bloqueo Then
                        If ME_Tools_Triggers.herramientaInternaBloqueo <> 0 Then
                        
                            'Dependiendo la herramienta es el grosor y el color del borde
                            If ME_Tools_Triggers.herramientaInternaBloqueo = InsertarDoble Then
                                grosor = 4
                                Color = &HFFFC0000
                            ElseIf ME_Tools_Triggers.herramientaInternaBloqueo = BorrarSimple Then
                                grosor = 2
                                Color = mzWhite
                            Else
                                grosor = 2
                                Color = &HFFFC0000
                            End If
                            
                            'Dibujo el borde sobre el cual esta el mouse
                            Select Case ME_Tools.obtenerBordeSeleccionado()
                                Case E_Heading.EAST
                                    Draw_FilledBox x + 30, y, grosor, 32, Color, 0, 0
                                Case E_Heading.WEST
                                    Draw_FilledBox x, y, grosor, 32, Color, 0, 0
                                Case E_Heading.NORTH
                                    Draw_FilledBox x, y, 32, grosor, Color, 0, 0
                                Case E_Heading.SOUTH
                                    Draw_FilledBox x, y + 30, 32, grosor, Color, 0, 0
                            End Select
                        End If
                    End If
                    
                    'Objetos
                    If TOOL_SELECC And Tools.tool_obj Then
                        If Me_Tools_Objetos.herramientaInternaOBJ = eHerramientasOBJ.insertar Then  'Esto me dice si hay algun borde bloqueado
                            
                            'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(Me_Tools_Objetos.objIndexSeleccionado, 1) + xRelativa Mod (UBound(Me_Tools_Objetos.objIndexSeleccionado, 1) - LBound(Me_Tools_Objetos.objIndexSeleccionado, 1) + 1)
                            yAreaInsertar = LBound(Me_Tools_Objetos.objIndexSeleccionado, 2) + yRelativa Mod (UBound(Me_Tools_Objetos.objIndexSeleccionado, 2) - LBound(Me_Tools_Objetos.objIndexSeleccionado, 2) + 1)
                            
                            'Debug.Print xAreaInsertar
                            'Debug.Print yAreaInsertar
                            'Obtengo el index del objeto
                            tempInt = Me_Tools_Objetos.objIndexSeleccionado(xAreaInsertar, yAreaInsertar)
                            
                            If tempInt > 0 Then
                                tempLong = ObjData(tempInt).GrhIndex
   
                                Grh_Render_new tempLong, x + GrhData(tempLong).offsetX, y + GrhData(tempLong).offsetY, xArea, yArea
                                Engine.Text_Render_ext CStr(Me_Tools_Objetos.objCantidadSeleccionado(xAreaInsertar, yAreaInsertar)), y + 7, x + 5, 100, 10, &HFFFFFFFF
                            End If
                           
                        ElseIf Me_Tools_Objetos.herramientaInternaOBJ = eHerramientasOBJ.borrar Then
                            Grh_Render_Blocked mzYellow And &H7FFFFFFF, x, y, xArea, yArea
                        End If
                    End If
                    
                    'Npcs
                    If TOOL_SELECC And Tools.tool_npc Then
                        If Me_Tools_Npc.herramientaInternaNPC = eHerramientasNPC.insertar Then
                            
                            'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(Me_Tools_Npc.NPCSeleccionado, 1) + xRelativa Mod (UBound(Me_Tools_Npc.NPCSeleccionado, 1) - LBound(Me_Tools_Npc.NPCSeleccionado, 1) + 1)
                            yAreaInsertar = LBound(Me_Tools_Npc.NPCSeleccionado, 2) + yRelativa Mod (UBound(Me_Tools_Npc.NPCSeleccionado, 2) - LBound(Me_Tools_Npc.NPCSeleccionado, 2) + 1)
                            
                            'Obtengo el index de la criatura
                            tempInt = Me_Tools_Npc.NPCSeleccionado(xAreaInsertar, yAreaInsertar).Index
                        
                            'Dibujar NPC
                            If tempInt > 0 Then
                                tempLong = BodyData(NpcData(tempInt).body).Walk(E_Heading.SOUTH).GrhIndex
                                
                                'Si es una animacion, elijo un grafico.
                                If GrhData(tempLong).NumFrames > 0 Then
                                    tempLong = GrhData(tempLong).Frames(1)
                                End If

                                Grh_Render_new tempLong, x + GrhData(tempLong).offsetX, y + GrhData(tempLong).offsetY, xArea, yArea
                            End If
                            
                        ElseIf Me_Tools_Npc.herramientaInternaNPC = eHerramientasNPC.borrar Then
                            Grh_Render_Blocked mzBlue And &H4FFFFFFF, x, y, xArea, yArea
                        End If
                    End If
        
                    'Acciones
                    If TOOL_SELECC And Tools.tool_acciones Then
                    
                        If ME_Tools_Acciones.herramientainterna = eHerramientasAccion.insertar Then
                        
                            'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(ME_Tools_Acciones.accionSeleccionada, 1) + xRelativa Mod (UBound(ME_Tools_Acciones.accionSeleccionada, 1) - LBound(ME_Tools_Acciones.accionSeleccionada, 1) + 1)
                            yAreaInsertar = LBound(ME_Tools_Acciones.accionSeleccionada, 2) + yRelativa Mod (UBound(ME_Tools_Acciones.accionSeleccionada, 2) - LBound(ME_Tools_Acciones.accionSeleccionada, 2) + 1)
                            
                            If Not ME_Tools_Acciones.accionSeleccionada(xAreaInsertar, yAreaInsertar) Is Nothing Then
                                Grh_Render_Blocked mzRed And &H7FFFFFFF, x, y, xArea, yArea
                                Engine.Text_Render_ext ME_Tools_Acciones.accionSeleccionada(xAreaInsertar, yAreaInsertar).GetNombre, y, x + 5, 0, 0, &HFFFFFFFF
                            End If
                        ElseIf ME_Tools_Acciones.herramientainterna = eHerramientasAccion.borrar Then
                            Grh_Render_Blocked mzRed And &H4FFFFFFF, x, y, MouseTileX, MouseTileY
                            Engine.Text_Render_ext "Borrar", y, x + 5, 0, 0, &HFFFFFFFF
                        End If
                    End If
                    
                    'Entidades
                    If TOOL_SELECC And Tools.tool_entidades Then
                        If Me_Tools_Entidades.herramientaInternaEntidades = eHerramientasEntidades.insertar Then

                            'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(Me_Tools_Entidades.entidadesSeleccionadas, 1) + xRelativa Mod (UBound(Me_Tools_Entidades.entidadesSeleccionadas, 1) - LBound(Me_Tools_Entidades.entidadesSeleccionadas, 1) + 1)
                            yAreaInsertar = LBound(Me_Tools_Entidades.entidadesSeleccionadas, 2) + yRelativa Mod (UBound(Me_Tools_Entidades.entidadesSeleccionadas, 2) - LBound(Me_Tools_Entidades.entidadesSeleccionadas, 2) + 1)
                        
                        
                            
                            If Not Me_Tools_Entidades.entidadesSeleccionadas(xAreaInsertar, yAreaInsertar).infoEntidades(1).IndexEntidad = 0 Then
                                'Obtengo el index de la entidad
                                tempInt = Me_Tools_Entidades.entidadesSeleccionadas(xAreaInsertar, yAreaInsertar).infoEntidades(1).IndexEntidad
                        
                                If EntidadesIndexadas(tempInt).Graficos(0) > 0 Then
                                    tempLong = EntidadesIndexadas(tempInt).Graficos(0)
 
                                    If GrhData(tempLong).NumFrames > 1 Then
                                        tempLong = GrhData(tempLong).Frames(1)
                                    End If
                                    
                                    Grh_Render_new tempLong, x + GrhData(tempLong).offsetX, y + GrhData(tempLong).offsetY, xArea, yArea
                                End If
                            End If
                        ElseIf Me_Tools_Entidades.herramientaInternaEntidades = eHerramientasEntidades.borrar Then
                            Grh_Render_Blocked mzRed And &H4FFFFFFF, x, y, MouseTileX, MouseTileY
                            Engine.Text_Render_ext "Borrar", y, x + 5, 0, 0, &HFFFFFFFF
                        End If
                    End If
                End If
            End If
        Next yArea
    Next xArea
    
    
     'LISTA DE TILESETS DISPONIBLES
    If TOOL_SELECC And Tools.tool_tileset Then
        Engine.Text_Render_ext "Tileset: [" & tileset_actual & "] " & Chr$(255) & Tilesets(tileset_actual).nombre, 26, 10, 0, 0, &H7FCCCCCC
        
        ' Esto es cuando alguien quiere ver un tileset en crudo (configuración de pisos)
        If MOSTRAR_TILESET And tileset_actual > 0 And tileset_actual <= Tilesets_count Then
            Engine.Draw_FilledBox 0, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, &H8F000000, 0
            Grh_Render_Simple_box Tilesets(tileset_actual).filenum, 5, 0, &HFFFFFFFF, 512
        End If
    End If

    If TOOL_SELECC And Tools.tool_filtros Then
        Engine.Text_Render_ext Chr$(255) & "Filtros" & Chr$(255) & ">", 58, 10, 0, 0, &HFFCCFFCC
    End If

End Sub

Public Sub click_tool(Optional ByVal Button As MouseButtonConstants)
    
    'Variables de comandos
    Dim comando As cComandoInsertarTrigger
    Dim comandoAccion As cComandoInsertarAccion
    Dim comandoNpc As cComandoInsertarNPC
    Dim comandoObjeto As cComandoInsertarObjeto
    Dim comandoTileSet As cComandoInsertarTileSet
    Dim comandoGrafico As cComandoInsertarGrafico
    Dim comandoLuz As cComandoInsertarLuz
    Dim comandoEntidad As cComandoInsertarEntidad
    Dim comandoParticula As cComandoInsertarParticula
    
    Dim conjuntoDeComandos As Collection
    
    'Variables auxiliares
    Dim puedeModificarComportamiento As Boolean 'Permisos
    Dim puedeModificarVisual As Boolean
    
    Dim xArea As Integer 'Tile que estoy procesando
    Dim yArea As Integer
    
    Dim xRelativa As Integer 'Posicion relativa que estoy procesando dentro del area
    Dim yRelativa As Integer 'De 1 al ancho(x), largo (y)
        
    Dim xAreaInsertar As Integer
    Dim yAreaInsertar As Integer

    Dim tempInt As Integer
    Dim tempInt2 As Integer
    Dim tempbyte1 As Byte
    Dim tempbyte2 As Byte
    Dim i As Long

    Dim auxGrhInfo(1 To CANTIDAD_CAPAS) As tCapasPosicion
    Dim luz As tLuzPropiedades
    
    Set conjuntoDeComandos = New Collection

    'Ninguna herramienta seleccionada
    If (TOOL_SELECC Or Tools.tool_none) = Tools.tool_none Then Exit Sub
    
    If (TOOL_SELECC And Tools.tools_seleccionMinima) Then
        If areaSeleccionada.derecha = areaSeleccionada.izquierda And areaSeleccionada.arriba = areaSeleccionada.abajo Then
            areaSeleccionada.derecha = areaSeleccionada.izquierda + ME_Tools.areaSeleccionadaAmpliacionAncho - 1
            areaSeleccionada.abajo = areaSeleccionada.arriba + ME_Tools.areaseleccioandaAmpliacionAlto - 1
        End If
    End If

    'Proceso cada tile que tengo seleccionado
    For xArea = areaSeleccionada.izquierda To areaSeleccionada.derecha
        For yArea = areaSeleccionada.arriba To areaSeleccionada.abajo
            
            If ME_Mundo.existePosicion(xArea, yArea) Then
            
                'Permisos que se tienen sobre esta posicion
                puedeModificarComportamiento = ME_Mundo.puedeModificarComporamientoTile(xArea, yArea)
                puedeModificarVisual = ME_Mundo.puedeModificarAspectoTile(xArea, yArea)
                
               'La posicion relativa del area. desde 1 hasta el ancho/largo del area
                xRelativa = xArea - areaSeleccionada.izquierda
                yRelativa = yArea - areaSeleccionada.arriba
                
                If puedeModificarComportamiento Then
    
                    ' ACCIONES
                    If TOOL_SELECC And Tools.tool_acciones Then
        
                        Set comandoAccion = New cComandoInsertarAccion
            
                        If ME_Tools_Acciones.herramientainterna = eHerramientasAccion.insertar Then
                             'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(ME_Tools_Acciones.accionSeleccionada, 1) + xRelativa Mod (UBound(ME_Tools_Acciones.accionSeleccionada, 1) - LBound(ME_Tools_Acciones.accionSeleccionada, 1) + 1)
                            yAreaInsertar = LBound(ME_Tools_Acciones.accionSeleccionada, 2) + yRelativa Mod (UBound(ME_Tools_Acciones.accionSeleccionada, 2) - LBound(ME_Tools_Acciones.accionSeleccionada, 2) + 1)
                            
                            Call comandoAccion.crear(xArea, yArea, ME_Tools_Acciones.accionSeleccionada(xAreaInsertar, yAreaInsertar))
                            Call conjuntoDeComandos.Add(comandoAccion)
                        ElseIf ME_Tools_Acciones.herramientainterna = eHerramientasAccion.borrar Then
                            Call comandoAccion.crear(xArea, yArea, Nothing)
                            Call conjuntoDeComandos.Add(comandoAccion)
                        End If
                    End If
                
                    'NPCS
                    If TOOL_SELECC And Tools.tool_npc Then
                        If Button = vbLeftButton Then
                    
                            Set comandoNpc = New cComandoInsertarNPC
                                    
                            If Me_Tools_Npc.herramientaInternaNPC = eHerramientasNPC.insertar Then
                            
                                'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(Me_Tools_Npc.NPCSeleccionado, 1) + xRelativa Mod (UBound(Me_Tools_Npc.NPCSeleccionado, 1) - LBound(Me_Tools_Npc.NPCSeleccionado, 1) + 1)
                                yAreaInsertar = LBound(Me_Tools_Npc.NPCSeleccionado, 2) + yRelativa Mod (UBound(Me_Tools_Npc.NPCSeleccionado, 2) - LBound(Me_Tools_Npc.NPCSeleccionado, 2) + 1)
                                
                                'Obtengo el index de la criatura y la zona donde puede nacer
                                tempInt = Me_Tools_Npc.NPCSeleccionado(xAreaInsertar, yAreaInsertar).Index
                                tempbyte1 = Me_Tools_Npc.NPCSeleccionado(xAreaInsertar, yAreaInsertar).Zona
                                
                                If tempInt > 0 Then
                                    Call comandoNpc.crear(tempInt, tempbyte1, CByte(xArea), CByte(yArea))
                                    Call conjuntoDeComandos.Add(comandoNpc)
                                End If
                            
                            ElseIf Me_Tools_Npc.herramientaInternaNPC = eHerramientasNPC.borrar Then
                                Call comandoNpc.crear(0, 0, CByte(xArea), CByte((yArea)))
                                Call conjuntoDeComandos.Add(comandoNpc)
                            End If
                        Else
                            If mapdata(MouseTileX, MouseTileY).NpcIndex > 0 Then
                                Call frmMain.ListaConBuscadorNpcs.seleccionarID(mapdata(MouseTileX, MouseTileY).NpcIndex)
                            End If
                        End If
                    End If
                    
                    'OBJETOS
                    If TOOL_SELECC And Tools.tool_obj Then
                        If Button = vbLeftButton Then
        
                            Set comandoObjeto = New cComandoInsertarObjeto
                                                           
                            If Me_Tools_Objetos.herramientaInternaOBJ = eHerramientasOBJ.insertar Then
                                'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(Me_Tools_Objetos.objIndexSeleccionado, 1) + xRelativa Mod (UBound(Me_Tools_Objetos.objIndexSeleccionado, 1) - LBound(Me_Tools_Objetos.objIndexSeleccionado, 1) + 1)
                                yAreaInsertar = LBound(Me_Tools_Objetos.objIndexSeleccionado, 2) + yRelativa Mod (UBound(Me_Tools_Objetos.objIndexSeleccionado, 2) - LBound(Me_Tools_Objetos.objIndexSeleccionado, 2) + 1)
                                
                                'Obtengo el index del objeto
                                tempInt = Me_Tools_Objetos.objIndexSeleccionado(xAreaInsertar, yAreaInsertar)
                                
                                If tempInt > 0 Then
                                    Call comandoObjeto.crear(xArea, yArea, Me_Tools_Objetos.objCantidadSeleccionado(xAreaInsertar, yAreaInsertar), tempInt)
                                    Call conjuntoDeComandos.Add(comandoObjeto)
                                End If
                                
                            ElseIf Me_Tools_Objetos.herramientaInternaOBJ = eHerramientasOBJ.borrar Then
                                Call comandoObjeto.crear(xArea, yArea, 0, 0)
                                Call conjuntoDeComandos.Add(comandoObjeto)
                            End If
                        Else
                            If mapdata(MouseTileX, MouseTileY).OBJInfo.objIndex > 0 Then
                                Call frmMain.ListaConBuscadorObjetos.seleccionarID(mapdata(xArea, yArea).OBJInfo.objIndex)
                            End If
                        End If
                    End If
                
                    ' TRIGGERS
                    If TOOL_SELECC And Tools.tool_triggers Then
                        
                        
                        If Button = vbRightButton Then
                            For i = 0 To frmMain.lstTriggers.ListCount - 1
                                frmMain.lstTriggers.Selected(i) = (mapdata(MouseTileX, MouseTileY).Trigger And bitwisetable(i))
                            Next i
                        Else
                            If ME_Tools_Triggers.herramientaInternaTrigger = herramientasTriggers.insertar Then
                                
                                'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(ME_Tools_Triggers.triggerSeleccionado, 1) + xRelativa Mod (UBound(ME_Tools_Triggers.triggerSeleccionado, 1) - LBound(ME_Tools_Triggers.triggerSeleccionado, 1) + 1)
                                yAreaInsertar = LBound(ME_Tools_Triggers.triggerSeleccionado, 2) + yRelativa Mod (UBound(ME_Tools_Triggers.triggerSeleccionado, 2) - LBound(ME_Tools_Triggers.triggerSeleccionado, 2) + 1)
                                    
                                If ME_Tools_Triggers.triggerSeleccionado(xAreaInsertar, yAreaInsertar) > 0 Then
                                
                                    Set comando = New cComandoInsertarTrigger
            
                                    Call comando.crear(xArea, yArea, ME_Tools_Triggers.triggerSeleccionado(xAreaInsertar, yAreaInsertar))
                                    
                                    Call conjuntoDeComandos.Add(comando)
                                End If
                            ElseIf ME_Tools_Triggers.herramientaInternaTrigger = herramientasTriggers.borrar Then
                                Set comando = New cComandoInsertarTrigger
                                
                                Call comando.crear(xArea, yArea, CLng(0))
                                
                                Call conjuntoDeComandos.Add(comando)
                            End If
                        End If
                    End If
                
                    'BLOQUEOS
                    If TOOL_SELECC And Tools.tool_bloqueo Then
                        If Button = vbLeftButton Then
                             If ME_Tools_Triggers.herramientaInternaBloqueo = InsertarSimple Then
                                 Call conjuntoDeComandos.Add(ME_Tools_Triggers.BloquearLinea(CByte(xArea), CByte(yArea), ME_Tools.obtenerBordeSeleccionado(), False))
                             ElseIf ME_Tools_Triggers.herramientaInternaBloqueo = InsertarDoble Then
                                 Call conjuntoDeComandos.Add(ME_Tools_Triggers.BloquearLinea(CByte(xArea), CByte(yArea), ME_Tools.obtenerBordeSeleccionado(), True))
                             ElseIf ME_Tools_Triggers.herramientaInternaBloqueo = BorrarSimple Then
                                 Call conjuntoDeComandos.Add(ME_Tools_Triggers.DesBloquearLinea(CByte(xArea), CByte(yArea), ME_Tools.obtenerBordeSeleccionado(), False))
                             End If
                         Else
                             If ME_Tools_Triggers.herramientaInternaBloqueo = InsertarSimple Then
                                 Call conjuntoDeComandos.Add(ME_Tools_Triggers.BloquearTile(CByte(xArea), CByte(yArea), False))
                             ElseIf ME_Tools_Triggers.herramientaInternaBloqueo = InsertarDoble Then
                                 Call conjuntoDeComandos.Add(ME_Tools_Triggers.BloquearTile(CByte(xArea), CByte(yArea), True))
                             ElseIf ME_Tools_Triggers.herramientaInternaBloqueo = BorrarSimple Then
                                 Call conjuntoDeComandos.Add(ME_Tools_Triggers.DesBloquearTile(CByte(xArea), CByte(yArea)))
                             End If
                         End If
                    End If
                    
                    'Entidades
                    If TOOL_SELECC And Tools.tool_entidades Then
                        If Button = vbLeftButton Then
                            If Me_Tools_Entidades.herramientaInternaEntidades = eHerramientasEntidades.insertar Then
                                'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(Me_Tools_Entidades.entidadesSeleccionadas, 1) + xRelativa Mod (UBound(Me_Tools_Entidades.entidadesSeleccionadas, 1) - LBound(Me_Tools_Entidades.entidadesSeleccionadas, 1) + 1)
                                yAreaInsertar = LBound(Me_Tools_Entidades.entidadesSeleccionadas, 2) + yRelativa Mod (UBound(Me_Tools_Entidades.entidadesSeleccionadas, 2) - LBound(Me_Tools_Entidades.entidadesSeleccionadas, 2) + 1)
                            
                                Set comandoEntidad = New cComandoInsertarEntidad
                                
                                Call comandoEntidad.crear(xArea, yArea, Me_Tools_Entidades.entidadesSeleccionadas(xAreaInsertar, yAreaInsertar).infoEntidades)
                                Call conjuntoDeComandos.Add(comandoEntidad)
                            ElseIf Me_Tools_Entidades.herramientaInternaEntidades = eHerramientasEntidades.borrar Then
                                Set comandoEntidad = New cComandoInsertarEntidad
                            
                                Call comandoEntidad.crear(xArea, yArea, Me_Tools_Entidades.entidadesSeleccionadasBorrado.infoEntidades)
                                Call conjuntoDeComandos.Add(comandoEntidad)
                            End If
                        Else
                            'Copio
                            Call frmMain.actualizarListaEntidadesEnTile(MouseTileX, MouseTileY)
                        End If
                    End If
                End If
            
            
                If puedeModificarVisual Then
            
                    'Suelo. TileSet
                    If TOOL_SELECC And Tools.tool_tileset Then

                        If Me_Tools_TileSet.herramientaInternaTileSet = eHerramientasTileSet.insertar Then
                            'Obtego el elemento con el que voy a tratar
                            xAreaInsertar = LBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 1) + xRelativa Mod (UBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 1) - LBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 1) + 1)
                            yAreaInsertar = LBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 2) + yRelativa Mod (UBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 2) - LBound(Me_Tools_TileSet.tilesetNumeroSeleccionado, 2) + 1)
      
                            tempbyte1 = 0 ' Se puede insertar
                            
                            '¿Tiene restriccion en donde se puede insertar?
                            If Me_Tools_TileSet.TileSetSectorAncho > 0 And Me_Tools_TileSet.TileSetSectorAlto > 0 Then
                            
                                If Not (xRelativa Mod Me_Tools_TileSet.TileSetSectorAncho = (xArea + GRILLA_OFFSET_X - 1) Mod Me_Tools_TileSet.TileSetSectorAncho And yRelativa Mod Me_Tools_TileSet.TileSetSectorAlto = (yArea + GRILLA_OFFSET_Y - 1) Mod Me_Tools_TileSet.TileSetSectorAlto) Then
                                    tempbyte1 = 1
                                End If
                                
                            End If
                            
                             ' ¿Acá esta bien?
                            If tempbyte1 = 0 Then
                                tempInt = Me_Tools_TileSet.TilesetSeleccionado(xAreaInsertar, yAreaInsertar)
                                tempInt2 = Me_Tools_TileSet.tilesetNumeroSeleccionado(xAreaInsertar, yAreaInsertar)
                                                    
                                If tempInt > 0 And tempInt2 >= 0 Then
                                    Set comandoTileSet = New cComandoInsertarTileSet
                                    Call comandoTileSet.crear(CByte(xArea), CByte(yArea), tempInt, tempInt2)
                                    Call conjuntoDeComandos.Add(comandoTileSet)
                                End If
                            End If
        
                            
                        Else
                            Set comandoTileSet = New cComandoInsertarTileSet
                            Call comandoTileSet.crear(CByte(xArea), CByte(yArea), 0, 0)
                            Call conjuntoDeComandos.Add(comandoTileSet)
                        End If
                        
                    End If
        
                    'Graficos
                       
                    If TOOL_SELECC And Tools.tool_grh Then
                            
                        Set comandoGrafico = New cComandoInsertarGrafico
                            
                        If ME_Tools_Graficos.herramientaInternaGraficos = eHerramientaGraficos.insertar Then
                            
                            If Button = vbLeftButton Then
                            
                               'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) + xRelativa Mod (UBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) - LBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) + 1)
                                yAreaInsertar = LBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) + yRelativa Mod (UBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) - LBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) + 1)
                                'Inserto en la capa seleccionada
                                Call comandoGrafico.crear(ME_Tools_Graficos.grhInfoSeleccionada(xAreaInsertar, yAreaInsertar).grhInfoPosicion, CByte(xArea), CByte(yArea))
                                Call conjuntoDeComandos.Add(comandoGrafico)
                                
                                Call ME_Tools_Graficos.actualizarListaUltimosUsados(ME_Tools_Graficos.grhInfoSeleccionada(xAreaInsertar, yAreaInsertar).grhInfoPosicion)
                                
                            Else
                                'Copio
                                frmMain.lblGraficosEnPos = "Graficos en (" & MouseTileX & "," & MouseTileY & ")"
                                frmMain.lstGraficosCopiados.Clear
                                For i = 1 To CANTIDAD_CAPAS
                                    If mapdata(MouseTileX, MouseTileY).Graphic(i).GrhIndex Then
                                        frmMain.lstGraficosCopiados.AddItem "Capa " & i & ": (" & mapdata(MouseTileX, MouseTileY).Graphic(i).GrhIndex & ") "
                                    Else
                                        frmMain.lstGraficosCopiados.AddItem "Capa " & i & ": -"
                                    End If
                                Next
                            End If
                            
                        ElseIf ME_Tools_Graficos.herramientaInternaGraficos = eHerramientaGraficos.borrar Then
                                
                            If Button = vbLeftButton Then
                            
                                'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) + xRelativa Mod (UBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) - LBound(ME_Tools_Graficos.grhInfoSeleccionada, 1) + 1)
                                yAreaInsertar = LBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) + yRelativa Mod (UBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) - LBound(ME_Tools_Graficos.grhInfoSeleccionada, 2) + 1)
                                
                                'Borro en la capa seleccionad
                                For i = 1 To CANTIDAD_CAPAS
                                    If xAreaInsertar > 0 And yAreaInsertar > 0 Then
                                        If ME_Tools_Graficos.grhInfoSeleccionada(xAreaInsertar, yAreaInsertar).grhInfoPosicion(i).seleccionado Then
                                            auxGrhInfo(i).seleccionado = True
                                            auxGrhInfo(i).GrhIndex = 0
                                        End If
                                    Else
                                            auxGrhInfo(i).seleccionado = True
                                            auxGrhInfo(i).GrhIndex = 0
                                    End If
                                Next i
                                    
                                Call comandoGrafico.crear(auxGrhInfo, CByte(xArea), CByte(yArea))
                                Call conjuntoDeComandos.Add(comandoGrafico)
                                
        
                            Else
                                'Borro todo!

                                For i = 1 To CANTIDAD_CAPAS
                                    auxGrhInfo(i).seleccionado = True
                                    auxGrhInfo(i).GrhIndex = 0
                                Next i
                                    
                                Call comandoGrafico.crear(auxGrhInfo, CByte(xArea), CByte(yArea))
                                Call conjuntoDeComandos.Add(comandoGrafico)
                                     
                            End If
                        End If
                        
                    End If
                    
                    'Luces
                    If TOOL_SELECC And Tools.tool_luces Then
                    
                        If Button = vbLeftButton Then
                            
                             If Me_Tools_Luces.herramientaInternaLuces = eHerramientasLuces.insertar Then  'Insertar Luz
                                 
                                 'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(Me_Tools_Luces.infoLuzSeleccionada, 1) + xRelativa Mod (UBound(Me_Tools_Luces.infoLuzSeleccionada, 1) - LBound(Me_Tools_Luces.infoLuzSeleccionada, 1) + 1)
                                yAreaInsertar = LBound(Me_Tools_Luces.infoLuzSeleccionada, 2) + yRelativa Mod (UBound(Me_Tools_Luces.infoLuzSeleccionada, 2) - LBound(Me_Tools_Luces.infoLuzSeleccionada, 2) + 1)
                                
                                luz = Me_Tools_Luces.infoLuzSeleccionada(xAreaInsertar, yAreaInsertar)
    
                                Set comandoLuz = New cComandoInsertarLuz
                                Call comandoLuz.crear(xArea, yArea, luz)
                                Call conjuntoDeComandos.Add(comandoLuz)
                                
    '                            Pre_Render_Light MapData(xArea, yArea).luz
                
                             ElseIf Me_Tools_Luces.herramientaInternaLuces = eHerramientasLuces.borrar Then  'Borrar luz
                                 
                                 If mapdata(xArea, yArea).luz Then
                                    luz.LuzRadio = 0
                                    
                                    Set comandoLuz = New cComandoInsertarLuz
                                    Call comandoLuz.crear(xArea, yArea, luz)
                                    Call conjuntoDeComandos.Add(comandoLuz)
                                 End If
                            End If
                        Else
                            If mapdata(xArea, yArea).luz Then
                                'Copio las propiedades
                                DLL_Luces.Get_Light mapdata(xArea, yArea).luz, tempbyte1, tempbyte2, luz.LuzColor.r, luz.LuzColor.g, luz.LuzColor.b, luz.LuzRadio, luz.LuzBrillo, luz.LuzTipo, luz.luzInicio, luz.luzFin
                                Call Me_Tools_Luces.mostrarLuzEnFormulario(luz)
                                Call Me_Tools_Luces.seleccionarLuz(luz)
                            End If
                        End If
                    End If
            
                    'Particulas
                    If TOOL_SELECC And Tools.tool_particles Then
                        If Button = vbRightButton Then 'Copiar
                        
                            With mapdata(xArea, yArea)
                                Call ME_Tools_Particulas.establecerParticula(.Particles_groups)
                            End With
                            
                            ValidarParticulasSeleccionadas
                            
                        Else
                                                             
                            If ME_Tools_Particulas.herramientaInternaParticula = eHerramientasParticulas.insertar Then
                                'Obtego el elemento con el que voy a tratar
                                xAreaInsertar = LBound(ME_Tools_Particulas.infoParticulasSeleccion, 1) + xRelativa Mod (UBound(ME_Tools_Particulas.infoParticulasSeleccion, 1) - LBound(ME_Tools_Particulas.infoParticulasSeleccion, 1) + 1)
                                yAreaInsertar = LBound(ME_Tools_Particulas.infoParticulasSeleccion, 2) + yRelativa Mod (UBound(ME_Tools_Particulas.infoParticulasSeleccion, 2) - LBound(ME_Tools_Particulas.infoParticulasSeleccion, 2) + 1)
                            
                                For i = 0 To 2
                                  If Not ME_Tools_Particulas.infoParticulasSeleccion(xAreaInsertar, yAreaInsertar).particulaSeleccionada(i) Is Nothing Then
                                        Set comandoParticula = New cComandoInsertarParticula
                                        comandoParticula.crear i, xArea, yArea, ME_Tools_Particulas.infoParticulasSeleccion(xAreaInsertar, yAreaInsertar).particulaSeleccionada(i).PGID
                                        Call conjuntoDeComandos.Add(comandoParticula)
                                    End If
                                Next i
                            
                            ElseIf ME_Tools_Particulas.herramientaInternaParticula = eHerramientasParticulas.borrar Then
                            
                                For i = 0 To 2
                                    Set comandoParticula = New cComandoInsertarParticula
                                    Call comandoParticula.crear(i, xArea, yArea, -1)
                                    Call conjuntoDeComandos.Add(comandoParticula)
                                Next i
                                
                            End If
                        End If
                    End If
                End If
            End If
        Next yArea
    Next xArea
    
    
    'Genere más de un comando'
    If conjuntoDeComandos.count > 0 Then
        If conjuntoDeComandos.count = 1 Then
            Call ME_Tools.ejecutarComando(conjuntoDeComandos.item(1))
        Else
            Dim comandoCompuesto As cComandoCompuesto
            Set comandoCompuesto = New cComandoCompuesto
            Call comandoCompuesto.crear(conjuntoDeComandos, ME_Tools.tool_act_name)
            Call ME_Tools.ejecutarComando(comandoCompuesto)
        End If
        miniMap_Redraw
    End If
    
    'Actualizo la vista
    'Map_render_2array
    
    rm2a
    Cachear_Tiles = True
    
    
    
End Sub

Public Sub selec_TOOL()

ME_Tools.editando_montaña = False
'ME_Tools.editando_sol = False

TipoEditorParticulas = False

frmMain.ckbMostrarAcciones.Enabled = True
frmMain.cmdInsertarNpc.Enabled = True
frmMain.cmdInsertarObjeto.Enabled = True

    If (TOOL_SELECC And Tools.tool_bloqueo) Then
        'Si es bloqueo pero no tengo activada la opción de ver bloqueos, la activo temporalmente
        If Not DRAWBLOQUEOS = vbChecked Then
            DRAWBLOQUEOS = vbGrayed
        End If
        HerramientaIndiceInternoMaximo = 5
        HerramientaIndiceInterno = HerramientaIndiceInterno Mod HerramientaIndiceInternoMaximo
    Else
        ' Si no es la de bloqueos, me fijo si tenia una activacion temporal.
        If DRAWBLOQUEOS = vbGrayed Then
            DRAWBLOQUEOS = vbUnchecked
        End If
    End If
    
    If (TOOL_SELECC And Tools.tool_triggers) Then
        'Si es triggers pero no tengo activada la opción de ver triggers, la activo temporalmente
        If Not DRAWTRIGGERS = vbChecked Then
            DRAWTRIGGERS = vbGrayed
        End If
    Else
        ' Si no es la de bloqueos, me fijo si tenia una activacion temporal.
        If DRAWTRIGGERS = vbGrayed Then
            DRAWTRIGGERS = vbUnchecked
        End If
    End If
       
    If TOOL_SELECC And Tools.tool_montaña Then
        Backup_HM
    End If
    
    If TOOL_SELECC And Tools.tool_particles Then
        If VentanaSelectorParticulas Is Nothing Then
            Set VentanaSelectorParticulas = New vw_Part_Select
        End If
        
        GUI_SetFocus VentanaSelectorParticulas
    Else
        If Not VentanaSelectorParticulas Is Nothing Then
            GUI_Quitar VentanaSelectorParticulas
        End If
    End If
    
    
frmMain.ctriggers.value = DRAWTRIGGERS
frmMain.ver_triggers.checked = DRAWTRIGGERS = vbChecked
frmMain.Bloqueos.value = DRAWBLOQUEOS
frmMain.ver_bloqueos.checked = DRAWBLOQUEOS = vbChecked

tool_act_name = Split(toolnames, "-")(TOOL_SELECC Mod (toolcount + 1))
End Sub

'Public Sub RENDER_SUN()
'
'    Dim tBottom!, tRight! ', tTop!, tLeft!
'    Static inta As Single
'    Dim ll As Long
'    Dim TGRH As GrhData
'    Dim altU As AUDT
'
'    Dim Color As Long
'
'    Color = &H7FFFFFFF
'
'
'    Call GetTexture(0)
'
'        Dim tBox As Box_Vertex
'        With tBox
'                .x0 = sunposee.X
'                .y0 = sunposee.Y * 0.7
'
'                .x1 = sunpose.X
'                .y1 = sunpose.Y * 0.7
'
'                .x2 = .x1
'                .y2 = .y1 - sunposa.Y * 0.7 + 2
'
'                .Color2 = Color
'                .color1 = Color
'                .color0 = Color
'                .tu0 = 0
'                .tv0 = 0
'                .tu1 = 0
'                .tv1 = 0
'                .tu2 = 0
'                .tv2 = 0
'        End With
'
'        'D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, tBox, TL_size
'        'D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
'        If inta = 0 Then inta = 138
'        inta = inta + 1 * Engine.timerElapsedTime * 0.025
'        If inta > 152 Then inta = 139
'        'Render_Particle CInt(tBox.x2), CInt(tBox.y2), inta, , 64
'        If inta > 152 Then inta = 138
'
'        'MapData(map_x, map_y).tile_render = 255
'
'
'
'End Sub

Public Sub calcular_montaña(ya As Integer)
Dim x%, y%, d!, j!
    Dim caja As RECT
    
    caja.top = maxl(clickpos.y - radio_montana, Y_MINIMO_VISIBLE)
    caja.left = maxl(clickpos.x - radio_montana, X_MINIMO_VISIBLE)
    caja.bottom = minl(clickpos.y + radio_montana, Y_MAXIMO_VISIBLE)
    caja.right = minl(clickpos.x + radio_montana, X_MAXIMO_VISIBLE)

Dim puedoY As Boolean
Dim puedoX As Boolean


Select Case mt_select
Case mtools.mt_clean
    For x = caja.left To caja.right
        For y = caja.top To caja.bottom
            d = radio_montana - Sqr(((clickpos.y - y) * (clickpos.y - y) + (clickpos.x - x) * (clickpos.x - x)))
            
            
            If d > 0 And esPosicionJugable(x, y) Then
                puedoX = InMapBounds(x - 1, y)
                puedoY = InMapBounds(x, y + 1)
                
                d = 0
                hMapData(x, y).hs(0) = d
                If puedoY Then
                    hMapData(x, y + 1).hs(1) = d
                    If puedoX Then hMapData(x - 1, y + 1).hs(3) = d
                End If
                If puedoX Then hMapData(x - 1, y).hs(2) = d
            End If

        Next y
    Next x
Case mtools.mt_promedio

        For x = caja.left To caja.right
            For y = caja.top To caja.bottom
                d = radio_montana - Sqr(((clickpos.y - y) * (clickpos.y - y) + (clickpos.x - x) * (clickpos.x - x)))
                If d > 0 And esPosicionJugable(x, y) Then
                    d = hMapDataORIGINAL(x, y).hs(0) + d * ya
                    If d < 255 And d >= 0 Then
                        puedoX = InMapBounds(x - 1, y)
                        puedoY = InMapBounds(x, y + 1)
                    
                        hMapData(x, y).hs(0) = d
                        If puedoY Then
                            hMapData(x, y + 1).hs(1) = d
                            If puedoX Then hMapData(x - 1, y + 1).hs(3) = d
                        End If
                        If puedoX Then hMapData(x - 1, y).hs(2) = d
                    End If
                End If
            Next y
        Next x
Case mtools.mp_suma
        For x = caja.left To caja.right
            For y = caja.top To caja.bottom
                d = radio_montana - Sqr(((clickpos.y - y) * (clickpos.y - y) + (clickpos.x - x) * (clickpos.x - x)))
                If d > 0 And esPosicionJugable(x, y) Then
                    d = hMapDataORIGINAL(x, y).hs(0) + ya
                    If d < 255 And d >= -255 Then
                        puedoX = InMapBounds(x - 1, y)
                        puedoY = InMapBounds(x, y + 1)
                        
                        hMapData(x, y).hs(0) = d
                        If puedoY Then
                            hMapData(x, y + 1).hs(1) = d
                            If puedoX Then hMapData(x - 1, y + 1).hs(3) = d
                        End If
                        If puedoX Then hMapData(x - 1, y).hs(2) = d
                    End If
                End If
            Next y
        Next x
Case mtools.mt_pie
        x = clickpos.x
        y = clickpos.y
        d = AlturaPieORIGINAL(x, y) + ya
        If d < 255 And d >= -255 Then
            AlturaPie(x, y) = d
        End If

    
Case mtools.mt_slerp
        For x = caja.left To caja.right
            For y = caja.top To caja.bottom
                d = Sqr(((clickpos.y - y) * (clickpos.y - y) + (clickpos.x - x) * (clickpos.x - x)))
                If d <= radio_montana And esPosicionJugable(x, y) Then
                    d = d / radio_montana
                    j = CosInterp(maxs(hMapDataORIGINAL(x, y).hs(0), Abs(ya)), mins(hMapDataORIGINAL(x, y).hs(0), Abs(ya)), d)
                    'If d < 255 And d >= 0 Then
                    If j < 255 And j > 0 Then
                        If ya > 0 Then
                            hMapData(x, y).hs(0) = CosInterp(hMapDataORIGINAL(x, y).hs(0) + j, hMapDataORIGINAL(x, y).hs(0), d)
                        Else
                            hMapData(x, y).hs(0) = CosInterp(hMapDataORIGINAL(x, y).hs(0) - j, hMapDataORIGINAL(x, y).hs(0), d)
                        End If
                        
                        puedoX = InMapBounds(x - 1, y)
                        puedoY = InMapBounds(x, y + 1)

                        If puedoY Then
                            hMapData(x, y + 1).hs(1) = hMapData(x, y).hs(0)
                            If puedoX Then hMapData(x - 1, y + 1).hs(3) = hMapData(x, y).hs(0)
                        End If
                        If puedoX Then hMapData(x - 1, y).hs(2) = hMapData(x, y).hs(0)
                    End If
                End If
            Next y
        Next x
Case mtools.mt_meseta
        For x = caja.left To caja.right
            For y = caja.top To caja.bottom
                d = radio_montana - Sqr(((clickpos.y - y) * (clickpos.y - y) + (clickpos.x - x) * (clickpos.x - x)))
                If d > 0 Then
                    d = hMapDataORIGINAL(x, y).hs(0) + ya
                    If d < 255 And d >= 0 And esPosicionJugable(x, y) Then
                        puedoX = InMapBounds(x - 1, y)
                        puedoY = InMapBounds(x, y + 1)
                    
                        hMapData(x, y).hs(0) = d
                        If puedoY Then
                            hMapData(x, y + 1).hs(1) = d
                            If puedoX Then hMapData(x - 1, y + 1).hs(3) = d
                        End If
                        If puedoX Then hMapData(x - 1, y).hs(2) = d
                    End If
                End If
            Next y
        Next x
Case mtools.mt_blur
        For x = caja.left To caja.right
            For y = caja.top To caja.bottom
                If InMapBounds(x, y) And InMapBounds(x - 1, y + 1) And esPosicionJugable(x, y) Then
                    d = hMapData(x, y).hs(3)
                    d = d + hMapData(x, y + 1).hs(2)
                    d = d + hMapData(x - 1, y + 1).hs(0)
                    d = d + hMapData(x - 1, y).hs(1)
                    d = d + hMapData(x, y).hs(0)
                    d = d + hMapData(x, y + 1).hs(1)
                    d = d + hMapData(x - 1, y + 1).hs(3)
                    d = d + hMapData(x - 1, y).hs(2)
                    d = d / 8
                    d = d + hMapData(x, y).hs(0)
                    d = d / 2
                    hMapData(x, y).hs(0) = d
                    hMapData(x, y + 1).hs(1) = d
                    hMapData(x - 1, y + 1).hs(3) = d
                    hMapData(x - 1, y).hs(2) = d
                End If
            Next y
        Next x
End Select
End Sub

Public Sub Backup_HM()
    Dim tamaño_ As Long

    tamaño_ = Len(hMapData(1, 1)) * TILES_POR_MAPA
    Call DXCopyMemory(hMapDataORIGINAL(1, 1), hMapData(1, 1), tamaño_)
'    tamaño_ = Len(Alturas(1, 1)) * TILES_POR_MAPA
'    Call DXCopyMemory(alturasORIGINAL(1, 1), Alturas(1, 1), tamaño_)
'    tamaño_ = Len(NormalData(1, 1)) * TILES_POR_MAPA
'    Call DXCopyMemory(NormalDataORIGINAL(1, 1), NormalData(1, 1), tamaño_)
'    tamaño_ = 2 * TILES_POR_MAPA
'    Call DXCopyMemory(AlturaPieORIGINAL(1, 1), AlturaPie(1, 1), tamaño_)
End Sub

Public Sub Restore_HM()
    Dim tamaño_ As Long

    tamaño_ = Len(hMapData(1, 1)) * TILES_POR_MAPA
    Call DXCopyMemory(hMapData(1, 1), hMapDataORIGINAL(1, 1), tamaño_)
'    tamaño_ = Len(Alturas(1, 1)) * TILES_POR_MAPA
'    Call DXCopyMemory(Alturas(1, 1), alturasORIGINAL(1, 1), tamaño_)
'    tamaño_ = Len(NormalData(1, 1)) * TILES_POR_MAPA
'    Call DXCopyMemory(NormalData(1, 1), NormalDataORIGINAL(1, 1), tamaño_)
'    tamaño_ = 2 * TILES_POR_MAPA
'    Call DXCopyMemory(AlturaPie(1, 1), AlturaPieORIGINAL(1, 1), tamaño_)
End Sub

Public Sub Grh_Render_PIE(ByVal Color As Long, ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte)
'*********************************************
'Author: menduz
'*********************************************
    Dim tBottom!, tRight! ', tTop!, tLeft!
    
    'If GrhIndex = 0 Then Exit Sub
    If map_x < X_MINIMO_VISIBLE Or map_y < Y_MINIMO_VISIBLE Then Exit Sub
    If map_x > X_MAXIMO_VISIBLE Or map_y > Y_MAXIMO_VISIBLE Then Exit Sub
    Dim tBox As Box_Vertex
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse PeekTexture(16500)
    

        
        tTop = tTop + 16
        tLeft = tLeft + 16
        tBottom = tTop + 32
        tRight = tLeft + 32
        
        
        With tBox 'With tBox
                .x0 = tLeft
                .y0 = tBottom - AlturaPie(map_x, map_y)
                .color0 = Color
                .x1 = tLeft
                .y1 = tTop - AlturaPie(map_x, map_y - 1)
                .Color1 = Color
                .X2 = tRight
                .Y2 = tBottom - AlturaPie(map_x + 1, map_y)
                .Color2 = Color
                .x3 = tRight
                .y3 = tTop - AlturaPie(map_x + 1, map_y - 1)
                .color3 = Color
                .tu0 = 0
                .tv0 = 1
                .tu1 = 0
                .tv1 = 0
                .tu2 = 1
                .tv2 = 1
                .tu3 = 1
                .tv3 = 0
                .rhw0 = 1
                .rhw1 = 1
                .rhw2 = 1
                .rhw3 = 1
        End With
        
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size

End Sub

Public Sub Grh_Render_PIE_Tool(ByVal tLeft As Single, ByVal tTop As Single, ByVal map_x As Byte, ByVal map_y As Byte)
'*********************************************
'Author: menduz
'*********************************************
    Dim tBottom!, tRight! ', tTop!, tLeft!
    


    'If GrhIndex = 0 Then Exit Sub
    If map_x < X_MINIMO_VISIBLE Or map_y < Y_MINIMO_VISIBLE Then Exit Sub
    If map_x > X_MAXIMO_VISIBLE Or map_y > Y_MAXIMO_VISIBLE Then Exit Sub
    Dim tBox As Box_Vertex
    Dim Color As Long
    Color = &H7F00FF00
    
    Engine_PixelShaders.Engine_PixelShaders_SetTexture_Diffuse Nothing
    

        
        tTop = tTop + 8
        tLeft = tLeft + 8
        tBottom = tTop + 16
        tRight = tLeft + 16
        
        
        With tBox 'With tBox
                .x0 = tLeft
                .y0 = tBottom - AlturaPie(map_x, map_y)
                .color0 = Color
                .x1 = tLeft
                .y1 = tTop - AlturaPie(map_x, map_y)
                .Color1 = Color
                .X2 = tRight
                .Y2 = tBottom - AlturaPie(map_x, map_y)
                .Color2 = Color
                .x3 = tRight
                .y3 = tTop - AlturaPie(map_x, map_y)
                .color3 = Color
                .tu0 = 0
                .tv0 = 1
                .tu1 = 0
                .tv1 = 0
                .tu2 = 1
                .tv2 = 1
                .tu3 = 1
                .tv3 = 0
                .rhw0 = 1
                .rhw1 = 1
                .rhw2 = 1
                .rhw3 = 1
        End With
        
        Engine_PixelShaders.Engine_PixelShaders_Utilizar ePixelShaders.Ninguno
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tBox, TL_size

End Sub



Sub init_map_editor()

    InitChars
    Init_weapons
    
    tileset_actual = 1
    tileset_actual_virtual = 0
'
'    j = NextOpenChar
'
'    Call MakeChar(j, 1, 1, SOUTH, 10, 10, 0, 0, 0)
'
'    UserCharIndex = j
'
'    'CharList(j).active = True
'
'    ActivateChar j
'
'    CharList(j).nombre = "Mapeador"
CrearCharWalkMode
'CharList(UserCharIndex).Velocidad.x = 40
'CharList(UserCharIndex).Velocidad.y = 40
    'CharList(j).Velocidad.X = 20
    'CharList(j).Velocidad.y = 20
End Sub

Public Sub Quitar_Capa(ByVal Capa As Byte)
If EditWarning Then Exit Sub

Dim y As Integer
Dim x As Integer


For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
            mapdata(x, y).Graphic(Capa).GrhIndex = 0
    Next x
Next y


End Sub


Public Sub Quitar_NPCs(ByVal Hostiles As Boolean)

If EditWarning Then Exit Sub

Dim y As Integer
Dim x As Integer

For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
    For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        If mapdata(x, y).NpcIndex > 0 Then
            If (Hostiles = True And mapdata(x, y).NpcIndex >= 500) Or (Hostiles = False And mapdata(x, y).NpcIndex < 500) Then
                Call EraseIndexChar(CharMap(x, y))
                Call EraseChar(CharMap(x, y))
                mapdata(x, y).NpcIndex = 0
                CharMap(x, y) = 0
            End If
        End If
    Next x
Next y

End Sub

Public Function obtenerBordeSeleccionado() As E_Heading

    Dim relativaX As Double
    Dim relativaY As Double
    Dim distanciaMinima As Double
    
    relativaX = frmMain.MouseX Mod modPantalla.PixelesPorTile.x
    relativaY = frmMain.MouseY Mod modPantalla.PixelesPorTile.y
    
    'Suponemos que la minima es la distancia del oeste
    distanciaMinima = relativaX
    obtenerBordeSeleccionado = E_Heading.WEST
    
    'Sur
    If distanciaMinima >= Abs(relativaY) Then
        distanciaMinima = relativaY
        obtenerBordeSeleccionado = E_Heading.NORTH
    End If
    
    'Este
    If distanciaMinima >= Abs(relativaX - 31) Then
        distanciaMinima = Abs(relativaX - 31)
        obtenerBordeSeleccionado = E_Heading.EAST
    End If
    
    'Sur. EL X,Y del mouse es al revez de l
    If distanciaMinima >= Abs(relativaY - 31) Then
        distanciaMinima = Abs(relativaY - 31)
        obtenerBordeSeleccionado = E_Heading.SOUTH
    End If

End Function

Public Sub establecerAmpliacionDeArea(AmpliarAncho As Integer, AmpliarAlto As Integer)
    ME_Tools.TOOL_SELECC = (ME_Tools.TOOL_SELECC Or Tools.tools_seleccionMinima)
    ME_Tools.areaseleccioandaAmpliacionAlto = AmpliarAlto
    ME_Tools.areaSeleccionadaAmpliacionAncho = AmpliarAncho
End Sub

