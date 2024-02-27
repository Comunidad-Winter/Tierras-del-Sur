Attribute VB_Name = "Me_Tools_Seleccion"
'Modulo que brinda las herramientas de
'Copiar
'Pegar
'Eliminar
'Cortar (Copiar + Pegar + Eliminar Origen)
Option Explicit

Private copia() As MapBlock

Private Const TAMANIO_PORTAPAPELES = 5

Private Type tElementoPortapapeles
    vacio As Boolean
    nombre As String
    informacion() As MapBlock
End Type

Public portapapeles() As tElementoPortapapeles

Private backupHerramientaInternaTrigger As Byte
Private backupHerramientaInternaOBJ As Byte
Private backupHerramientaInternaNPC As Byte
Private backupHerramientainternaAcciones As Byte
Private backupHerramientaInternaGraficos As Byte
Private backupHerramientaInternaTileSet As Byte
Private backupHerramientaInternaParticula As Byte
Private backupHerramientaInternaLuces As Byte
Private backupHerramientaInternaEntidades As Byte

Private backup_ObjetoIndexSeleccionado() As Integer
Private backup_ObjetoCantidadSeleccionado() As Integer
Private backup_triggerSeleccionado() As Long
Private backup_NPCSeleccionado() As tNPCSeleccionado
Private backup_accionSeleccionada() As iAccionEditor
Private backup_infoParticulasSeleccion() As tParticulaSeleccionada
Private backup_infoLuzSeleccionada() As tLuzPropiedades
Private backup_grhInfoSeleccionada() As tGhInfoSeleccionada
Private backup_TilesetSeleccionado() As Integer
Private backup_tilesetNumeroSeleccionado() As Integer
Private backup_EntidadesSeleccionadas() As tEntidadSeleccionada

Public copiando As Boolean

Public Sub iniciarPortapapeles()

    ReDim portapapeles(1 To TAMANIO_PORTAPAPELES) As tElementoPortapapeles
    
    Dim loopP As Integer
    
    For loopP = 1 To TAMANIO_PORTAPAPELES
        portapapeles(loopP).vacio = True
    Next
    
End Sub

Private Function obtenerIndiceLibre() As Byte
    Dim loopP As Integer

    For loopP = TAMANIO_PORTAPAPELES To 2 Step -1
            portapapeles(loopP) = portapapeles(loopP - 1)
    Next loopP
   
    obtenerIndiceLibre = 1
End Function

'Agregar al portapapeles el area seleccionada
Public Sub cargarAlPortapeles(area As tAreaSeleccionada)
    
    Dim copia() As MapBlock
    Dim x As Integer
    Dim y As Integer
    Dim indicePortapapeles As Integer
    
    ReDim copia(area.izquierda To area.derecha, area.arriba To area.abajo)

    For x = area.izquierda To area.derecha
        For y = area.arriba To area.abajo
            copia(x, y) = mapdata(x, y)
        Next y
    Next x
    
    indicePortapapeles = obtenerIndiceLibre
    
    portapapeles(indicePortapapeles).informacion = copia
    portapapeles(indicePortapapeles).vacio = False
    portapapeles(indicePortapapeles).nombre = "(" & area.izquierda & ", " & area.arriba & ") a ( " & area.derecha & ", " & area.abajo & ")"
End Sub

Public Sub backupearEstadoHerramientas()
    backupHerramientaInternaTrigger = ME_Tools_Triggers.herramientaInternaTrigger
    backupHerramientaInternaOBJ = Me_Tools_Objetos.herramientaInternaOBJ
    backupHerramientaInternaNPC = Me_Tools_Npc.herramientaInternaNPC
    backupHerramientainternaAcciones = ME_Tools_Acciones.herramientainterna
    backupHerramientaInternaGraficos = ME_Tools_Graficos.herramientaInternaGraficos
    backupHerramientaInternaTileSet = Me_Tools_TileSet.herramientaInternaTileSet
    backupHerramientaInternaParticula = ME_Tools_Particulas.herramientaInternaParticula
    backupHerramientaInternaLuces = Me_Tools_Luces.herramientaInternaLuces
    backupHerramientaInternaEntidades = Me_Tools_Entidades.herramientaInternaEntidades
    
    backup_ObjetoIndexSeleccionado = Me_Tools_Objetos.objIndexSeleccionado
    backup_ObjetoCantidadSeleccionado = Me_Tools_Objetos.objCantidadSeleccionado
    
    backup_triggerSeleccionado() = ME_Tools_Triggers.triggerSeleccionado
    backup_NPCSeleccionado() = Me_Tools_Npc.NPCSeleccionado
    backup_accionSeleccionada() = ME_Tools_Acciones.accionSeleccionada
    backup_infoParticulasSeleccion() = ME_Tools_Particulas.infoParticulasSeleccion
    backup_infoLuzSeleccionada() = Me_Tools_Luces.infoLuzSeleccionada
    backup_grhInfoSeleccionada() = ME_Tools_Graficos.grhInfoSeleccionada
    backup_TilesetSeleccionado() = Me_Tools_TileSet.TilesetSeleccionado
    backup_tilesetNumeroSeleccionado() = Me_Tools_TileSet.tilesetNumeroSeleccionado
    'backup_EntidadesSeleccionadas() = Me_Tools_Entidades.entidadesSeleccionadas
End Sub

Public Sub restablecerBackupHerramientas()
    ME_Tools_Triggers.herramientaInternaTrigger = backupHerramientaInternaTrigger
    Me_Tools_Objetos.herramientaInternaOBJ = backupHerramientaInternaOBJ
    Me_Tools_Npc.herramientaInternaNPC = backupHerramientaInternaNPC
    ME_Tools_Acciones.herramientainterna = backupHerramientainternaAcciones
    ME_Tools_Graficos.herramientaInternaGraficos = backupHerramientaInternaGraficos
    ME_Tools_Particulas.herramientaInternaParticula = backupHerramientaInternaParticula
    Me_Tools_Luces.herramientaInternaLuces = backupHerramientaInternaLuces
    Me_Tools_TileSet.herramientaInternaTileSet = backupHerramientaInternaTileSet
    'Me_Tools_Entidades.herramientaInternaEntidades = backupHerramientaInternaEntidades
    
    Me_Tools_Objetos.objIndexSeleccionado = backup_ObjetoIndexSeleccionado
    Me_Tools_Objetos.objCantidadSeleccionado = backup_ObjetoCantidadSeleccionado
    ME_Tools_Triggers.triggerSeleccionado = backup_triggerSeleccionado()
    Me_Tools_Npc.NPCSeleccionado = backup_NPCSeleccionado()
    ME_Tools_Acciones.accionSeleccionada = backup_accionSeleccionada()
    ME_Tools_Particulas.infoParticulasSeleccion = backup_infoParticulasSeleccion()
    Me_Tools_Luces.infoLuzSeleccionada = backup_infoLuzSeleccionada()
    ME_Tools_Graficos.grhInfoSeleccionada = backup_grhInfoSeleccionada()
    Me_Tools_TileSet.TilesetSeleccionado = backup_TilesetSeleccionado()
    Me_Tools_TileSet.tilesetNumeroSeleccionado = backup_tilesetNumeroSeleccionado()
  '  Me_Tools_Entidades.entidadesSeleccionadas = backup_EntidadesSeleccionadas()
    
    copiando = False
End Sub

Public Function copiarDesdeMapBlock(info() As MapBlock, fuenteLuces As LucesManager, elementos As Tools) As Long
    
    Dim x As Integer
    Dim y As Integer
    Dim loopCapa As Integer
    Dim area As tAreaSeleccionada
    Dim toolAux As Tools
    Dim loopEntidad As Byte
    Dim idEntidad As Byte
    
    copiando = True
    'Activo las herramientas
    'toolAux = (0 Or Tools.tool_triggers)
   ' toolAux = (toolAux Or Tools.tool_obj)
   ' toolAux = (toolAux Or Tools.tool_npc)
   ' toolAux = (toolAux Or Tools.tool_acciones)
  '  toolAux = (toolAux Or Tools.tool_tileset)
  '  toolAux = (toolAux Or Tools.tool_grh)
 '  toolAux = (toolAux Or Tools.tool_particles)
   ' toolAux = (toolAux Or Tools.tool_luces)
    toolAux = elementos
    toolAux = (toolAux Or Tools.tool_copiar)
    
    'Guardamos para luego poder restablecer
    Call backupearEstadoHerramientas
    
    ME_Tools_Triggers.herramientaInternaTrigger = herramientasTriggers.insertar
    Me_Tools_Objetos.herramientaInternaOBJ = eHerramientasOBJ.insertar
    Me_Tools_Npc.herramientaInternaNPC = eHerramientasNPC.insertar
    ME_Tools_Acciones.herramientainterna = eHerramientasAccion.insertar
    ME_Tools_Graficos.herramientaInternaGraficos = eHerramientaGraficos.insertar
    ME_Tools_Particulas.herramientaInternaParticula = eHerramientasParticulas.insertar
    Me_Tools_Luces.herramientaInternaLuces = eHerramientasLuces.insertar
    Me_Tools_TileSet.herramientaInternaTileSet = eHerramientasTileSet.insertar
    Me_Tools_TileSet.TileSetSectorAncho = 0
    Me_Tools_TileSet.TileSetSectorAlto = 0
    
   ' Me_Tools_Entidades.herramientaInternaEntidades = eHerramientasEntidades.insertar
    
    'Obtengo el tamanio del area del portapapeles
    area.derecha = UBound(info, 1)
    area.izquierda = LBound(info, 1)
    area.arriba = LBound(info, 2)
    area.abajo = UBound(info, 2)
    
    'Obtengo memoria
    ReDim ME_Tools_Triggers.triggerSeleccionado(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim Me_Tools_Objetos.objIndexSeleccionado(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim Me_Tools_Objetos.objCantidadSeleccionado(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim Me_Tools_Npc.NPCSeleccionado(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim ME_Tools_Acciones.accionSeleccionada(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim ME_Tools_Particulas.infoParticulasSeleccion(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim Me_Tools_Luces.infoLuzSeleccionada(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim ME_Tools_Graficos.grhInfoSeleccionada(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim Me_Tools_TileSet.TilesetSeleccionado(area.izquierda To area.derecha, area.arriba To area.abajo)
    ReDim Me_Tools_TileSet.tilesetNumeroSeleccionado(area.izquierda To area.derecha, area.arriba To area.abajo)
    'ReDim Me_Tools_Entidades.entidadesSeleccionadas(area.izquierda To area.derecha, area.arriba To area.abajo)
    
    'Seteo lo que voy a agregar
    For x = area.izquierda To area.derecha
        For y = area.arriba To area.abajo
            ME_Tools_Triggers.triggerSeleccionado(x, y) = info(x, y).Trigger
            Me_Tools_Objetos.objIndexSeleccionado(x, y) = info(x, y).OBJInfo.objIndex
            Me_Tools_Objetos.objCantidadSeleccionado(x, y) = info(x, y).OBJInfo.Amount
            
            Me_Tools_Npc.NPCSeleccionado(x, y).Index = info(x, y).NpcIndex
            Me_Tools_Npc.NPCSeleccionado(x, y).Zona = info(x, y).NpcZona
             
            Me_Tools_TileSet.tilesetNumeroSeleccionado(x, y) = info(x, y).tile_number
            Me_Tools_TileSet.TilesetSeleccionado(x, y) = info(x, y).tile_texture

            For loopCapa = 1 To CANTIDAD_CAPAS
                If info(x, y).Graphic(loopCapa).GrhIndex > 0 Then
                    ME_Tools_Graficos.grhInfoSeleccionada(x, y).grhInfoPosicion(loopCapa).seleccionado = True
                    ME_Tools_Graficos.grhInfoSeleccionada(x, y).grhInfoPosicion(loopCapa).GrhIndex = info(x, y).Graphic(loopCapa).GrhIndex
                Else
                    ME_Tools_Graficos.grhInfoSeleccionada(x, y).grhInfoPosicion(loopCapa).seleccionado = False
                End If
            Next
            
            For loopCapa = 0 To 2
                Set ME_Tools_Particulas.infoParticulasSeleccion(x, y).particulaSeleccionada(loopCapa) = info(x, y).Particles_groups(loopCapa)
            Next
            
            With Me_Tools_Luces.infoLuzSeleccionada(x, y)
                If info(x, y).luz > 0 Then
                    fuenteLuces.Get_Light info(x, y).luz, CByte(x), CByte(y), .LuzColor.r, .LuzColor.g, .LuzColor.b, .LuzRadio, .LuzBrillo, .LuzTipo, .luzInicio, .luzFin
                End If
            End With
             
                
            Set ME_Tools_Acciones.accionSeleccionada(x, y) = info(x, y).accion
        Next y
    Next x
    
    copiarDesdeMapBlock = toolAux
End Function

Public Sub copiar(numeroElementoPortapapeles As Integer, elementos As Tools)
    Dim area As tAreaSeleccionada
    Dim toolAux As Long
    
    area.derecha = UBound(portapapeles(numeroElementoPortapapeles).informacion, 1)
    area.izquierda = LBound(portapapeles(numeroElementoPortapapeles).informacion, 1)
    area.arriba = LBound(portapapeles(numeroElementoPortapapeles).informacion, 2)
    area.abajo = UBound(portapapeles(numeroElementoPortapapeles).informacion, 2)

    toolAux = copiarDesdeMapBlock(portapapeles(numeroElementoPortapapeles).informacion, DLL_Luces, elementos)
    
    Call ME_Tools.selectToolMultiple(toolAux, "Copiar desde (" & area.arriba & "," & area.izquierda & ") hasta (" & area.abajo & "," & area.derecha & ")")
    
    'Le aviso al sistema que el area no es de 1x1 sino que es de otro tamaño.
    Call ME_Tools.establecerAmpliacionDeArea(area.derecha - area.izquierda + 1, area.abajo - area.arriba + 1)
    
End Sub

Public Sub pegar(area As tAreaSeleccionada)
    Call click_tool(vbLeftButton)
End Sub

Private Function cargarHerramientasBorrado(elementos As Tools) As Long
    Dim toolAux As Tools
    
    'Activo las herramientas
    'toolAux = (0 Or Tools.tool_triggers)
   ' toolAux = (toolAux Or Tools.tool_seleccion)
   ' toolAux = (toolAux Or Tools.tool_obj)
   ' toolAux = (toolAux Or Tools.tool_copiar)
  '  toolAux = (toolAux Or Tools.tool_npc)
  '  toolAux = (toolAux Or Tools.tool_acciones)
  '  toolAux = (toolAux Or Tools.tool_tileset)
  '  toolAux = (toolAux Or Tools.tool_grh)
  '  toolAux = (toolAux Or Tools.tool_particles)
  '  toolAux = (toolAux Or Tools.tool_entidades)
    
    toolAux = elementos
    toolAux = (toolAux Or Tools.tool_seleccion)
    toolAux = (toolAux Or Tools.tool_copiar)
    
    
    ME_Tools_Triggers.herramientaInternaTrigger = herramientasTriggers.borrar
    Me_Tools_Objetos.herramientaInternaOBJ = eHerramientasOBJ.borrar
    Me_Tools_Npc.herramientaInternaNPC = eHerramientasNPC.borrar
    ME_Tools_Acciones.herramientainterna = eHerramientasAccion.borrar
    ME_Tools_Graficos.herramientaInternaGraficos = eHerramientaGraficos.borrar
    Me_Tools_TileSet.herramientaInternaTileSet = eHerramientasTileSet.borrar
    ME_Tools_Particulas.herramientaInternaParticula = eHerramientasParticulas.borrar
    Me_Tools_Luces.herramientaInternaLuces = eHerramientasLuces.borrar
    Me_Tools_Entidades.herramientaInternaEntidades = eHerramientasEntidades.borrar
    
    ME_Tools_Graficos.setGrhInfoBorrado
    
    cargarHerramientasBorrado = toolAux
End Function
Public Sub eliminar(area As tAreaSeleccionada, elementos As Tools)
    Dim tool As Long
    Dim area_vieja As tAreaSeleccionada
    
    area_vieja = areaSeleccionada
    
    areaSeleccionada = area
    backupearEstadoHerramientas
    
    tool = cargarHerramientasBorrado(elementos)
    Call ME_Tools.selectToolMultiple(tool, "Borrar desde (" & area.arriba & "," & area.izquierda & ") hasta (" & area.abajo & "," & area.derecha & ")")
       
    Call ME_Tools.click_tool(vbLeftButton)
    Call ME_Tools.deseleccionarTool
    
    areaSeleccionada = area_vieja
    restablecerBackupHerramientas
    

End Sub

Public Sub cortar(area As tAreaSeleccionada, elementos As Tools)
    Call Me_Tools_Seleccion.cargarAlPortapeles(areaSeleccionada)
    Call Me_Tools_Seleccion.eliminar(areaSeleccionada, elementos)
    Call Me_Tools_Seleccion.copiar(1, elementos)
End Sub

Public Function crearPresetDesdeMapa(area As tAreaSeleccionada, nombre As String, elementos As Tools) As Boolean

    Dim preset As PresetData
    Dim x As Integer
    Dim y As Integer
    Dim loopE As Byte
    
    preset.alto = area.abajo - area.arriba + 1
    preset.ancho = area.derecha - area.izquierda + 1
    
    preset.nombre = nombre
    
    ReDim preset.infoPos(1 To preset.ancho, 1 To preset.alto)
    
    For x = area.izquierda To area.derecha
        For y = area.arriba To area.abajo
        
            If elementos And Tools.tool_grh Then
                For loopE = 1 To CANTIDAD_CAPAS
                    preset.infoPos(x - area.izquierda + 1, y - area.arriba + 1).Graphic(loopE) = mapdata(x, y).Graphic(loopE)
                Next
            End If
            
            If elementos And Tools.tool_triggers Then
                preset.infoPos(x - area.izquierda + 1, y - area.arriba + 1).Trigger = mapdata(x, y).Trigger
            End If
            
            If elementos And Tools.tool_luces Then
                 preset.infoPos(x - area.izquierda + 1, y - area.arriba + 1).luz = mapdata(x, y).luz
            End If
            
            If elementos And Tools.tool_npc Then
                preset.infoPos(x - area.izquierda + 1, y - area.arriba + 1).NpcIndex = mapdata(x, y).NpcIndex
            End If
            
            If elementos And Tools.tool_obj Then
                preset.infoPos(x - area.izquierda + 1, y - area.arriba + 1).OBJInfo = mapdata(x, y).OBJInfo
            End If
            
            If elementos And Tools.tool_particles Then
                For loopE = 0 To 2
                    Set preset.infoPos(x - area.izquierda + 1, y - area.arriba + 1).Particles_groups(loopE) = mapdata(x, y).Particles_groups(loopE)
                Next
            End If
            
            If elementos And Tools.tool_tileset Then
                preset.infoPos(x - area.izquierda + 1, y - area.arriba + 1).tile_texture = mapdata(x, y).tile_texture
                preset.infoPos(x - area.izquierda + 1, y - area.arriba + 1).tile_number = mapdata(x, y).tile_number
            End If
            
        Next y
    Next x
    
    crearPresetDesdeMapa = ME_presets.agregarNuevoPreset(preset)
    
End Function
