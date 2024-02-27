Attribute VB_Name = "SV_Mundo"
'/* CARGA DE NUEVOS MAPAS
Option Explicit

Public NumMaps As Integer

Private Function crearAccion(archivoOrigen As Integer) As iAccion
    Dim tempID As Integer
    
    Get archivoOrigen, , tempID
    
    Set crearAccion = Sv_Acciones.obtenerAccion(tempID)
    
    Call crearAccion.cargar(archivoOrigen)

End Function

Private Sub cargarListaAcciones(archivoOrigen As Integer, ByRef lista() As iAccion)
    Dim cantidadAcciones As Integer
    Dim loopAccion As Integer
    
    Get archivoOrigen, , cantidadAcciones
    
    If cantidadAcciones = 0 Then
    ReDim lista(0)
    Else
    ReDim lista(1 To cantidadAcciones) As iAccion
    
    For loopAccion = 1 To cantidadAcciones
            Set lista(loopAccion) = crearAccion(archivoOrigen)
    Next loopAccion
    End If
End Sub

' Esta funcion carga la configuracion de cada uno de los mapas.
' Tambén va a cargar la cantidad de mapas que hay en el mundo
Private Function cargarConfiguraciones() As Boolean

    Dim cantidad As Integer
    Dim datos As New cIniManager
    Dim loopMapa As Integer
    
    ' Abrimos el archivo
    Call datos.Initialize(MapPath & "mapas.dat")

    ' Obtenemos la cantidad de mapas
    cantidad = CInt(val(datos.getNameLastSection()))

    If cantidad <= 0 Then
        Call LogError("Cantidad de mapas menor o igual a 0.")
        Exit Function
    End If
    
    NumMaps = cantidad
    
    ' Preparamos para guardar la info
    ReDim MapInfo(1 To NumMaps) As MapInfo
     
    ' Cargamos
    For loopMapa = 1 To NumMaps
    
        With MapInfo(loopMapa)
        
            .Name = datos.getValue(loopMapa, "NOMBRE")
            .UsuariosMaximo = val(datos.getValue(loopMapa, "MAXPERSONAJES"))
            
            If .UsuariosMaximo = 0 Then
                .UsuariosMaximo = 999
            End If
            
            .zona = val(datos.getValue(loopMapa, "ZONA"))
            
            .MaxLevel = val(datos.getValue(loopMapa, "NIVELMAXIMO"))
            
            If .MaxLevel = 0 Then
                .MaxLevel = STAT_MAXELV
            End If
     
            .Nivel = val(datos.getValue(loopMapa, "NIVELMINIMO"))
            
            .Music = val(datos.getValue(loopMapa, "MUSICA"))
            
            .Pk = val(datos.getValue(loopMapa, "ZONASEGURA")) = 0

            .SeCaeiItems = val(datos.getValue(loopMapa, "CAENITEMS"))
            
            .PermiteRoboNPC = val(datos.getValue(loopMapa, "PERMITIRROBO"))
             
            .restringir = val(datos.getValue(loopMapa, "NEWBIE"))
            
            .Frio = val(datos.getValue(loopMapa, "FRIO"))
            
            .BackUp = val(datos.getValue(loopMapa, "GUARDAR"))
            
            .Continente = val("")
            
            
            ' TODO: ACCESOSTATUS'
            Dim status As Byte
            status = val(datos.getValue(loopMapa, "GUARDAR"))
            
            .SoloCiudas = 0
            .SoloCrimis = 0
            .SoloCaos = 0
            .SoloArmada = 0
'

'        .Terreno = GetVar(MapPath & numeroMapa & ".dat", "Mapa" & numeroMapa, "Terreno")
'
'        .Aotromapa.map = val(GetVar(MapPath & numeroMapa & ".dat", "Mapa" & numeroMapa, "SeVanumeroMapa"))
'        .Aotromapa.x = val(GetVar(MapPath & numeroMapa & ".dat", "Mapa" & numeroMapa, "SeVaX"))
'        .Aotromapa.y = val(GetVar(MapPath & numeroMapa & ".dat", "Mapa" & numeroMapa, "SeVaY"))
'
         .AntiHechizosPts = val(datos.getValue(loopMapa, "ANTIPTS"))
        End With
        
    Next
    
    Set datos = Nothing
    
    cargarConfiguraciones = True
End Function

' Esta funcion carga la estructura (acciones, criaturas, bloqueos) de cada
' uno de los mapas existentes
Private Function cargarEstructuras() As Boolean
    Dim loopMapa As Integer
    Dim archivoMapa As String
    
    ' Preparamos el espacio donde lo vamos a guardar
    ReDim MapData(1 To NumMaps, X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE, Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE) As MapBlock
    
    ' Cargamos mapa por mapa
    For loopMapa = 1 To NumMaps
    
        archivoMapa = MapPath & loopMapa & ".servermap"

        If FileExist(archivoMapa) Then
            Call SV_Mundo.cargarMapa(archivoMapa, CInt(loopMapa))
            MapInfo(loopMapa).Existe = True
        Else
           ' LogError "Se intenta abrir el archivo de mapa " & archivoMapa & " y no existe."
            MapInfo(loopMapa).Existe = False
        End If
        
    Next
    
    cargarEstructuras = True
End Function

Public Sub cargarMundo()
    
    Call cargarConfiguraciones
    
    Call cargarEstructuras
    
End Sub

Public Sub cargarMapa(ruta As String, numeroMapa As Integer)

Dim handle As Integer
Dim TempInt As Integer
Dim tempLong As Long
Dim y As Long
Dim x As Long
Dim byflags As Integer

Dim listaAcciones() As iAccion

Dim nombre As String * 32
Dim numero As Integer

handle = FreeFile

Open ruta For Binary Access Read As handle

Seek handle, 1
    
    Get handle, , nombre
    Get handle, , numero
    
    'inf Header
    Get handle, , tempLong 'TODO Se supone que siempre va a ser cuadrado..
    Get handle, , TempInt
    Get handle, , TempInt
    Get handle, , TempInt
    Get handle, , TempInt
     
    
    'Sistema de acciones
    Call cargarListaAcciones(handle, listaAcciones)

    'Creo el listado de usuarios que estan en este mapa
    Set MapInfo(numeroMapa).usuarios = New EstructurasLib.ColaConBloques
    'Creo el listado de los npcs que tiene el mapa
    Set MapInfo(numeroMapa).NPCs = New EstructurasLib.ColaConBloques

    Set MapInfo(numeroMapa).fogatas = New Collection
    
    'Leo la información de los tiles
    For y = Y_MINIMO_VISIBLE To Y_MAXIMO_VISIBLE
        For x = X_MINIMO_VISIBLE To X_MAXIMO_VISIBLE
        
            'Leo la bandera que me dice que propiedades tiene el tile
            Get handle, , byflags

            '1) ¿Tiene un evento este tile?
            If byflags And 1 Then
                Get handle, , TempInt 'Tomo el ID de la accion relativa
                'Busco esta ID
                'se la asigno
                '3b)
                Set MapData(numeroMapa, x, y).accion = listaAcciones(TempInt)
            End If
                
            '2) ¿Hay un NPC?
            If byflags And 2 Then
                Get handle, , MapData(numeroMapa, x, y).npcIndex
                
                ' Si el npc debe hacer respawn en la pos
                'original la guardamos
                MapData(numeroMapa, x, y).npcIndex = OpenNPC(MapData(numeroMapa, x, y).npcIndex)

                NpcList(MapData(numeroMapa, x, y).npcIndex).pos.map = numeroMapa
                NpcList(MapData(numeroMapa, x, y).npcIndex).pos.x = x
                NpcList(MapData(numeroMapa, x, y).npcIndex).pos.y = y
                Call MakeNPCChar(ToNone, 0, 0, MapData(numeroMapa, x, y).npcIndex, numeroMapa, x, y)
 
            End If
            
            '3) Hay un objeto?
            If byflags And 4 Then
                '5b)
                Get handle, , MapData(numeroMapa, x, y).OBJInfo.ObjIndex
                Debug.Print ObjData(MapData(numeroMapa, x, y).OBJInfo.ObjIndex).Name
                Get handle, , MapData(numeroMapa, x, y).OBJInfo.Amount
            End If
            
            '4)
            If byflags And 8 Then
                Get handle, , MapData(numeroMapa, x, y).Trigger
                'Debug.Print MapData(numeroMapa, x, y).Trigger
            End If
            
        Next x
    Next y
    
    Close handle
End Sub

