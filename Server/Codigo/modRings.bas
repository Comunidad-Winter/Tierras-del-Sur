Attribute VB_Name = "modRings"
Option Explicit

Public Enum eRingTipo
    ringCualquiera = 1
    ringReto = 2
    ringTorneo = 4
    ringPlantado = 8
    ringAcuatico = 16
End Enum

Public Type tRing
    id As Byte 'ID del ring
    
    tipoRing As eRingTipo
    
    cantidadEquipos As Byte 'Cantidad de equipos que soporta
    capacidadPorEquipo As Byte 'Cantidad de integrantes que soporta por equipo
    
    mapa As Integer 'Mapa donde se encuentra el ring
    Esquina() As Position 'Posiciones donde estan las distintas esquinas (Equipo, Integrante)
    descanso() As Position 'Posiciones donde va el personaje mientras esta muerto esperando que el round termine
    
    superiorIzquierdo As Position 'Parte superior izquierda del ring. Variable para saber donde limpiar
    inferiorDerecho As Position 'Variable para saber donde limpiar
End Type


Private isRingDisponible() As Byte 'Establece si el ring esta disponible o no. 0: disponible. 1: No disponible
Private cantidadRingsDisponibles As Integer 'Cantidad de rings que hay disponibles
Private rings() As tRing 'Lista de rings


Public Sub reCargarRings()
    'La re-carga no pone a los rings como disponibles
    'Esto es para no hacer caos en caso de que haya rings que se estan utilizado
    'Si marco a todos como liberados puede que otro evento pueda pedir un ring
    'y que se le de uno que se esta utilizando.
    Dim cantidadActuales As Byte
    
    cantidadActuales = UBound(isRingDisponible)
    
    Call cargarRings
    
    'Se agregaron nuevos rings?
    If cantidadActuales < UBound(isRingDisponible) Then
        cantidadRingsDisponibles = cantidadRingsDisponibles + (UBound(isRingDisponible) - cantidadActuales)
    End If
    
End Sub

Public Sub iniciarRings()
    Dim loopRing As Integer
    Dim cantidadRings As Integer
    
    'Los cargo
    Call cargarRings
    
    'Los marco a todos como disponibles
    cantidadRings = UBound(isRingDisponible)
    
    For loopRing = 1 To cantidadRings
        isRingDisponible(loopRing) = 0
    Next loopRing
    
    'Todos estan disponibles
    cantidadRingsDisponibles = cantidadRings
End Sub


Private Sub cargarRings()

    Dim ring As tRing
    
    Dim cantidadRings As Integer
    Dim cantidadRingsDistintos As Integer
    Dim ruta As String
    
    Dim loopRings As Integer
    Dim loopEquipo As Integer
    Dim loopEsquina As Integer
    
    Dim infoEsquina As String
    Dim infoDescanso As String
    
    Dim comienzoInfoCoordenada As Integer
    Dim comienzoInfoCoordenadaDescanso As Integer
    
    ruta = App.Path & "\Dat\Rings.dat"
    
    cantidadRings = val(GetVar(ruta, "INIT", "NumRings"))
    cantidadRingsDistintos = val(GetVar(ruta, "INIT", "NumRingsDistintos"))
    
    ReDim Preserve rings(1 To cantidadRings) As tRing
    ReDim Preserve isRingDisponible(1 To cantidadRingsDistintos) As Byte
    
    For loopRings = 1 To cantidadRings
        
        ring.id = val(GetVar(ruta, "RING" & loopRings, "ID"))
        
        ring.tipoRing = val(val(GetVar(ruta, "RING" & loopRings, "Tipo")))
        
        If (ring.tipoRing = 0) Then
            Call LogError("No tiene tipo el ring " & loopRings)
            Exit Sub
        End If
        
        ring.mapa = val(GetVar(ruta, "RING" & loopRings, "Mapa"))
        
        ring.cantidadEquipos = val(GetVar(ruta, "RING" & loopRings, "CantidadEquipos"))
        ring.capacidadPorEquipo = val(GetVar(ruta, "RING" & loopRings, "CapacidadPorEquipo"))
        
        'Redimensiono los arrays donde voy a guardar la info
        ReDim ring.Esquina(1 To ring.cantidadEquipos, 1 To ring.capacidadPorEquipo) As Position
        ReDim ring.descanso(1 To ring.cantidadEquipos, 1 To ring.capacidadPorEquipo) As Position
    
        'Cargo la parte superior
        infoEsquina = Trim(GetVar(ruta, "RING" & loopRings, "SuperiorIzquierda"))
        
        ring.superiorIzquierdo.x = obtenerCoordenadas(infoEsquina)(0)
        ring.superiorIzquierdo.y = obtenerCoordenadas(infoEsquina)(1)
        
        If ring.superiorIzquierdo.x = 0 Or ring.superiorIzquierdo.y = 0 Then
            Call LogError("No se pudo cargar la parte superior izquierda del ring " & loopRings)
            Exit Sub
        End If
        
        'Cargo la parte inferior
        infoEsquina = Trim(GetVar(ruta, "RING" & loopRings, "InferiorDerecha"))
        comienzoInfoCoordenada = InStr(1, infoEsquina, "(")
        
        ring.inferiorDerecho.x = obtenerCoordenadas(infoEsquina)(0)
        ring.inferiorDerecho.y = obtenerCoordenadas(infoEsquina)(1)
        
        If ring.inferiorDerecho.x = 0 Or ring.inferiorDerecho.y = 0 Then
            Call LogError("No se pudo cargar la parte inferior derecha del ring " & loopRings)
            Exit Sub
        End If
        
             
        For loopEquipo = 1 To ring.cantidadEquipos
            
            'Cargo la info de la esquina y de su correspondiente descanso
            infoEsquina = GetVar(ruta, "RING" & loopRings, "Esquina" & loopEquipo)
            infoDescanso = Trim(GetVar(ruta, "RING" & loopRings, "Descanso" & loopEquipo))
        
            loopEsquina = 1
            comienzoInfoCoordenada = 1
            comienzoInfoCoordenadaDescanso = 1
            
            If Len(infoEsquina) > 0 Then
            
                Do While loopEsquina <= ring.capacidadPorEquipo
                    
                    comienzoInfoCoordenada = InStr(comienzoInfoCoordenada, infoEsquina, "(")
                
                    If comienzoInfoCoordenada > 0 Then
                        
                        Dim coordenada As String
                        
                        coordenada = mid$(infoEsquina, comienzoInfoCoordenada, InStr(comienzoInfoCoordenada, infoEsquina, ")") - comienzoInfoCoordenada + 1)
                        
                        ring.Esquina(loopEquipo, loopEsquina).x = obtenerCoordenadas(coordenada)(0)
                        ring.Esquina(loopEquipo, loopEsquina).y = obtenerCoordenadas(coordenada)(1)
                    
                        If ring.Esquina(loopEquipo, loopEsquina).x > 0 And ring.Esquina(loopEquipo, loopEsquina).y > 0 Then
                            'Chequeo que este el parentesis que cierra

                            'Todo ok en la carga de la posicion de la esquina
                            'Ahora cargo la posicion del dencaso correspondiete
                            'Cargo el correspondiente descanso
                            'Esta la informacion?
                            If Len(infoDescanso) > 0 Then
                                'Donde comienza la info de la coordenada?
                                comienzoInfoCoordenadaDescanso = InStr(comienzoInfoCoordenadaDescanso, infoDescanso, "(")
                            
                                'Esta el parentesis que abre?
                                If comienzoInfoCoordenadaDescanso > 0 Then
                                    Dim coordenadaDescanso As String
                                    
                                    coordenadaDescanso = mid$(infoDescanso, comienzoInfoCoordenadaDescanso, InStr(comienzoInfoCoordenadaDescanso, infoDescanso, ")") - comienzoInfoCoordenadaDescanso + 1)
                        
                                    ring.descanso(loopEquipo, loopEsquina).x = obtenerCoordenadas(coordenadaDescanso)(0)
                                    ring.descanso(loopEquipo, loopEsquina).y = obtenerCoordenadas(coordenadaDescanso)(1)
                                
                                    If ring.descanso(loopEquipo, loopEsquina).x > 0 And ring.descanso(loopEquipo, loopEsquina).y > 0 Then
   
                                          'Esta todo ok
                                        comienzoInfoCoordenada = InStr(comienzoInfoCoordenada, infoEsquina, ")") + 1
                                        comienzoInfoCoordenadaDescanso = InStr(comienzoInfoCoordenadaDescanso, infoDescanso, ")") + 1
                                        loopEsquina = loopEsquina + 1
                                    Else
                                        Call LogError("No se pudo cargar el descanso " & loopEsquina & " del ring " & loopRings & ". Las coordenadas no se cargaron correctamente.")
                                        Exit Sub
                                    End If
                                Else
                                    Call LogError("No se pudo cargar el descanso " & loopEsquina & " del ring " & loopRings & ". No se encontro la apertura de la coordenada.")
                                    Exit Sub
                                End If
                            Else
                                Call LogError("No se pudo cargar el descanso " & loopEsquina & " del ring " & loopRings & ". No se encontro la información del descanso.")
                                Exit Sub
                            End If
        
                        Else
                            Call LogError("No se pudo cargar la esquina " & loopEsquina & " del ring " & loopRings & ". Las coordenadas no se cargaron correctamente.")
                            Exit Sub
                        End If
                        
                    Else
                        Call LogError("No se pudo cargar la esquina " & loopEsquina & " del ring " & loopRings & ". No se encontro la apertura de la coordenada.")
                        Exit Sub
                    End If
                    
                Loop
            Else
                Call LogError("No se pudo cargar la esquina " & loopEsquina & " del ring " & loopRings)
                Exit Sub
            End If
        
        Next loopEquipo
    
        'Termino la carga
        'lo agrego
        rings(loopRings) = ring
    Next loopRings
    
End Sub


Public Function obtenerRing(cantidadEquipos As Byte, capacidadPorEquipo As Byte, tipoRing As eRingTipo) As tRing
    
    Dim encontrado As Boolean
    Dim loopC As Integer
       
    If cantidadRingsDisponibles > 0 Then
        
        loopC = 0
        encontrado = False
        
        'Busco el ring que necesito
        Do While Not encontrado And loopC < UBound(rings)
            'Siguiente
            loopC = loopC + 1
            'Esta disponible
            If isRingDisponible(rings(loopC).id) = 0 Then
                'El ring me sirve?. ¿Es del tipo que necesito?
                If (rings(loopC).tipoRing = tipoRing) Then
                    ' ¿Tiene el espacio suficiente?
                    If rings(loopC).cantidadEquipos >= cantidadEquipos And rings(loopC).capacidadPorEquipo >= capacidadPorEquipo Then
                        encontrado = True
                    End If
                End If
            End If
        Loop
    
        If encontrado Then
            'Marco al ring como no disponible
            isRingDisponible(rings(loopC).id) = 1
            cantidadRingsDisponibles = cantidadRingsDisponibles - 1
            obtenerRing = rings(loopC)
        Else
            'No hay ring disponible
            obtenerRing.id = 0
        End If
    Else
        'No hay ring disponible
        obtenerRing.id = 0
    End If

End Function

Public Sub liberarRing(ByRef ring As tRing)
    'Lo limpio

    Call limpiarRing(ring)
    
    'Lo marco como disponible
    isRingDisponible(ring.id) = 0
    'Aumento la cantidad de rings disponibles
    cantidadRingsDisponibles = cantidadRingsDisponibles + 1

    Call LogNuevosRetos("Libero el ring " & ring.id)
    ring.id = 0

End Sub

'Limpio el ring de los objetos que pudieron llegar a tirar los participantes
Private Sub limpiarRing(ByRef ring As tRing)
    Dim mapa As Integer
    Dim y As Byte
    Dim x As Byte
    
    mapa = ring.mapa
   
    For y = ring.superiorIzquierdo.y To ring.inferiorDerecho.y
        For x = ring.superiorIzquierdo.x To ring.inferiorDerecho.x
            If MapData(mapa, x, y).OBJInfo.ObjIndex > 0 Then
                If ItemNoEsDeMapa(MapData(mapa, x, y).OBJInfo.ObjIndex) Then
                    Call EraseObj(ToMap, 0, mapa, 10000, mapa, x, y)
                End If
            End If
        Next x
    Next y


End Sub

'FUNCIONES DE DEBUG

Public Sub verEstado(ByRef lista As VB.ListBox)
    Dim loopRing As Byte
    Dim cantidadDisponibles As Byte
    Dim loopRing2 As Byte
    Dim linea1 As String
    
    cantidadDisponibles = 0
    
    For loopRing = 1 To UBound(isRingDisponible)
        If isRingDisponible(loopRing) = 0 Then
            Call lista.AddItem(loopRing & ": Disponible")
            cantidadDisponibles = cantidadDisponibles + 1
        Else
            Call lista.AddItem(loopRing & ": NO disponible ")
        End If
        
        Call LogDesarrollo("Ring " & loopRing)
        
        For loopRing2 = 1 To UBound(rings)
            If rings(loopRing2).id = loopRing Then
                linea1 = "  " & getDescTipoRing(rings(loopRing2).tipoRing) & "-> " & rings(loopRing2).cantidadEquipos & "x" & rings(loopRing2).capacidadPorEquipo & ". Mapa: " & rings(loopRing2).mapa
                Call lista.AddItem(linea1)
                Call LogDesarrollo(linea1)
            End If
        Next
    Next loopRing
    
    Call lista.AddItem("Disponibles: " & cantidadDisponibles)
End Sub

Public Function getCantidadRingsDisponibles() As Integer
    getCantidadRingsDisponibles = cantidadRingsDisponibles
End Function

Public Function getDescTipoRing(tipoRing As eRingTipo) As String
    Dim temp As String
    
    temp = ""
    
    If (tipoRing And eRingTipo.ringReto) Then temp = temp & " reto"
    
    If (tipoRing And eRingTipo.ringTorneo) Then temp = temp & " torneo"
    
    If (tipoRing And eRingTipo.ringPlantado) Then temp = temp & " plante"
        
    If (tipoRing And eRingTipo.ringAcuatico) Then temp = temp & " acuatico"
    
    getDescTipoRing = temp
    
End Function

