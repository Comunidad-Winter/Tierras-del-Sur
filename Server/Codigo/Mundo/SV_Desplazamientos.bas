Attribute VB_Name = "SV_Desplazamientos"
Option Explicit

Public Function personajePuedeMoverse(Usuario As User) As Boolean
    
    If Usuario.Counters.combateRegresiva > 0 Then 'No se puede mover si esta en cuenta regresiva
        personajePuedeMoverse = False
        Exit Function
    End If
        
    'No se puede moversi esta paralizado o inmovilizado
    If Usuario.flags.Paralizado = 1 Or Usuario.flags.Inmovilizado = 1 Then
        personajePuedeMoverse = False
        Exit Function
    End If

    personajePuedeMoverse = True
End Function


'* Devuelve True si puedo trasportar al usuario.
'False no pudo y no le hizo nada al pj
Public Function avanzarPersonajeOtroMapa(Usuario As User, mapaDestino As Integer, _
                                                            xDestino As Byte, _
                                                            yDestino As Byte) As Boolean

    'El mapa existe? esta cargado? la posicion es correcta?
    If SV_PosicionesValidas.existePosicionMundo(mapaDestino, xDestino, yDestino) Then
        If SV_PosicionesValidas.personajePuedeIngresarMapa(Usuario, MapInfo(mapaDestino)) Then
            
            'Parecido al mover
            'Bloqueado por donde quiero ir?
            If SV_Bloqueos.sePuedeIngresarTile(MapData(mapaDestino, xDestino, yDestino).Trigger, Usuario.Char.heading) Then

                If SV_PosicionesValidas.esPosicionUsablePersonaje(Usuario, MapData(mapaDestino, xDestino, yDestino)) Then
                    
                    If MapData(mapaDestino, xDestino, yDestino).UserIndex = 0 _
                        And MapData(mapaDestino, xDestino, yDestino).npcIndex = 0 Then
                    
                        'Cambiar
                        Call WarpUserChar(Usuario.UserIndex, mapaDestino, xDestino, yDestino)
                    
                        'Ejecuto la accion que aqui hay
                        If Not MapData(Usuario.pos.map, xDestino, yDestino).accion Is Nothing Then
                            Call MapData(Usuario.pos.map, xDestino, yDestino).accion.ejecutar(Usuario.pos.map, CByte(Usuario.pos.x), CByte(Usuario.pos.y))
                        End If
                        
                        avanzarPersonajeOtroMapa = True
                    Else
                        Dim auxPos As WorldPos
                        Dim nPos As WorldPos
                        
                        auxPos.map = mapaDestino
                        auxPos.x = xDestino
                        auxPos.y = yDestino
                        
                        Call ClosestLegalPos(auxPos, nPos, Usuario)
                        
                        If nPos.x <> 0 And nPos.y <> 0 Then
                            Call WarpUserChar(Usuario.UserIndex, nPos.map, nPos.x, nPos.y)
                            avanzarPersonajeOtroMapa = True
                        Else
                            avanzarPersonajeOtroMapa = False
                        End If

                    End If
                End If
            End If
        End If
    End If

    'Si llegue hasta acá es porque no puede ingresar al mapa
    avanzarPersonajeOtroMapa = False
    
End Function




Public Function moverPersonajeHacia(Usuario As User, ByVal NuevoX As Integer, ByVal NuevoY As Integer, nuevoHeading As eHeading)

Dim avanzo As Boolean
Dim otroUser As Integer

avanzo = True
'El personaje en si mismo puede moverse?
If personajePuedeMoverse(Usuario) Then
    'Donde se queire mover es una posicion que existe?
    If SV_PosicionesValidas.esPosicionJugable(NuevoX, NuevoY) Then
        'Puedo entrar por el lado que quiero a la tile?
        If SV_Bloqueos.sePuedeIngresarTile(MapData(Usuario.pos.map, NuevoX, NuevoY).Trigger, nuevoHeading) Then
            'Puede el usuario estar en esa posicion?
            If SV_PosicionesValidas.esPosicionUsablePersonaje(Usuario, MapData(Usuario.pos.map, NuevoX, NuevoY)) Then
                'No hay un npc
                If MapData(Usuario.pos.map, NuevoX, NuevoY).npcIndex = 0 Then
                    
                    otroUser = MapData(Usuario.pos.map, NuevoX, NuevoY).UserIndex
                    
                    If otroUser > 0 Then
                    
                        'Bajo ciertas circunstancias puede pasar
                        avanzo = False
                        
                        If Usuario.flags.Muerto = 0 And Usuario.flags.Navegando = 0 Then
                            If UserList(otroUser).flags.Muerto = 1 Then
                                avanzo = True
                                Call intercambiarPersonajesAdyacantes(Usuario, UserList(otroUser), nuevoHeading)
                            End If
                        End If
                        
                    Else 'NO hay nadie. Todo bien
                         
                        EnviarPaquete Paquetes.MoveChar, ITS(Usuario.Char.charIndex) & ITS(NuevoX) & ITS(NuevoY), Usuario.UserIndex, ToAreaButIndex
        
                        MapData(Usuario.pos.map, Usuario.pos.x, Usuario.pos.y).UserIndex = 0
        
                        Usuario.pos.x = NuevoX
                        Usuario.pos.y = NuevoY
        
                        Usuario.Char.heading = nuevoHeading
        
                        MapData(Usuario.pos.map, NuevoX, NuevoY).UserIndex = Usuario.UserIndex
        
                        Call ActualizarArea(Usuario.UserIndex, nuevoHeading)
                    End If
                    
                    'La acción es valida?
                    If avanzo Then
                        If Not MapData(Usuario.pos.map, NuevoX, NuevoY).accion Is Nothing Then
                            Call MapData(Usuario.pos.map, NuevoX, NuevoY).accion.ejecutar(Usuario.pos.map, CByte(Usuario.pos.x), CByte(Usuario.pos.y))
                        End If
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
End If
'Si llegamos aca
'No se pudo mover, le actualizo la posicion por si se movio en el cliente

Call enviarPosicion(Usuario)

End Function

'Si o si lo transporta
Public Sub transportarUsuario(Usuario As User, mapa As Integer, xDestino As Byte, yDestino As Byte)

    'Esta no puede fallar
    If Not trasportarUsuarioOnline(Usuario, mapa, xDestino, yDestino, 0) Then
        Call transportarUsuarioOffline(Usuario.id, mapa, xDestino, yDestino)
    End If

End Sub

Public Function trasportarUsuarioOnline(Usuario As User, ByVal mapa As Integer, ByVal xDestino As Byte, ByVal yDestino As Byte, ByVal radio As Byte) As Boolean

Dim x As Byte
Dim y As Byte
Dim intentos As Byte

If Not SV_PosicionesValidas.existeMapa(mapa) Then
    trasportarUsuarioOnline = False
    Exit Function
End If

If radio > 0 Then
    xDestino = xDestino - radio + Int(Rnd * (radio * 2 + 1))
    yDestino = yDestino - radio + Int(Rnd * (radio * 2 + 1))
End If

intentos = 0

Do While intentos < 2

    For x = xDestino - intentos To xDestino + intentos
        For y = yDestino - intentos To yDestino + intentos
            
            If SV_PosicionesValidas.esPosicionJugable(x, y) Then
                If SV_PosicionesValidas.esPosicionUsablePersonaje(Usuario, MapData(mapa, x, y)) Then
                    If MapData(mapa, x, y).UserIndex = 0 And MapData(mapa, x, y).npcIndex = 0 Then
                        'Lo muevo
                        Call WarpUserChar(Usuario.UserIndex, mapa, x, y)
                                
                        trasportarUsuarioOnline = True
                        Exit Function
                    End If
                End If
            End If
        Next y
    Next x

intentos = intentos + 1
Loop

trasportarUsuarioOnline = False

End Function
'/*
Sub intercambiarPersonajesAdyacantes(personaje1 As User, personaje2 As User, headingPJ1 As eHeading)

   On Error GoTo intercambiarPersonajesAdyacantes

Dim auxPos As Position
Dim ladoContrario As eHeading

ladoContrario = SV_Bloqueos.obtenerBordeContrario(headingPJ1)
'Pongo al  1 donde esta el 2
auxPos.x = personaje2.pos.x
auxPos.y = personaje2.pos.y

personaje2.pos.x = personaje1.pos.x
personaje2.pos.y = personaje1.pos.y

'Pongo al pj 2 donde esta el 1
personaje1.pos.x = auxPos.x
personaje1.pos.y = auxPos.y

MapData(personaje1.pos.map, personaje1.pos.x, personaje1.pos.y).UserIndex = personaje1.UserIndex
MapData(personaje2.pos.map, personaje2.pos.x, personaje2.pos.y).UserIndex = personaje2.UserIndex

'EnviarPaquete Paquetes.MoveChar, ITS(Usuario.Char.charIndex) & Chr$(NuevoX) & Chr$(NuevoY), Usuario.UserIndex, ToAreaButIndex
    
    
EnviarPaquete Paquetes.MoverMuerto, ITS(personaje1.Char.charIndex) & headingPJ1, 0, ToMap, personaje1.pos.map
EnviarPaquete Paquetes.MoverMuerto, ITS(personaje2.Char.charIndex) & ladoContrario, 0, ToMap, personaje2.pos.map

Call ActualizarArea(personaje1.UserIndex, headingPJ1)
Call ActualizarArea(personaje2.UserIndex, ladoContrario)
Exit Sub

intercambiarPersonajesAdyacantes:

     LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WarpUserCharEspecial of Módulo UsUaRiOs"
End Sub
