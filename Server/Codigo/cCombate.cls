VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCombate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**
' El objetivo de esta clase es organizar un combate, llevar a los
' integrantes de los equipos al ring, prepararlos, contar
' llevando la cantidad de rounds que gano cada equipo.
' No envia mensajes a los usuarios
' No trata con la tabla del evento. No descalifica de la tabla general, solo del combate.
' No pide recursos
'@see modRings modDescansos modUsuarios modUtilitarios

Option Explicit

Public Enum eModoCombate
    alMejorDe = 1 'Al mejor de X rounds
    AlGanar = 2 'El que gane X rounds
End Enum

Public Enum eResultadoCombate 'El computarMuerte o los descalificador, devuelven un resultado. Este resultado puede ser:
    Sin_resultado = 1 'La funcion no causo ningun resultado
    Parcial = 2 'El ganador es parcial (gano o round no todo), no hay un ganador final.
    final = 3 'El ganadodr devuelto, es el que gano el combate.
End Enum

Private Enum eEstadoEquipo
    Jugando = 1
    descalificado = 2
End Enum

Private Enum eEstadoParticipante
    Jugando = 1
    descalificado = 2
End Enum

Private Type tIntegranteEquipoCombate
    IndexIntegrante As Integer 'UserIndex del integrante
    posOriginal As WorldPos 'Posicion desde donde lo traje y donde lo voy a devolver una vez terminado el reto
    IDIntegrante As Long
    Estado As eEstadoParticipante
End Type

Private Type tEquipoCombate
    id As Byte 'Indentificador unico del equipo
    cantidadIntegrantes As Byte 'Cantidad de integrantes del equipo
    integrantes() As tIntegranteEquipoCombate 'Integrantes del equipo
    cantidadRoundsGanados As Byte 'Cantidad de rounds ganados por el equipo
    cantidadIntegrantesVivos As Byte 'Cantidad de integrantes del equipo que quedan vivos
    cantidadIntegrantesDescalificados As Byte
    Estado As eEstadoEquipo 'Estado del equipo
End Type

Private Enum eEstadoCombate
    Preparando = 1
    EnCombate = 2
    Terminado = 3
End Enum

Private equipos() As tEquipoCombate 'Equipos que estan participando

Private ring As modRings.tRing 'Ring asociado al combate

Private cantidadEquipos As Byte 'Cantidad de equipos que se inscribieron en el combate
Private cantidadEquiposDescalificados As Byte 'Cantidad de equipos que fueron descalificados a lo largo del combate por alguna irregularidad
Private cantidadEquiposVivos As Byte 'Durante el round, cantidad de equipos que quedan vivos

Private modo As eModoCombate 'Modo del combate, como se declara la victoria
Private victoriasNecesarias As Byte 'Variable que, junto con modo, establece como se declara la victoria

Private Const TIEMPO_DESCANSO_PRIMER_COMBATE = 16 'segundos de descanso hasta que empieza el round
Private Const TIEMPO_DESCANSO_COMBATE = 11 'segundos de descanso hasta que empieza el round

Private cantidadRoundsJugados As Byte 'Cuenta la cantidad de rounds que se jugaron al momento

Private Sub Class_Initialize()
    'Cargo los valores predefinidos
    cantidadEquiposDescalificados = 0
    
    'Al mejor de 3
    modo = eModoCombate.alMejorDe
    victoriasNecesarias = 3
End Sub

Public Sub setCantidadEquipos(cantidad As Byte)
    cantidadEquipos = cantidad
    ReDim equipos(1 To cantidadEquipos) As tEquipoCombate
End Sub

Friend Sub setRing(ring_ As modRings.tRing)
    ring = ring_
End Sub

Public Sub configurarDecisionVictoria(modoCombate As eModoCombate, variable As Byte)
    modo = modoCombate
    victoriasNecesarias = variable
End Sub


Public Sub agregarEquipo(integrantesIndex() As Integer, integrantesid() As Long, idEquipo As Byte)
    Dim loopC As Byte

    equipos(idEquipo).id = idEquipo
    equipos(idEquipo).cantidadIntegrantes = UBound(integrantesIndex)
    equipos(idEquipo).Estado = eEstadoEquipo.Jugando
    equipos(idEquipo).cantidadIntegrantesDescalificados = 0
    
    'Obtengo memoria para guardar esta información
    ReDim equipos(idEquipo).integrantes(1 To equipos(idEquipo).cantidadIntegrantes) As tIntegranteEquipoCombate
    
    For loopC = 1 To equipos(idEquipo).cantidadIntegrantes
        equipos(idEquipo).integrantes(loopC).IndexIntegrante = integrantesIndex(loopC)
        equipos(idEquipo).integrantes(loopC).IDIntegrante = integrantesid(loopC)
        equipos(idEquipo).integrantes(loopC).Estado = eEstadoParticipante.Jugando
    Next

End Sub

'**
' Reinicia los contadores de victorias.
' Guarda la posicion original donde se encuentran los usuariros
' Pone a los usuarios en el ring
'@rem
Public Sub comenzar()
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    
    'Obtengo las posiciones donde se encuentran los usuarios antes de que los traiga al ring
    For loopEquipo = 1 To cantidadEquipos
        'Por cada integrante obtengo la posicion actual
        For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            'Guardo la posicion
            equipos(loopEquipo).integrantes(loopIntegrante).posOriginal = UserList(equipos(loopEquipo).integrantes(loopIntegrante).IndexIntegrante).pos
        Next loopIntegrante
                
    Next loopEquipo
        
    cantidadRoundsJugados = 0
    
    'Reseteo los contadores
    Call reiniciarContadoresVictorias
    
    'Le doy inicio al primer round
    Call iniciarRound(True)

End Sub

'Reinicia los contadores de victorias de los equipos que no estan descalificados
Private Sub reiniciarContadoresVictorias()
    Dim loopC As Byte
    
   
    For loopC = 1 To cantidadEquipos
        If equipos(loopC).Estado <> eEstadoEquipo.descalificado Then
            equipos(loopC).cantidadRoundsGanados = 0
        End If
    Next loopC

    
End Sub

'Reinicia los contadores de muertes
Private Sub reiniciarContadoresMuertes()
    Dim loopC As Byte
    
    cantidadEquiposVivos = cantidadEquipos - cantidadEquiposDescalificados
    
    For loopC = 1 To cantidadEquipos
        If equipos(loopC).Estado <> eEstadoEquipo.descalificado Then
            equipos(loopC).cantidadIntegrantesVivos = equipos(loopC).cantidadIntegrantes - equipos(loopC).cantidadIntegrantesDescalificados
        End If
    Next loopC

End Sub

'Continua con el siguiente round
Public Sub continuar()
    cantidadRoundsJugados = cantidadRoundsJugados + 1
    Call iniciarRound(False)
End Sub

'**
'Preparar al usuario (revivirlo, desparalizarlo, etc)
'Asignar el tiempo del contador de "timeleft" antes del comienzo del combate
'Llevarlos a las esquinas
'@rem
Private Sub iniciarRound(Optional ByVal esElPrimero As Boolean = False)
'Voy a:
'

Dim loopEquipo As Byte
Dim loopIntegrante As Byte
Dim UserIndex As Integer

Dim posicionEnElRing As Byte

'Reinicio los contadores de las muertes de los equipos
Call reiniciarContadoresMuertes

'¿Que pasa si cuando transporto hay alguien en la esquina?
'Primero recorro posicionando a los vivos. luego a los muertos
'Positivo: poco codigo,minima cantidad de warpuserchar

Dim UserEstadoProcesar As Byte

'Primero Proceso los que esten vivos y despues a los muertos
For UserEstadoProcesar = 0 To 1
    'Por cada equipo...
    For loopEquipo = 1 To cantidadEquipos
        'Solo inicia el round sino fue descalificado
        
        If equipos(loopEquipo).Estado <> eEstadoEquipo.descalificado Then
            
            posicionEnElRing = ((loopEquipo + cantidadRoundsJugados) Mod cantidadEquipos) + 1
            'Por cada integrante
            For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
    
                'Solo trato con los participantes que siguen jugando
                If equipos(loopEquipo).integrantes(loopIntegrante).Estado = eEstadoParticipante.Jugando Then
                    'Obtengo el index del integrante
                    UserIndex = equipos(loopEquipo).integrantes(loopIntegrante).IndexIntegrante
                    
                    If UserList(UserIndex).flags.Muerto = UserEstadoProcesar Then
                        'Lo preparo
                        Call modUtilitarios.Preparando(UserIndex)
            
                        'seteo el tiempo. El tiempo depende de si es el primer round o no
                        If esElPrimero Then
                            UserList(UserIndex).Counters.combateRegresiva = TIEMPO_DESCANSO_PRIMER_COMBATE
                            'Le aviso del tiempo
                            EnviarPaquete Paquetes.TiempoReto, ByteToString(TIEMPO_DESCANSO_PRIMER_COMBATE), UserIndex, ToIndex
                        Else
                            UserList(UserIndex).Counters.combateRegresiva = TIEMPO_DESCANSO_COMBATE
                            'Le aviso
                            EnviarPaquete Paquetes.TiempoReto, ByteToString(TIEMPO_DESCANSO_COMBATE), UserIndex, ToIndex
                        End If
                        
                        'Lo transporto a menos qe el usuariro ya este en esa posicion
                        If Not MapData(ring.mapa, ring.Esquina(posicionEnElRing, loopIntegrante).x, ring.Esquina(posicionEnElRing, loopIntegrante).y).UserIndex = UserIndex Then
                            'Si ya esta ahi no lo transporto
                            Call WarpUserChar(UserIndex, ring.mapa, ring.Esquina(posicionEnElRing, loopIntegrante).x, ring.Esquina(posicionEnElRing, loopIntegrante).y, False)
                        End If
                    End If
                End If
            Next loopIntegrante
        End If
    Next loopEquipo

Next UserEstadoProcesar
'Ya estan en el ring, cara a cara

End Sub

Private Function esGanadorCombate(cantidadVictorias As Byte) As Boolean

    If modo = eModoCombate.AlGanar Then 'Tiene que ganar X cantidad de rounds para ser el ganador
        If victoriasNecesarias = cantidadVictorias Then
            esGanadorCombate = True
        Else
            esGanadorCombate = False
        End If
    ElseIf modo = eModoCombate.alMejorDe Then 'Al mejor de X cantidad de rounds. Tiene que ganar X/2
        'Al mejor de 1 ... 1 , 2 ... 2 , 3 ... 2 , 4 ... 3 , 5 ... 3 , 100 ... 51
        'Division entera. La mitad + 1
        If ((victoriasNecesarias \ 2) + 1) = cantidadVictorias Then
            esGanadorCombate = True
        Else
            esGanadorCombate = False
        End If
    Else
        esGanadorCombate = True
        LogError ("Hay un combate que no tiene un modo pre definido")
    End If
        
End Function
'Verifica si un equipo gano. Si es así devuelve cual es
Public Function existeGanador(ByRef tipoResultado As eResultadoCombate) As Byte
    Dim loopEquipo As Byte
    
    'Quedo solo el ganador. Fue descalificado el resto
    'O mato a todos

    If cantidadEquiposDescalificados = cantidadEquipos - 1 Then
        'Busco al único equipo que no quedo descalificado
        For loopEquipo = 1 To cantidadEquipos
            If equipos(loopEquipo).Estado = eEstadoEquipo.Jugando Then
                existeGanador = loopEquipo
                tipoResultado = eResultadoCombate.final
                Exit Function
            End If
        Next loopEquipo
        
    ElseIf cantidadEquiposVivos = 1 Then 'Mato a todos
        'Busco cual es
        For loopEquipo = 1 To cantidadEquipos
            If equipos(loopEquipo).Estado <> eEstadoEquipo.descalificado Then
                'El que tenga al menos un integrante vivo es el ganador...
                If equipos(loopEquipo).cantidadIntegrantesVivos > 0 Then
                    'Este es el que gano!
                    equipos(loopEquipo).cantidadRoundsGanados = equipos(loopEquipo).cantidadRoundsGanados + 1
                    'Esta cantidad de victorias le sirve para ganar el combate
                    If esGanadorCombate(equipos(loopEquipo).cantidadRoundsGanados) Then
                        'Gano!
                        existeGanador = loopEquipo
                        tipoResultado = eResultadoCombate.final
                        Exit Function
                    Else
                        'Gano el round....
                        existeGanador = loopEquipo
                        '... pero no el combate
                        tipoResultado = eResultadoCombate.Parcial
                        'Listo
                        Exit Function
                    End If
                End If
            End If
        Next loopEquipo
    
    Else 'Hay otros equipos en juego
        tipoResultado = eResultadoCombate.Sin_resultado
        existeGanador = 0
    End If
   
End Function

'Devuelve el ID del equipo al cual pertenece el usuario
'**
'
'@param        IDPersonaje Required. Long.
'@return       Byte. EL ID del equipo del usuario
'@rem
Private Function obtenerEquipoUsuario(idPersonaje As Long) As Byte
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    
    For loopEquipo = 1 To cantidadEquipos
        'No tiene sentido firjarme en los equipos descalificados
        If equipos(loopEquipo).Estado <> eEstadoEquipo.descalificado Then
            'Me fijo en los integrantes
            For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
                'Es el?
                If equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante = idPersonaje Then
                    obtenerEquipoUsuario = loopEquipo
                    Exit Function
                End If
            Next loopIntegrante
        End If
    Next loopEquipo
    
    obtenerEquipoUsuario = 0

End Function

Public Function obtenerUserIndex(idPersonaje As Long) As Integer

    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte

     For loopEquipo = 1 To cantidadEquipos
        'Me fijo en los integrantes
        For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            'Es el?
            If equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante = idPersonaje Then
                obtenerUserIndex = equipos(loopEquipo).integrantes(loopIntegrante).IndexIntegrante
                Exit Function
            End If
        Next loopIntegrante
    Next loopEquipo
    obtenerUserIndex = -1
End Function

'Se computa la muerte del usuario pasado por parametro
'Puede que esto traiga como consecuencia que un equipo gane el evento
'En ese caso se devuelve el ID del equipo ganador
Public Function computarMuerte(UserIndex As Integer, ByRef tipoResultado As eResultadoCombate, ByVal personajeOffline As Boolean) As Byte
   
    Dim equipo As Byte
    Dim posicionEnElRing As Byte
     'Obtengo el equipo al cual pertenece el usuario
   
    equipo = obtenerEquipoUsuario(UserList(UserIndex).id)
    
    'Actualizo la cantidad de usuariros vivos que tiene el equipo
    equipos(equipo).cantidadIntegrantesVivos = equipos(equipo).cantidadIntegrantesVivos - 1
    
    'Era el ultimo vivo del equipo
    If equipos(equipo).cantidadIntegrantesVivos = 0 Then
        cantidadEquiposVivos = cantidadEquiposVivos - 1
    End If
    
    'Lo llevo al descanso
    'Lo mando al descanso, hasta que termine el round
    posicionEnElRing = ((equipo + cantidadRoundsJugados) Mod cantidadEquipos) + 1
    
    ' TODO. Esto hay que cambiarlo
    If personajeOffline Then
        Call modUsuarios.transportarUsuario(UserList(UserIndex).id, 0, ring.mapa, ring.descanso(posicionEnElRing, equipos(equipo).cantidadIntegrantesVivos + 1).x, ring.descanso(posicionEnElRing, equipos(equipo).cantidadIntegrantesVivos + 1).y, True, (ring.tipoRing And eRingTipo.ringAcuatico))
    Else
        Call modUsuarios.transportarUsuario(UserList(UserIndex).id, UserIndex, ring.mapa, ring.descanso(posicionEnElRing, equipos(equipo).cantidadIntegrantesVivos + 1).x, ring.descanso(posicionEnElRing, equipos(equipo).cantidadIntegrantesVivos + 1).y, True, (ring.tipoRing And eRingTipo.ringAcuatico))
    End If
    
    'Queda más de un equipo vivo, esto quiere decir que hay usuarios luchando aun
    If cantidadEquiposVivos > 1 Then
        computarMuerte = 0
        tipoResultado = eResultadoCombate.Sin_resultado
    Else 'Quedo un solo equipo vivo.
        'Veo si este equipo gano
        computarMuerte = existeGanador(tipoResultado)
    End If
   
End Function

Public Sub reviviruUsuario(UserIndex As Integer)
   
    Dim equipo As Byte
    Dim auxUserIndex As Integer
    Dim posRestuticion As Position
    
     'Obtengo el equipo al cual pertenece el usuario
    equipo = obtenerEquipoUsuario(UserList(UserIndex).id)
    
    'Todo el equipo estaba muerto?
    If equipos(equipo).cantidadIntegrantesVivos = 0 Then
        cantidadEquiposVivos = cantidadEquiposVivos + 1
    End If
    
    'Actualizo la cantidad de usuariros vivos que tiene el equipo
    equipos(equipo).cantidadIntegrantesVivos = equipos(equipo).cantidadIntegrantesVivos + 1
    
    'Lo meto en alguna parte valida del ring
    auxUserIndex = obtenerUserIndex(UserList(UserIndex).id)
    
    posRestuticion.x = RandomNumber(ring.superiorIzquierdo.x, ring.inferiorDerecho.x)
    posRestuticion.y = RandomNumber(ring.superiorIzquierdo.y, ring.inferiorDerecho.y)
    
    Call modUsuarios.transportarUsuario(UserList(UserIndex).id, UserIndex, ring.mapa, posRestuticion.x, posRestuticion.y, False, ring.tipoRing And eRingTipo.ringAcuatico)
    
End Sub


Public Sub actualizarUserIndexPersonajeID(idPersonaje As Long, NuevoUserIndex As Integer)

    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte

     For loopEquipo = 1 To cantidadEquipos
        'No tiene sentido firjarme en los equipos descalificados
        If equipos(loopEquipo).Estado <> eEstadoEquipo.descalificado Then
            'Me fijo en los integrantes
            For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
                'Es el?
                If equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante = idPersonaje Then
                    equipos(loopEquipo).integrantes(loopIntegrante).IndexIntegrante = NuevoUserIndex
                    Exit Sub
                End If
            Next loopIntegrante
        End If
    Next loopEquipo

End Sub
'Se descalifica al equipo de este individuo

Public Sub descalificarUsuario2(idPersonaje As Long)
    
    Dim equipo As Byte
    
    equipo = obtenerEquipoUsuario(idPersonaje)
    
    Dim loopIntegrante
    
    For loopIntegrante = 1 To equipos(equipo).cantidadIntegrantes
        If equipos(equipo).integrantes(loopIntegrante).IDIntegrante = idPersonaje _
            And equipos(equipo).integrantes(loopIntegrante).Estado = eEstadoParticipante.Jugando Then
            
            equipos(equipo).integrantes(loopIntegrante).Estado = eEstadoParticipante.descalificado
            equipos(equipo).cantidadIntegrantesDescalificados = equipos(equipo).cantidadIntegrantesDescalificados + 1

            Call deshacerIntegrante(equipos(equipo).integrantes(loopIntegrante))
        End If
    
    Next loopIntegrante
End Sub
Public Sub descalificarEquipoDelUsuario2(idPersonaje As Long)
    Dim equipo As Byte
    'Obtengo el equipo el usuario
    equipo = obtenerEquipoUsuario(idPersonaje)
    
    'Lo descalifico
    equipos(equipo).Estado = eEstadoEquipo.descalificado
    
    'Aumento la cantidad de equipos descalificados
    cantidadEquiposDescalificados = cantidadEquiposDescalificados + 1
    
    'Me deshago del equipo
    Call deshacerEquipo(equipo)
End Sub

'Se descalifica al equipo de este individuo
Public Function descalificarEquipoDelUsuario(UserIndex As Integer, ByRef tipoResultado As eResultadoCombate) As Byte
    Dim equipo As Byte
    'Obtengo el equipo el usuario
    equipo = obtenerEquipoUsuario(UserList(UserIndex).id)
    
    'Me fijo el estado del equipo.
    'Si el equipo aun no habia perdido el round, (1 o m?s integrantes vivos)
    'Ahora se consdiera que perdio
    If equipos(equipo).cantidadIntegrantesVivos > 0 Then
        cantidadEquiposVivos = cantidadEquiposVivos - 1
    End If
    
    'Lo descalifico
    equipos(equipo).Estado = eEstadoEquipo.descalificado
    
    'Aumento la cantidad de equipos descalificados
    cantidadEquiposDescalificados = cantidadEquiposDescalificados + 1
    
    'Me deshago del equipo
    Call deshacerEquipo(equipo)
    
    'Que haya un equipo menos causo que haya un ganador?
    descalificarEquipoDelUsuario = existeGanador(tipoResultado)

End Function

Public Sub cancelar()
    Call terminar
End Sub
Public Sub terminar()
    Dim loopEquipo As Byte
      
    'Me deshago de todos los equipo que aun estaban participando
    For loopEquipo = 1 To cantidadEquipos
        If equipos(loopEquipo).Estado <> eEstadoEquipo.descalificado Then
            Call deshacerEquipo(loopEquipo)
        End If
    Next loopEquipo

End Sub

Private Sub deshacerIntegrante(Integrante As tIntegranteEquipoCombate)
    'Lo transporto a donde estaban antes del combate
    If Integrante.IndexIntegrante > 0 Then
        Call modUtilitarios.Preparando(Integrante.IndexIntegrante)
    End If

    Call modUsuarios.transportarUsuario(Integrante.IDIntegrante, Integrante.IndexIntegrante, Integrante.posOriginal.map, Integrante.posOriginal.x, Integrante.posOriginal.y, False, False)
End Sub


Private Sub deshacerEquipo(idEquipo As Byte)
    Dim loopIntegrante As Byte
    Dim Integrante As tIntegranteEquipoCombate

    For loopIntegrante = 1 To equipos(idEquipo).cantidadIntegrantes
        Integrante = equipos(idEquipo).integrantes(loopIntegrante)

        If Integrante.Estado = eEstadoParticipante.Jugando Then
            'Lo transporto a donde estaban antes del combate
            Call deshacerIntegrante(Integrante)
        End If
    Next loopIntegrante

End Sub
