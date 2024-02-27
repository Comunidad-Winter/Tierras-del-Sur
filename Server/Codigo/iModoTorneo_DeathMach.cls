VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iModoTorneo_DeathMach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**
' Verifica que los equipos esten online. Maneja las advertencias
' Verifica que el equipo tenga el minimo de integrantes necesarios para poder jugar
' Procesa los abandonos
' Descalifica usuariros
' Envia mensajes a los equipos
'@see cCombate modRings modTorneos
Option Explicit

Implements iModoTorneo

Private tablaEquipos() As tEquipoTablaTorneo 'Tabla donde esta la informacion de los equipos
Private puntosEquipo() As Byte 'Cantidad de rounds ganados por el equipo

Private nombreEvento As String 'Nombre del "sub" evento

Private Estado As eEstadoEvento 'Estado del sub evento

Private tiempoTranscurrido As Integer 'Minutos

Private cantidadEquipos As Byte 'Cantida de equipos

Private combate As cCombate 'Estrucutra de combate utilizada

Private rings() As tRing 'rings que tiene a disposicion pra utilizar el evento.

Private relacionIDPosicion() As Byte 'Relativo -> IDPos en tabla

Private necesidadRoundGanadosGanador As Byte

Private maxima_cantidad_advertencias As Byte ' Cantidad de llamados a un usuario que debe participar antes de ser descalificado

Private esperandoIngreso As Boolean

Private delayComienzo As Byte

Private Sub Class_Initialize()
    Estado = eEstadoEvento.Preparacion
    necesidadRoundGanadosGanador = 3
    esperandoIngreso = False
    maxima_cantidad_advertencias = 3
End Sub

Private Sub Class_Terminate()
    Call LogTorneos(nombreEvento & "-> Se elimina de memoria.")
End Sub

Public Sub iModoTorneo_cancelar()
    'Termino el combate
    Call terminar
End Sub

Private Sub iModoTorneo_guardar()
End Sub

Public Sub iModoTorneo_setMaximaCantidadAdvertencias(ByVal cantidad As Byte)
    maxima_cantidad_advertencias = cantidad
End Sub


Private Sub iModoTorneo_setParametros(parametros As String)

End Sub

'Tabla de equipos
Friend Sub iModoTorneo_setTablaEquipos(tabla() As tEquipoTablaTorneo)

    tablaEquipos = tabla
    
    cantidadEquipos = UBound(tablaEquipos)

    ReDim puntosEquipo(1 To cantidadEquipos)
    ReDim relacionIDPosicion(1 To cantidadEquipos)
    
    Call reiniciar

End Sub
 
Friend Sub iModoTorneo_setRings(ring() As tRing)
    rings() = ring()
End Sub

Friend Sub iModoTorneo_setRing(ring As tRing)
    ReDim rings(1 To 1) As tRing
    rings(1) = ring
End Sub

Public Sub reiniciar()
    Dim loopEquipo As Byte
    
    For loopEquipo = 1 To cantidadEquipos
        puntosEquipo(loopEquipo) = 0
    Next loopEquipo
   
End Sub

'**
' Revisa un equipo en busca de usuarios que hayan igualaldo la mayor cantidad de advertencias posibles
' Lo descalifica del evento y en caso de que este jugando en el combate, del mismo.
'@param        equipo Required. tEquipoTablaTorneo object.
'@rem
Private Sub descalificarUsuariosSuperaronAdvertencias(equipo As tEquipoTablaTorneo)

    Dim loopIntegrante As Byte
    
    For loopIntegrante = 1 To equipo.cantidadIntegrantes

        If equipo.integrantes(loopIntegrante).cantidadAdvertencias >= maxima_cantidad_advertencias _
         And Not equipo.integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando Then
            
            If Not combate Is Nothing Then
                Call combate.descalificarUsuario2(equipo.integrantes(loopIntegrante).id)
            End If
            
            'Lo descalifico
            equipo.integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando
            'Me deshecho de el
           Call modTorneos.desecharIntegranteEquipo(equipo.integrantes(loopIntegrante), 0)
            'Aviso que descalifique a uno
           Call modTorneos.enviarMensajeEquipos(tablaEquipos, equipo.integrantes(loopIntegrante).nick & " descalificado por no estar presente.", nombreEvento)
            'Aumento la cantidad de descalificados
            equipo.cantidadIntegrantesDescalificados = equipo.cantidadIntegrantesDescalificados + 1
        End If

    Next loopIntegrante
End Sub


Private Sub descalificarEquipo(equipo As tEquipoTablaTorneo)

    If Not combate Is Nothing Then
        Call combate.descalificarEquipoDelUsuario2(equipo.integrantes(1).id)
    End If
    
    'Envio el mensaje de que el equipo fue descalificado
    Call modTorneos.enviarMensajeEquipos(tablaEquipos, "Equipo de " & modTorneos.obtenerStringEquipo(equipo, False, eFormatoDisplayEquipo.NombreEquipo) & " descalificado por no encontrarse listo.", nombreEvento)
    
    'Marco el estado como descalificado
    equipo.Estado = eEstadoEquipoTorneo.descalificado
    
    'Chau participantes
    Call modTorneos.desecharEquipo(equipo, 0)

End Sub


Private Function tieneCantidadMinimaDeParticipantes(equipo As tEquipoTablaTorneo) As Boolean
    tieneCantidadMinimaDeParticipantes = (equipo.cantidadIntegrantesDescalificados < equipo.cantidadIntegrantes)
End Function

'**
'
'@param        tablaEquipos() Required. tEquipoTablaTorneo object.
'@param        cantidadEquiposConfirmados Required. Byte. Cantidad de equipos listos para jugar.
'@return       Boolean. Si estan todos los equipos listos.
'@rem
Private Function estaLaAsistenciaCorrecta(tablaEquipos() As tEquipoTablaTorneo, ByRef cantidadequiposconfirmados As Byte, Optional ByVal advertir As Boolean = True) As Boolean
    
    Dim hayError As Boolean
    Dim cantidadIntegrantesSuperaronAdvertencia As Byte
    Dim cantidadIntegrantesOffline As Byte
    Dim loopEquipo As Byte
    Dim cantidadParticipantesEquipo As Byte
    Dim loopIntegrante As Byte
    Dim UserIndex As Integer
        
    hayError = False
    
    cantidadequiposconfirmados = 0
    
    For loopEquipo = 1 To cantidadEquipos
    
        If tablaEquipos(loopEquipo).Estado = eEstadoEquipoTorneo.participando Then
            cantidadParticipantesEquipo = UBound(tablaEquipos(loopEquipo).integrantes)
            cantidadIntegrantesOffline = 0
            cantidadIntegrantesSuperaronAdvertencia = 0
            
            For loopIntegrante = 1 To cantidadParticipantesEquipo
            
                'Solo me intersan los usuarios no descalificados
                If tablaEquipos(loopEquipo).integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Jugando Then
                    
                    'Primero voy a obtener el index que ya tengo pre guardado
                    UserIndex = tablaEquipos(loopEquipo).integrantes(loopIntegrante).UserIndex
                
                    'La puta madre, no esta online
                    If UserIndex = 0 Then
                        'Si es 0 quiere decir que no esta online.
                        cantidadIntegrantesOffline = cantidadIntegrantesOffline + 1
                        
                        
                        If advertir Then
                            tablaEquipos(loopEquipo).integrantes(loopIntegrante).cantidadAdvertencias = tablaEquipos(loopEquipo).integrantes(loopIntegrante).cantidadAdvertencias + 1
                    
                            If tablaEquipos(loopEquipo).integrantes(loopIntegrante).cantidadAdvertencias >= maxima_cantidad_advertencias Then
                                cantidadIntegrantesSuperaronAdvertencia = cantidadIntegrantesSuperaronAdvertencia + 1
                            End If
                        End If
                    End If
                End If
            Next loopIntegrante
            
            
            'Algun usuario esta offline
            If cantidadIntegrantesOffline = 0 Then
                cantidadequiposconfirmados = cantidadequiposconfirmados + 1
            Else
            
                'Hay usuarios que se van?
                If cantidadIntegrantesSuperaronAdvertencia > 0 Then
                    Call descalificarUsuariosSuperaronAdvertencias(tablaEquipos(loopEquipo))
                End If
            
                'El equipo esta en condiciones de seguir participando?
                If Not tieneCantidadMinimaDeParticipantes(tablaEquipos(loopEquipo)) Then
                    Call descalificarEquipo(tablaEquipos(loopEquipo))
                Else 'El equipo puede seguir participando
                    
                    'Hay usuarios que no estan... pero aun no los vole
                    If cantidadIntegrantesSuperaronAdvertencia < cantidadIntegrantesOffline Then
                    
                        'Uno o mas de uno offline?
                        If cantidadIntegrantesOffline = 1 Then
                            Call modTorneos.enviarMensajeEquipo(tablaEquipos(loopEquipo), "Un compañero tuyo no esta online. Sino ingresa en 1 minuto será descalificado.", nombreEvento)
                        Else 'Hay mas de un integrant offline
                            Call modTorneos.enviarMensajeEquipo(tablaEquipos(loopEquipo), "Faltan integrates online en tu equipo. Sino ingresan en 1 minuto el equipo puede ser descalificado.", nombreEvento)
                        End If
    
                        hayError = True ' Los tengo que esperar un cacho mas
                    Else 'Los que estaban offline fueron sacados
                        cantidadequiposconfirmados = cantidadequiposconfirmados + 1
                    End If
                End If
            End If
        End If
    Next loopEquipo
    
    '*************************************************************************'
    estaLaAsistenciaCorrecta = Not hayError
End Function

'**
' Revisa que todos los equipos esten listos (descalificando si es necesario a los que falten)
' Si no estamos esperando a nadie y hay más de un equipo confirmado arma el combate y lo empieza.
' Si hay un solo equipo confirmado, lo declara ganador.
' Si no hay ningún equipo avisa que no quedo nadie.
'@rem
Public Sub iModoTorneo_comenzar(delay As Byte)
    Dim loopEquipo As Byte
    Dim cantidadEquiposJugando As Byte
    
    If delay = 0 Then 'Quiere que empecemos ya!
    
        Call comenzar
        
    Else
    
        cantidadEquiposJugando = obtenerCantidadEquiposJugando(tablaEquipos)
        
        If cantidadEquiposJugando > 1 Then
        
            If delay = 1 Then
                Call modTorneos.enviarMensajeEquipos(tablaEquipos, "En 1 minuto comenzará el combate: " & modTorneos.obtenerStringPrensetacion(tablaEquipos, " Vs ", True, eFormatoDisplayEquipo.NombreEquipo), nombreEvento)
            Else
                Call modTorneos.enviarMensajeEquipos(tablaEquipos, "En " & delay & " minutos comenzará el combate: " & modTorneos.obtenerStringPrensetacion(tablaEquipos, " Vs ", True, eFormatoDisplayEquipo.NombreEquipo), nombreEvento)
            End If
            
            delayComienzo = delay
            
            'Cambio el estado
            Estado = eEstadoEvento.Desarrollandose
            
        ElseIf cantidadEquiposJugando = 1 Then
        
            For loopEquipo = 1 To cantidadEquipos
                If tablaEquipos(loopEquipo).Estado = participando Then
                    Call procesarGanador(tablaEquipos(loopEquipo).idEquipo)
                    Exit For
                End If
            Next loopEquipo
        
        Else
        
            Call modTorneos.enviarMensajeGlobal("Los equipos que deberian participar fueron descalificados o abandonaron.", nombreEvento)
            Call terminar
            
        End If
    End If

    
End Sub
Private Sub comenzar()
    
    Dim loopIntegrante As Byte
    Dim loopEquipo As Byte
    Dim UserIndex As Integer
    
    Dim cantidadequiposconfirmados As Byte 'Cantidad de equipos que estan ok
    
    If Not estaLaAsistenciaCorrecta(tablaEquipos, cantidadequiposconfirmados) Then
        Call modTorneos.enviarMensajeEquipos(tablaEquipos, "No todos los equipos se encuentran listos. Se demora un minuto más el comienzo.", nombreEvento)
        esperandoIngreso = True
        Exit Sub
    End If
    
    esperandoIngreso = False
    'Si hay más de un equipo confirmado vale la pena. Si quedo uno solo fue
    If cantidadequiposconfirmados > 1 Then
        'Esta todo ok. Creo el combate.
        Set combate = New cCombate
        
        Call combate.configurarDecisionVictoria(alMejorDe, necesidadRoundGanadosGanador)
        Call combate.setCantidadEquipos(cantidadequiposconfirmados)
        
        Call combate.setRing(rings(1))

        Dim integrantes() As Integer
        Dim integrantesid() As Long
        
        Dim equipoRelativo As Byte
        Dim integranteRelativo As Byte
        
        equipoRelativo = 1
        
        For loopEquipo = 1 To cantidadEquipos
        
            If tablaEquipos(loopEquipo).Estado = participando Then

            ReDim integrantes(1 To tablaEquipos(loopEquipo).cantidadIntegrantes - tablaEquipos(loopEquipo).cantidadIntegrantesDescalificados) As Integer
            ReDim integrantesid(1 To tablaEquipos(loopEquipo).cantidadIntegrantes - tablaEquipos(loopEquipo).cantidadIntegrantesDescalificados) As Long
            
                integranteRelativo = 1
                
                For loopIntegrante = 1 To tablaEquipos(loopEquipo).cantidadIntegrantes
                            
                    If tablaEquipos(loopEquipo).integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Jugando Then
                        'Lo agrego al array que le voy a pasar al combate
                        integrantes(integranteRelativo) = tablaEquipos(loopEquipo).integrantes(loopIntegrante).UserIndex
                        integrantesid(integranteRelativo) = tablaEquipos(loopEquipo).integrantes(loopIntegrante).id
                        
                        integranteRelativo = integranteRelativo + 1
                    End If
                    
                Next loopIntegrante
                
                'Agrego el equipo al combate
                Call combate.agregarEquipo(integrantes, integrantesid, equipoRelativo)
                relacionIDPosicion(equipoRelativo) = tablaEquipos(loopEquipo).idEquipo
                
                equipoRelativo = equipoRelativo + 1
            End If
            
        Next loopEquipo
            
        'Empezamos de 0
        tiempoTranscurrido = 0
    
        'Anunciamos el combate de manera global
        Call modTorneos.enviarMensajeGlobal(modTorneos.obtenerStringPrensetacion(tablaEquipos, " Vs ", True, eFormatoDisplayEquipo.NombreEquipo), nombreEvento)
            
        'Estamos listos. Que comience.
        Call combate.comenzar
        
    ElseIf cantidadequiposconfirmados = 1 Then
    
        'Solo quedo un equipo confirmado, es el ganador
        'Busco al equipo que quedo y ese es el ganador
        For loopEquipo = 1 To cantidadEquipos
            If tablaEquipos(loopEquipo).Estado = participando Then
                Call procesarGanador(tablaEquipos(loopEquipo).idEquipo)
                Exit For
            End If
        Next loopEquipo
        
    Else 'No quedo nadie! Todos desloguearon!
        Call modTorneos.enviarMensajeGlobal("Los equipos que deberian participar fueron descalificados o abandonaron.", nombreEvento)
            
        Call terminar
    End If

   Exit Sub

End Sub
Friend Function iModoTorneo_obtenerEquipo(idEquipo As Byte) As tEquipoTablaTorneo
    Dim loopEquipo As Byte
    'Busco al equipo
    For loopEquipo = 1 To cantidadEquipos
        If tablaEquipos(loopEquipo).idEquipo = idEquipo Then
            iModoTorneo_obtenerEquipo = tablaEquipos(loopEquipo)
            Exit Function
        End If
    Next loopEquipo
        
End Function
' *************************************************************************************
' Eventos de los usuarios
Public Sub iModoTorneo_usuarioIngreso(UserIndex As Integer, UserID As Long)

    'Obtengo el equipo al cual pertenece
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    
    ' Log
    Call LogTorneos(nombreEvento & "-> Ingreso " & UserList(UserIndex).Name)
    
    'Actualizo mi info
    Call modTorneos.actualizarUserIndexPersonaje(tablaEquipos, UserID, UserIndex)
    
    'Si estoy en un combate le actualizo a el
    If Not combate Is Nothing Then
        Call combate.actualizarUserIndexPersonajeID(UserID, UserIndex)
    End If
    
   Exit Sub
    
End Sub

Public Sub iModoTorneo_usuarioAbandono(UserIndex As Integer)
    'Descalifico al integrante
    
    Dim posEquipoTabla As Byte
    Dim loopIntegrante As Byte
    Dim IDEquipoGanador_  As Byte
    Dim tipoResultado As eResultadoCombate
    
    
    Call LogTorneos(nombreEvento & "-> Abandono " & UserList(UserIndex).Name)
        
    posEquipoTabla = modTorneos.obtenerPosicionEnTablaPersonaje(tablaEquipos, UserList(UserIndex).id)
    
    With tablaEquipos(posEquipoTabla)

        For loopIntegrante = 1 To .cantidadIntegrantes

            If Not .integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando And _
                .integrantes(loopIntegrante).UserIndex = UserIndex Then
                
                If Not combate Is Nothing Then
                
                    If UserList(UserIndex).flags.Muerto = 0 And Not esperandoIngreso Then
                        IDEquipoGanador_ = combate.computarMuerte(UserIndex, tipoResultado, False)
                    End If
                    Call combate.descalificarUsuario2(.integrantes(loopIntegrante).id)
                End If
                
                'Lo descalifico
                .integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando
                'Me deshecho de el
               Call modTorneos.desecharIntegranteEquipo(.integrantes(loopIntegrante), 0)
                'Aviso que descalifique a uno
               Call modTorneos.enviarMensajeEquipos(tablaEquipos, .integrantes(loopIntegrante).nick & " abandono el evento.", nombreEvento)
                'Aumento la cantidad de descalificados
                .cantidadIntegrantesDescalificados = .cantidadIntegrantesDescalificados + 1
                
                If Not tieneCantidadMinimaDeParticipantes(tablaEquipos(posEquipoTabla)) Then
                    Call descalificarEquipo(tablaEquipos(posEquipoTabla))
                End If
                
                If IDEquipoGanador_ > 0 Then
                    Call procesarPosibleGanador(relacionIDPosicion(IDEquipoGanador_), tipoResultado)
                End If
                Exit For
            End If

        Next loopIntegrante
    End With

    
End Sub


Private Sub avanzar()
    Dim cantidadequiposconfirmados As Byte
    Dim loopEquipo As Byte
    
    If Not estaLaAsistenciaCorrecta(tablaEquipos, cantidadequiposconfirmados) Then
        Call modTorneos.enviarMensajeEquipos(tablaEquipos, "No todos los equipos se encuentran listos. Esperaremos un minuto más.", nombreEvento)
        esperandoIngreso = True
        Exit Sub
    End If
    
    esperandoIngreso = False
    If cantidadequiposconfirmados > 1 Then
        'Continuo con el combate
        Call combate.continuar
    ElseIf cantidadequiposconfirmados = 1 Then
        'Solo quedo un equipo confirmado, es el ganador
        'Busco al equipo que quedo y ese es el ganador
        For loopEquipo = 1 To UBound(tablaEquipos)
            If tablaEquipos(loopEquipo).Estado = participando Then
                Call procesarGanador(tablaEquipos(loopEquipo).idEquipo)
            End If
        Next loopEquipo
    Else 'No quedo nadie! Todos desloguearon!
        Call modTorneos.enviarMensajeGlobal("Los equipos que deberian participar fueron descalificados o abandonaron.", nombreEvento)
            
        Call terminar
    End If
End Sub
Private Sub procesarPosibleGanador(IDEquipoGanador_ As Byte, tipoResultadoCombate As eResultadoCombate)
    Dim mensaje As String
    Dim posEquipoTabla As Byte
    'Le sumo del round
        
    posEquipoTabla = modTorneos.obtenerPosTablaIDEquipo(tablaEquipos, IDEquipoGanador_)
        
    puntosEquipo(posEquipoTabla) = puntosEquipo(posEquipoTabla) + 1
    'Este round le sirvio para ganar?
    
    If tipoResultadoCombate = eResultadoCombate.final Then
        'Le doy las cosas, bla bla


        Call procesarGanador(IDEquipoGanador_)
    Else

        If tablaEquipos(posEquipoTabla).cantidadIntegrantes > 1 Then
            mensaje = "Ganador del round equipo de " & modTorneos.obtenerStringEquipo(tablaEquipos(posEquipoTabla), True, eFormatoDisplayEquipo.NombreEquipo) & "."
        Else
            mensaje = "Ganador del round " & modTorneos.obtenerStringEquipo(tablaEquipos(posEquipoTabla), False, eFormatoDisplayEquipo.NombreEquipo) & "."
        End If

        Call modTorneos.enviarMensajeEquipos(tablaEquipos, mensaje, nombreEvento)

        Call LogTorneos(nombreEvento & "-> " & mensaje)
        
        mensaje = obtenerStringPuntos()

        Call modTorneos.enviarMensajeEquipos(tablaEquipos, mensaje, nombreEvento)
        
        'Seguimos con el combate
        Call avanzar
       
    End If
End Sub
Public Sub iModoTorneo_usuarioCerro(UserIndex As Integer)

    Dim IDEquipoGanador_ As Byte
    Dim tipoResultado As eResultadoCombate
    Dim mensaje As String
    Dim equipoDescalificado As Byte


    ' ¿ Cual es el problema aca?
    ' Si consideramos que el usuario ya cerró tenemos que trabajar como si estuviese Offline
    ' Si consideramos que el usuario todavia no cerro pero que se está ejecutando el código del cerrar
        ' El problema se da si termina el Evento.
        ' a) Las lineas de actualizarUserIndexPersonaje y agregarUsuarioOffline
        ' deberian ir luego del UserDie y chequeando que el evento no haya terminado
        ' porque sino estaria haciendo algo qeu no me corresponde, ya que el fin del evento
        ' se proceso en la pila de llamadas generadas por el UserDie.
        ' Imaginemos que el Usuario Muere, se computa se  muerte, termina el combate,
        ' al terinar el combate termina el DeathMatch... este objeto quedaria invalido
        ' y el PlayOff o la liga nunca se enterarian de que el personaje esta Offline.
        
        ' b) En cambio si las lineas las pongo acá adelante, se llama luego al userdie
        ' y todos los eventos que se produzcan (finalizacion del torneo, etc)
        ' se van a hacer pensando que el usuario esta off, pero en realidad no lo está.
        
        ' c) La tercera opción sería dividir esto en dos partes.
        '       1) Cuando cierra se llama a un usuarioCerrando()
        '               Acá se lo materia y se lanzarian todos los eventos como corresponden.
        '               teniendo en cuenta que el usuario hasta acá se mantiene Online.
        '       2) Cuando se cerró totalmente, o sea que se llama al usuarioCerro()
        '               Si por algún motivo el evento termino debido a su muerte, el usuarioCerro
        '               no se llamaria
        '               Acá se agregaría el personaje a la lista de usuarios Offline
        '               sin ningun tipo de penalidad
        

    'Actualizo que el usuario cerro.
    Call modTorneos.actualizarUserIndexPersonaje(tablaEquipos, UserList(UserIndex).id, 0)
    
    'Lo agrego a la lista de users offline
    Call modEventos.agregarUsuarioOffline(UserList(UserIndex).id, UserList(UserIndex).evento)
    
    If Not combate Is Nothing Then
        '¿Esta muerto?'
        'Le aviso al combate que un userindex cambio de id.
        Call combate.actualizarUserIndexPersonajeID(UserList(UserIndex).id, 0)

        If UserList(UserIndex).flags.Muerto = 0 Then
            ' Lo mato
            ' TO-DO MMm esto esta bueno?.
            Call modUsuarios.matarUsuarioOffline(UserList(UserIndex).id)
            
            If Not esperandoIngreso Then  'TODO Mejorar
                '
                Call LogTorneos("Usuario cerro " & UserList(UserIndex).Name)
                '
                IDEquipoGanador_ = combate.computarMuerte(UserIndex, tipoResultado, True)
            
                If IDEquipoGanador_ > 0 Then
                    Call procesarPosibleGanador(relacionIDPosicion(IDEquipoGanador_), tipoResultado)
                End If
            End If
        End If
    End If
End Sub

Public Sub iModoTorneo_usuarioMuere(UserIndex As Integer)
    
    Dim IDEquipoGanador_ As Byte
    Dim tipoResultado As eResultadoCombate
    
    If Not combate Is Nothing Then
        'Si no hay ganador y esta desaroollandose es porque esta en la etapa del combate
        If Not esperandoIngreso Then 'TODO Mejorar
            Call LogTorneos("Usuario muere " & UserList(UserIndex).Name)
                    
            IDEquipoGanador_ = combate.computarMuerte(UserIndex, tipoResultado, False)
    
            If IDEquipoGanador_ > 0 Then
                Call procesarPosibleGanador(relacionIDPosicion(IDEquipoGanador_), tipoResultado)
                End If 'El morir no influyo en nada en el combate. Nadie gano el round.
            End If
    End If
    
End Sub
' ****************************************************************************


'Hay un equipo que gano el sub evento
Private Sub procesarGanador(idEquipoGanador As Byte)

    Dim auxEquipo As tEquipoTablaTorneo
    Dim loopEquipo As Byte
    Dim cantidadRoundsJugados As Byte
    Dim posEquipoTabla As Byte
    Dim mensaje As String

    posEquipoTabla = modTorneos.obtenerPosTablaIDEquipo(tablaEquipos, idEquipoGanador)

    If tablaEquipos(posEquipoTabla).cantidadIntegrantes > 1 Then
         mensaje = "¡Ganó el equipo de " & modTorneos.obtenerStringEquipo(tablaEquipos(posEquipoTabla), False, eFormatoDisplayEquipo.NombreEquipo) & "!."
     Else
         mensaje = "¡Ganó " & modTorneos.obtenerStringEquipo(tablaEquipos(posEquipoTabla), False, eFormatoDisplayEquipo.NombreEquipo) & "!."
     End If

     Call modTorneos.enviarMensajeGlobal(mensaje, nombreEvento)

     Call LogTorneos(nombreEvento & "-> " & mensaje)

    'Obtengo la cantidad de rounds jugados en tota
    cantidadRoundsJugados = 0
    For loopEquipo = 1 To cantidadEquipos
        cantidadRoundsJugados = cantidadRoundsJugados + puntosEquipo(loopEquipo)
    Next

    'Actualizo la tabla
    For loopEquipo = 1 To cantidadEquipos
        tablaEquipos(loopEquipo).cantidadCombatesJugados = tablaEquipos(loopEquipo).cantidadCombatesJugados + 1
        tablaEquipos(loopEquipo).cantidadRoundsJugados = tablaEquipos(loopEquipo).cantidadRoundsJugados + cantidadRoundsJugados
        tablaEquipos(loopEquipo).cantidadRoundsGanados = tablaEquipos(loopEquipo).cantidadRoundsGanados + puntosEquipo(loopEquipo)
    Next

    'Simplemente pongo al ganador en la pos primera de la tabla
    'Guardo el que esta en la primera posicion para no perder info

    If posEquipoTabla <> 1 Then
        auxEquipo = tablaEquipos(1)

        'Lo pongo en el primer lugar
        tablaEquipos(1) = tablaEquipos(posEquipoTabla)
        tablaEquipos(posEquipoTabla) = auxEquipo
    End If

    'Le sumo el combate que gano
    tablaEquipos(1).cantidadCombatesGanados = tablaEquipos(1).cantidadCombatesGanados + 1

    Call terminar
End Sub

Private Sub terminar()
    '
    'Termino el combate
    If Not combate Is Nothing Then
        Call combate.terminar
        
        Set combate = Nothing
    End If
    
    'El DeathMach termino
    Estado = eEstadoEvento.Terminado
End Sub
Public Sub iModoTorneo_pasarMinuto()

    If delayComienzo > 0 Then
    
        delayComienzo = delayComienzo - 1
        
        If delayComienzo = 0 Then
            Call comenzar
        End If
        
    ElseIf esperandoIngreso Then
        
        If combate Is Nothing Then
            Call comenzar
        Else
            Call avanzar
        End If
        
    End If
    
End Sub


' Devuelve true si el sub evento termino
Public Function iModoTorneo_isTerminado() As Boolean
    iModoTorneo_isTerminado = (Estado = eEstadoEvento.Terminado)
End Function

' Devuelve la tabla de equipos.
Friend Function iModoTorneo_obtenerTabla() As tEquipoTablaTorneo()
    iModoTorneo_obtenerTabla = tablaEquipos
End Function

'Establece un nombre para el evento
Public Sub iModoTorneo_setNombreEvento(nombre_ As String)
    nombreEvento = nombre_
End Sub

'Obtiene un string que contiene los equipos y su puntuacion
Private Function obtenerStringPuntos() As String

    Dim loopEquipo As Byte
    Dim cantidadColumnas As Integer
    Dim columnaActual As Byte
    
    obtenerStringPuntos = "Resultado parcial: " & vbCrLf
    
    
    columnaActual = 1
    
    If 5 - tablaEquipos(1).cantidadIntegrantes <= 0 Then
        cantidadColumnas = 1
    Else
        cantidadColumnas = 5 - tablaEquipos(1).cantidadIntegrantes
    End If

    For loopEquipo = 1 To cantidadEquipos
    
        If Not tablaEquipos(loopEquipo).Estado = eEstadoEquipoTorneo.descalificado Then
            obtenerStringPuntos = obtenerStringPuntos & modTorneos.obtenerStringEquipo(tablaEquipos(loopEquipo), True, eFormatoDisplayEquipo.NombreEquipo) & vbTab & puntosEquipo(loopEquipo)
        Else
            obtenerStringPuntos = obtenerStringPuntos & modTorneos.obtenerStringEquipo(tablaEquipos(loopEquipo), False, eFormatoDisplayEquipo.NombreEquipo) & vbTab & "(des)"
        End If
        
        If loopEquipo < cantidadEquipos Then
            
            columnaActual = columnaActual + 1
            
            If columnaActual > cantidadColumnas Then
                obtenerStringPuntos = obtenerStringPuntos & vbCrLf
                columnaActual = 1
            Else
                obtenerStringPuntos = obtenerStringPuntos & " - "
            End If
            
        Else
            obtenerStringPuntos = obtenerStringPuntos & "."
        End If
        
    Next

End Function

'TODO re ver
Public Sub iModoTorneo_setNecesidadRoundGanadosGanador(cantidad As Byte)
    necesidadRoundGanadosGanador = cantidad
End Sub

Public Sub iModoTorneo_cargar(idliga As Long)
End Sub

Private Function iModoTorneo_obtenerIDPersistencia() As Integer
End Function

Public Sub iModoTorneo_usuarioRevive(UserIndexRevivido As Integer, UserIndexRevividor As Integer)
    If Not combate Is Nothing Then
        If Not esperandoIngreso Then 'TODO Mejorar
            Call combate.reviviruUsuario(UserIndexRevivido)
        End If
    End If
End Sub
