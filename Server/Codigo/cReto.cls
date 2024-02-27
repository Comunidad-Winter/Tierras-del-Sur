VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iEvento

Option Explicit

'Define que ocsas estan en juego en el reto
Public Enum eModoApuesta
    oro = 1
    Items = 2
    oroitems = 3
End Enum

Private Type tIntegranteEquipoReto
    IDIntegrante As Long
    posOrig  As WorldPos 'Posicion desde donde lo traje y a donde lo voy a devolver
    Acepto As Byte '0: Aún no aceptó el reto. 1: sí lo aceptó. Tambien es usado para saber si el personaje cerro cuando estaba recogiendo los items. (1 no, 0 si)
    UserIndex As Integer ' UserIndex de este integrante
    nick As String 'Nick del integrante. Lo dejo por si alguno cierra. Si cierra no tengo forma de saberlo
End Type

Private Type tEquipoReto
    id As Byte 'Indentificador del equipo
    cantidadIntegrantes As Byte 'cantidad de integrantes del equipo
    integrantes() As tIntegranteEquipoReto 'Integrantes del equipo
    descalificado As Byte '0. El equipo no fue descalificado. 1 El equipo fue descalificado
    cantidadRoundsGanados As Byte 'Cantidad de rounds que gano el equipo
End Type

Private reglaHechizos() As Boolean 'Indica que hechizos estan permitidos usar y cuales no durante el desarrollo del evento

Private cantidadOroApostado As Long 'Cantidad de oro que fue apostado
Private modoApuesta As eModoApuesta 'El reto es por items o por oro o por ambas.

Private cantidadSolicitudesAceptadas As Byte 'Cantidad de solicitud de participar en el reto aceptada'
                                              'Esta variable esta para optimizar, ahorrandonos recorrer todos los equipos para ver si aceptaron todos cada vez que acepta un integrante
Private cantidadParticipantes As Byte 'Cantidad de personas en total que participan

Private cantidadEquipos As Byte 'cantidad de equipos que estan jugando en el reto
Private equipos() As tEquipoReto 'Equipos que conforman el reto

Private tiempoTranscurrido As Integer 'Minutos desde que el reto se esta desarrollando

Private tiempoMaximo As Byte 'Tiempo maximo, en minutos, que puede durar el reto

Private combate As cCombate 'Combate que utiliza el reto
Private descanso As tZonaDescanso 'Donde voy a poner a los individuos para explotarlos cuando el reto es por items
                                  'En los retos solo por oro no tiene sentido que pasen por acá
Private ring As tRing 'Ring que le vamos a dar al combate para que se haga la pelea
Private tipoRing As eRingTipo

Private Estado As eEstadoEvento 'Estado del evento

Private idEquipoGanador As Byte 'ID del equipo ganador

Private premioOroExtra As Long 'Oro extra de premio que se suma al pozo para repartir.
                                'Esta porsi el reto lo organiza un gm
                                
Private nombreReto As String 'Nombre del reto

Private listaCondiciones As Collection  ' Condiciones que deben cumplir para poder ingresar


Private Sub Class_Initialize()
    'Me aseguro que arranque en 0
    tiempoTranscurrido = 0
    'Cargo los default
    modoApuesta = eModoApuesta.oroitems
    cantidadOroApostado = 5000
    'Por ahora no uso ningún combate
    Set combate = Nothing
    Set listaCondiciones = Nothing
    
    tipoRing = eRingTipo.ringReto
    
    cantidadParticipantes = 0
    'Estado
    Estado = eEstadoEvento.Preparacion
    'Tiempo maximo
    tiempoMaximo = 20 '20 minutos
    '
    idEquipoGanador = 0
    'Por defecto no se le da nada de oro, solo lo que apostaron
    premioOroExtra = 0
    '
    nombreReto = "Reto"
    'Reglas
    Call modTorneos.cargarReglasbasicasHechizos(reglaHechizos)
    
End Sub

Public Sub setValeResu(vale As Boolean)
    reglaHechizos(eHechizos.Resucitar) = vale
End Sub
'Establece el modo de la apuesta y el monto
Public Sub setModoApuesta(modo As eModoApuesta, cantidadOro As Long)
    modoApuesta = modo
    
    If modoApuesta = oro Or modoApuesta = oroitems Then
        cantidadOroApostado = cantidadOro
    Else
        cantidadOroApostado = 0
    End If
End Sub

'Establece la cantidad de equipos que van a jugar al reto
Public Sub setCantidadEquipos(cantidad As Byte)

    cantidadEquipos = cantidad
    
    ReDim equipos(1 To cantidad) As tEquipoReto

End Sub
Public Sub setOroPremioExtra(cantidad As Long)
    premioOroExtra = cantidad
End Sub

Public Sub setNombre(nombre As String)
    nombreReto = nombre
End Sub

'establece el tiempo maximo que puede durar el evento
Public Sub setTiempoMaximo(minutos As Byte)
    tiempoMaximo = minutos
End Sub

'Agrega los equipos
Public Function agregarEquipo(integrantesid() As Long, idEquipo As Byte) As eResultadoInscripcion

    Dim loopIntegrante As Byte
    
    equipos(idEquipo).id = idEquipo
    equipos(idEquipo).cantidadIntegrantes = UBound(integrantesid)
    equipos(idEquipo).descalificado = 0
    equipos(idEquipo).cantidadRoundsGanados = 0
    
    ReDim equipos(idEquipo).integrantes(1 To UBound(integrantesid))
    For loopIntegrante = 1 To equipos(idEquipo).cantidadIntegrantes
        equipos(idEquipo).integrantes(loopIntegrante).IDIntegrante = integrantesid(loopIntegrante)
        equipos(idEquipo).integrantes(loopIntegrante).Acepto = 0

    Next loopIntegrante
    
    'Sumo a estas personas a la cantidad de integrantes
    cantidadParticipantes = cantidadParticipantes + equipos(idEquipo).cantidadIntegrantes

    agregarEquipo = correcta
End Function

'Devuelve true si el resto esta listo para ser jugado.
'Hay mas de un equipo y todos aceptaron la solicitud
Public Function listoParaEmpezar() As Boolean
    listoParaEmpezar = (cantidadSolicitudesAceptadas = cantidadParticipantes) And (cantidadEquipos > 1)
End Function

Public Function formaParteDelReto(IDIntegrante As Long)
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    
    'Busco al usuario dentro de los equipos
    For loopEquipo = 1 To cantidadEquipos
    
      For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            'Es este?
            If equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante = IDIntegrante Then
                formaParteDelReto = True
                Exit Function
            End If
        Next loopIntegrante
    Next loopEquipo

    formaParteDelReto = False
    
End Function

Public Function rechazarSolicitud(IDIntegrante As Long, UserIndex As Integer) As Boolean
    
    If Not formaParteDelReto(IDIntegrante) Then
        rechazarSolicitud = False
        Exit Function
    End If

    ' Cancelamos la aceptacion de estar persona
    Call cancelarSolicitudIDIntegrante(IDIntegrante)

    ' Enviamos
    Call enviarMensajeParticipantesQueAceptaron(UserList(UserIndex).Name & " rechazó el reto.")
    
    ' Cancelamos el reto
    Call iEvento_cancelar
    
    rechazarSolicitud = True
End Function

Private Sub cancelarSolicitudIDIntegrante(IDIntegrante As Long)

    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
        
    'Busco al usuario dentro de los equipos
    For loopEquipo = 1 To cantidadEquipos
    
      For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            'Es este?
            If equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante = IDIntegrante Then
                Call cancelarAceptacionSolicitud(equipos(loopEquipo).integrantes(loopIntegrante))
                Exit Sub
            End If
        Next loopIntegrante
    Next loopEquipo
    
End Sub

'Intenta aceptar la solicitud. Devuelve el reusltado
'0: Solicitud aceptada
'1: Ya habia sido aceptada
'2: No fue invitado

Public Function aceptarSolicitud(IDIntegrante As Long, UserIndex As Integer) As Byte

    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
        
    'Busco al usuario dentro de los equipos
    For loopEquipo = 1 To cantidadEquipos
    
      For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            'Es este?
            If equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante = IDIntegrante Then
                
                If equipos(loopEquipo).integrantes(loopIntegrante).Acepto = 0 Then
                    'Acepta
                    equipos(loopEquipo).integrantes(loopIntegrante).Acepto = 1
                    equipos(loopEquipo).integrantes(loopIntegrante).nick = UserList(UserIndex).Name
                    
                    cantidadSolicitudesAceptadas = cantidadSolicitudesAceptadas + 1
                    aceptarSolicitud = 0
                Else
                    'Ya aceptó
                    aceptarSolicitud = 1
                End If
                
                Exit Function
            End If
        Next loopIntegrante
    Next loopEquipo

    aceptarSolicitud = 2 'No lo encontre

End Function

Public Sub cancelarTodasLasAceptaciones()
    
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    
    'Busco al usuario dentro de los equipos

    For loopEquipo = 1 To cantidadEquipos
    
        For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            Call cancelarAceptacionSolicitud(equipos(loopEquipo).integrantes(loopIntegrante))
        Next loopIntegrante
        
    Next loopEquipo
    
End Sub
Private Sub cancelarAceptacionSolicitud(ByRef Integrante As tIntegranteEquipoReto)

    'Marco que el integrante no acepto
    If Integrante.Acepto = 1 Then
        Integrante.Acepto = 0
        'Marco que hay una solicitud menos aceptada
        cantidadSolicitudesAceptadas = cantidadSolicitudesAceptadas - 1
    End If
End Sub
Public Sub comenzar()

    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    Dim UserIndex As Integer
    Dim enOtroEvento As Boolean
    Dim maxCantidadIntegrantes As Byte 'Cantidad maxima de integrantes que hay en un equipo
    
    Dim hayError As Boolean
    'Validaciones finales
    'Lo mas importante es el oro!!! todos tienen que tener el oro suficiente!

    hayError = False
    maxCantidadIntegrantes = 0
    'Cuando hay un error, ovbiamente el reto no se inicia.
    'El reto no se cancela. Sino que los usuarios tienen tiempo de solucionar el problema
    'y volver a aceptar el reto.
    'Por lo tanto cuando el usuario tiene un error se cancela su aceptacion de solicitud
    For loopEquipo = 1 To cantidadEquipos
        For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            
            'Obtengo el userindex del usuario a partir de su ID
            UserIndex = IDIndex(equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante)
            
            If UserIndex > 0 Then
            
                'El pesonaje esta jugando otro evento en estos momentos?
                enOtroEvento = False
                
                'Debe tener un evento asociado y este no estar desarolandose
                If Not UserList(UserIndex).evento Is Nothing Then
                    If UserList(UserIndex).evento.getEstadoEvento = eEstadoEvento.Desarrollandose Then
                        enOtroEvento = True
                    End If
                End If
                
                'Si no esta en otro evento, ok.
                If Not enOtroEvento Then
                    If UserList(UserIndex).flags.Muerto = 0 Then
                        If MapInfo(UserList(UserIndex).pos.map).Pk = False Then
                            If UserList(UserIndex).Stats.GLD >= (VALOR_RETO + cantidadOroApostado) Then
                            
                                If puedeIngresar(UserIndex) Then
                                    'Todo OK. Ya que lo tengo guardo el userindex
                                    equipos(loopEquipo).integrantes(loopIntegrante).UserIndex = UserIndex
                                Else
                                    hayError = True
                                    Call enviarMensajeParticipantes("El personaje " & UserList(UserIndex).Name & " no cumple con las condiciones para ingresar al reto.")
                                    Call cancelarAceptacionSolicitud(equipos(loopEquipo).integrantes(loopIntegrante))
                                End If
                            Else
                                'El usuario no tiene el oro
                                hayError = True
                                Call enviarMensajeParticipantes("El personaje " & UserList(UserIndex).Name & " no tiene el oro suficiente.")
                                Call cancelarAceptacionSolicitud(equipos(loopEquipo).integrantes(loopIntegrante))
                            End If
                        Else
                            hayError = True
                            Call enviarMensajeParticipantes("El personaje " & UserList(UserIndex).Name & " se encuentra en zona insegugra.")
                            Call cancelarAceptacionSolicitud(equipos(loopEquipo).integrantes(loopIntegrante))
                        End If
                    Else
                        hayError = True
                        Call enviarMensajeParticipantes("El personaje " & UserList(UserIndex).Name & " esta muerto. Así no puuede jugar un reto.")
                        Call cancelarAceptacionSolicitud(equipos(loopEquipo).integrantes(loopIntegrante))
                    End If
                Else 'El usuario esta participando en otro evento
                    'Esto permite que cuando termine el otro evento pueda participar en este
                    hayError = True
                    Call enviarMensajeParticipantes("El personaje " & UserList(UserIndex).Name & " esta participando en otro evento.")
                    Call cancelarAceptacionSolicitud(equipos(loopEquipo).integrantes(loopIntegrante))
                End If
            Else
                'Un usuario no esta online
                hayError = True
                Call enviarMensajeParticipantes("El personaje " & equipos(loopEquipo).integrantes(loopIntegrante).nick & " no se encuentra online")
                Call cancelarAceptacionSolicitud(equipos(loopEquipo).integrantes(loopIntegrante))
            End If
            
        Next loopIntegrante
        
        If maxCantidadIntegrantes < equipos(loopEquipo).cantidadIntegrantes Then
            maxCantidadIntegrantes = equipos(loopEquipo).cantidadIntegrantes
        End If
        
    Next loopEquipo
    
    'Si hay un error, me voy.
    If hayError Then Exit Sub
    
    'SI LLEGUE HASTA ACÁ ES PORQUE LOS USUARIOS ESTAN OK

    'Obtengo el ring
    
    ring = modRings.obtenerRing(cantidadEquipos, maxCantidadIntegrantes, tipoRing)
    
    If ring.id = 0 Then
        'No hay un ring disponible, se cancela el reto
        Call enviarMensajeParticipantes("Lo sentimos, pero en estos momentos todas las arenas de combate de Tierras del Sur se encuentran ocupadas. Esperen unos minutos y vuelvan todos (incluido el creador) a aceptar el reto.")
        Call cancelarTodasLasAceptaciones
        Exit Sub
    End If

    'Obtengo el lugar que voy a utilizar para explotarlos en caso
    'de ser un reto por items
    If modoApuesta = eModoApuesta.Items Or modoApuesta = eModoApuesta.oroitems Then
    
        descanso = modDescansos.getZonaDescanso(cantidadParticipantes, (eDescansoTipo.reto + eDescansoTipo.conBoveda))
        
        Debug.Print "Tomo el descanso con el id " & descanso.id
        'Hay un descanso disponible?
        If descanso.id = 0 Then
            'No hay una zona de descanso disponible, se cancela el reto
            Call enviarMensajeParticipantes("Lo sentimos, pero en estos momentos todas las arenas de combate de Tierras del Sur se encuentran ocupadas. El reto ha sido cancelado.")
            Call iEvento_cancelar
            Exit Sub
        End If
    End If

    'TENEMOS TODO LO NECESARIO PARA ARMAR EL RETO!

    Dim nPos As WorldPos
    
    'Creo el combate donde se van a llevar a cbo la pelea
    Set combate = New cCombate
    
    'Le digo en que arena tiene que llevar a cabo el combate
    Call combate.setRing(ring)
    'Establezco la cantidad de equipos
    Call combate.setCantidadEquipos(cantidadEquipos)
                
    'Si llegue hasta acá es porque esta todo ok
    'Genero lso equipos y le asigno a cada usuario el evento
    Dim integrantes() As Integer
    Dim integrantesid() As Long
    
    For loopEquipo = 1 To cantidadEquipos
    
        ReDim integrantes(1 To equipos(loopEquipo).cantidadIntegrantes) As Integer
        ReDim integrantesid(1 To equipos(loopEquipo).cantidadIntegrantes) As Long
        
        For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            
            'Obtengo el userindex
            UserIndex = equipos(loopEquipo).integrantes(loopIntegrante).UserIndex
            
            'Le saco el oro se los saco.
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - (cantidadOroApostado + VALOR_RETO)
            'Lo actualizo
            EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
             
            'Si tiene mascotas se las saco
            Call BorrarMascotas(UserIndex)
            
            'Obtengo el lugar donde esta el usuario
            equipos(loopEquipo).integrantes(loopIntegrante).posOrig = UserList(UserIndex).pos
            
            'Guardo el nombre, por si cierra, ya lo tengo y no tengo que consultar algo como
            'la base de datos
            equipos(loopEquipo).integrantes(loopIntegrante).nick = UserList(UserIndex).Name
            
            'Lo muevo al descanso en caso de existir
            If descanso.id > 0 Then
                Call ClosestLegalPos(descanso.centro, nPos, UserList(UserIndex))
                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.y)
            End If
            
            'Lo agrego al array que le voy a pasar al combate
            integrantes(loopIntegrante) = UserIndex
            integrantesid(loopIntegrante) = UserList(UserIndex).id
            
            'Relaciono el evento al usuario
            Set UserList(UserIndex).evento = Me
            
        Next loopIntegrante
        
        'Agrego el equipo al combate
        Call combate.agregarEquipo(integrantes, integrantesid, loopEquipo)
        
    Next loopEquipo
    'Guardo en el logs
        
    Dim mensaje As String
    
    mensaje = obtenerStringPrensetacion()
    Call enviarMensajeParticipantes(mensaje)
    Call LogNuevosRetos("Comienza el reto " & mensaje & " con ring " & ring.id & " y descanso " & descanso.id)
    
    'Listo, comienza el combate
    Call combate.comenzar
    
    Call modRetos.comenzoReto
    'El reto cambia de estado a "desarrollandose"
    Estado = eEstadoEvento.Desarrollandose
    
    'Tiempo desde que el reto empezo a desarrollarse
    tiempoTranscurrido = 0

End Sub

'Ya tenemos un ganador
'Entonces hacemos las operativas pertinentenes

Private Sub procesarGanador(IDEquipoGanador_ As Byte)
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    Dim oroOtorgar  As Long 'Cantidad de oro que le corresponde a cada ganador
    Dim UserIndex As Integer
    Dim mensaje As String
    
    'Aviso el ganador
    
    mensaje = generarMensajeGanador(IDEquipoGanador_)

    'Guardo en el log
    If modoApuesta = eModoApuesta.oro Then 'Por oro
        mensaje = mensaje & " Apuesta por " & cantidadOroApostado & " monedas de oro."
    ElseIf modoApuesta = eModoApuesta.Items Then 'El reto es solo por los items
        mensaje = mensaje & " Apuesta por los items."
    Else  'Por el oro y pos los items
        mensaje = mensaje & " Apuesta por los items y " & cantidadOroApostado & " monedas de oro."
    End If

    'Call enviarMensajeParticipantes(mensaje)
    
    'Le aviso a toda la gente.
    EnviarPaquete Paquetes.mensajeinfo, nombreReto & "> " & obtenerStringPrensetacion() & " " & mensaje, 0, ToAll

    Call LogNuevosRetos(mensaje)
    '
    'Hay un ganador. El combate termino
    Call combate.terminar

    'Calculo cuanto es el oro que le voy a tner que dar al usuario
    If modoApuesta = eModoApuesta.oro Or modoApuesta = oroitems Then
        'El premio en oro es el equivalente a la suma de lo apostado
        'por cada uno de los que participaron en el reto
        'Todos porque previamente le saque el oro también a los que ahora ganaron.
        'A eso eventualmente se le puede sumar un premio extra en oro
        'que debe ser decidio por los gms y que no se da en los retos normales
        'Dividido la cantidad de integrantes que tiene el equipo.
        'El premio se reparte por partes iguales
        oroOtorgar = Int(((cantidadOroApostado * cantidadParticipantes) + premioOroExtra) / equipos(IDEquipoGanador_).cantidadIntegrantes)
    Else
        oroOtorgar = Int(premioOroExtra / equipos(IDEquipoGanador_).cantidadIntegrantes)
    End If
    

    'Recorro los equipos. Actualizo las estadisticas y los mando devuelta a casa
    For loopEquipo = 1 To cantidadEquipos

        If loopEquipo = IDEquipoGanador_ Then
            'Es el ganador
            'Le tengo que dar oro
            For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes

                UserIndex = equipos(loopEquipo).integrantes(loopIntegrante).UserIndex

                'Los ganadores se merecen terminar vivos.
                'Lo preparo por si esta muerto, inmovil, etc
                Call Preparando(UserIndex)

                'Actualizo sus estadisticas
                'Estadisticas solo en los retos 1vs1 y en el FACIL. ESTO es un poco feo
                '******************************************************************************
                '           ESTADISTICAS
                If equipos(loopEquipo).cantidadIntegrantes = 1 Then
                    'Actualizo la cantidad de retos ganados
                    UserList(UserIndex).Stats.RetosGanadoS = UserList(UserIndex).Stats.RetosGanadoS + 1
                    'Actualizo el oro ganado
                    If oroOtorgar > 0 Then UserList(UserIndex).Stats.OroGanado = UserList(UserIndex).Stats.OroGanado + (oroOtorgar - cantidadOroApostado)
                
                    #If TDSFacil Then
                        If loopEquipo = 1 Then
                            Call modRetos.registrar(UserList(UserIndex), UserList(equipos(2).integrantes(1).UserIndex), (oroOtorgar - cantidadOroApostado))
                        Else
                            Call modRetos.registrar(UserList(UserIndex), UserList(equipos(1).integrantes(1).UserIndex), (oroOtorgar - cantidadOroApostado))
                        End If
                    #End If
                End If
                '******************************************************************************
                
                'Hay que darle oro?
                If oroOtorgar > 0 Then
                    'Le doy el oro
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + oroOtorgar
                    'Se lo actualizo
                    EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex
                End If
        
                'Es por los items?
                If modoApuesta = eModoApuesta.Items Or modoApuesta = eModoApuesta.oroitems Then
                    'Les digo que tienen 30 segundos para agarrar las cosas
                    EnviarPaquete Paquetes.MensajeTalk, nombreReto & "> Tienes 2 minutos para recoger los items. Si terminas antes puedes escribir /ABANDONAR para salir del evento.", UserIndex, ToIndex
                Else
                    'Sino es por los items. Listo. No lo necesito más.
                    'Lo mando donde estaba
                    Call WarpUserChar(UserIndex, equipos(loopEquipo).integrantes(loopIntegrante).posOrig.map, equipos(loopEquipo).integrantes(loopIntegrante).posOrig.x, equipos(loopEquipo).integrantes(loopIntegrante).posOrig.y, True)
                    'Le quito la referencia del evento
                    Set UserList(UserIndex).evento = Nothing
                End If
                    
            Next loopIntegrante
        Else 'No es el equipo ganador
                'EQUIPO PERDEDOR
                Call procesarEquipoPerdedor(loopEquipo)
        End If
    Next loopEquipo
       
    tiempoTranscurrido = 0
    
    idEquipoGanador = IDEquipoGanador_
    
    'Si el reto no es por los items. El reto termino.
    If modoApuesta = eModoApuesta.oro Then
        Call terminar
    End If

End Sub

Private Sub procesarEquipoPerdedor(idEquipo As Byte)
    
    Dim loopIntegrante As Byte
    Dim UserIndex As Integer
    
    'Si esta descalificado no lo proceso porque ya lo procese antes

    If equipos(idEquipo).descalificado = 0 Then
        For loopIntegrante = 1 To equipos(idEquipo).cantidadIntegrantes
    
            UserIndex = equipos(idEquipo).integrantes(loopIntegrante).UserIndex
            
            'Cantidad de retos perdidos
            If equipos(idEquipo).cantidadIntegrantes = 1 Then
                'Actualizo las estadisticas
                UserList(UserIndex).Stats.RetosPerdidosB = UserList(UserIndex).Stats.RetosPerdidosB + 1
                'Actualizo el oro perdido
                UserList(UserIndex).Stats.OroPerdido = UserList(UserIndex).Stats.OroPerdido + cantidadOroApostado
            End If
            
            If modoApuesta = eModoApuesta.Items Or modoApuesta = eModoApuesta.oroitems Then
                'Le quito los items, muejeje,
                Call TirarTodo(UserList(UserIndex))
                'Lo mando del repositorio a donde estaba. Es un perdedor, no me sirve.
                Call WarpUserChar(UserIndex, equipos(idEquipo).integrantes(loopIntegrante).posOrig.map, equipos(idEquipo).integrantes(loopIntegrante).posOrig.x, equipos(idEquipo).integrantes(loopIntegrante).posOrig.y, True)
            End If

            'Le quito la referencia del evento
            Set UserList(UserIndex).evento = Nothing
            
        Next
    End If
    
    'El equipo perdio
    equipos(idEquipo).descalificado = 1

End Sub

Public Function puedeIngresar(UserIndex As Integer) As Boolean
    Dim condicion As iCondicionEvento
    Dim userIndexs(1 To 1) As Integer
    
    ' ¿Cumple la condicion?
    If Not listaCondiciones Is Nothing Then
        userIndexs(1) = UserIndex
        For Each condicion In listaCondiciones
            If Not condicion.puedeIngresarEquipo(userIndexs) Then
                 puedeIngresar = False
                Exit Function
            End If
        Next
    End If
    
    puedeIngresar = True
End Function

Public Sub iEvento_agregarCondicionIngreso(condicion As iCondicionEvento)

If listaCondiciones Is Nothing Then Set listaCondiciones = New Collection

' Agregamos la condicion
Call listaCondiciones.Add(condicion)

End Sub

Public Sub iEvento_cancelar()
'Cancelo el combate. Nadie gana.

If Estado = eEstadoEvento.Preparacion Then
    'No cancelo el combate porque no se inicio...
    'Los personajes tampoco fueron transportados
    'hay que sacarle la relacion de evento.
    Call desasAsignarEvento
ElseIf Estado = eEstadoEvento.Desarrollandose Then
    'Cancelo el combate
    Call combate.cancelar
    'Como nadie gano. Los descalifico a todos
    Call descalificarTodosEquipos
End If

'No hubo ningún ganador ya que cancele el reto
idEquipoGanador = 0
'Termino el reto
Call terminar

End Sub

Private Sub iEvento_establecerTablaDePagos(tablaPagos() As Long, tipoExpresion As eEventoPremio)
End Sub

Private Function iEvento_getCantidadEquiposMax() As Byte
    iEvento_getCantidadEquiposMax = 2
End Function

Private Sub iEvento_publicar()
End Sub

' ¿El usuario puede tirar un objeto?
Public Function iEvento_puedeTirarObjeto(UserIndex As Integer, objeto As Integer, ByVal cantidad As Long, destino As eDestinoObjeto, destinoIndex As Integer) As Boolean

If Estado = Desarrollandose Then
    If destino = eDestinoObjeto.criatura Then
        iEvento_puedeTirarObjeto = True
    Else
        iEvento_puedeTirarObjeto = False
        EnviarPaquete Paquetes.mensajeinfo, "No puedes tirar objetos al suelo en este momento.", UserIndex, ToIndex
    End If
Else
    iEvento_puedeTirarObjeto = True
End If

End Function

Private Sub iEvento_configurarApuestas(ByVal activadas As Boolean, ByVal pozoInicial As Long, ByVal tiempo As Byte)
End Sub

Private Function iEvento_establecerGanadorManualmente(nombreEquipoGanador As String) As Boolean
End Function


Private Sub iEvento_guardar()

End Sub

Private Function iEvento_isCumpleCondicionEquipo(userIndexs() As Integer) As Boolean
End Function

Private Function iEvento_obtenerIDPersistencia() As Integer

End Function

Private Function iEvento_obtenerInfoEstado() As String

End Function

Private Function iEvento_obtenerInfoExtendida() As String

End Function

Private Sub iEvento_setHechizosPermitidos(reglas() As Boolean)
End Sub

Private Sub iEvento_setHojaDeRuta(hojaRuta As String, parametros As String)
End Sub

Private Sub iEvento_setIDPersistencia(IDPersistencia As Long)

End Sub

Public Sub iEvento_setTiporing(tipoRing_ As eRingTipo)
    tipoRing = tipoRing_
End Sub

' Funciones de la interface
Public Sub iEvento_usuarioMuere(UserIndex As Integer)
    
    Dim IDEquipoGanador_ As Byte
    Dim tipoResultado As eResultadoCombate
    
    If Estado = Desarrollandose Then
        'Si no hay ganador y esta desaroollandose es porque esta en la etapa del combate
        If idEquipoGanador = 0 Then '
        
            IDEquipoGanador_ = combate.computarMuerte(UserIndex, tipoResultado, False)
    
            If IDEquipoGanador_ > 0 Then
                'Le sumo del round
                equipos(IDEquipoGanador_).cantidadRoundsGanados = equipos(IDEquipoGanador_).cantidadRoundsGanados + 1
                'Este round le sirvio para ganar?
                    
                If tipoResultado = eResultadoCombate.final Then
                    'Le doy las cosas, bla bla
                    Call procesarGanador(IDEquipoGanador_)
                Else
                    'Aviso...
                    Dim mensaje As String
                    
                    If equipos(IDEquipoGanador_).cantidadIntegrantes > 1 Then
                        mensaje = "Ganador del round equipo de " & obtenerStringEquipo(IDEquipoGanador_) & "."
                    Else
                        mensaje = "Ganador del round " & obtenerStringEquipo(IDEquipoGanador_) & "."
                    End If
                    
                    Call enviarMensajeParticipantes(mensaje)
                    'Call LogNuevosRetos(mensaje)
                    
                    mensaje = obtenerStringPuntos()
                    
                    Call enviarMensajeParticipantes(mensaje)
                    'Call LogNuevosRetos(mensaje)
                    'Continuo con el combate
                    Call combate.continuar
                End If
            End If 'El morir no influyo en nada en el combate. Nadie gano el round.
            
        Else    'Esta en el descanso agarrando los items
            Call modUtilitarios.Preparando(UserIndex)
        End If
    End If
    End Sub
Private Function generarMensajeGanador(idEquipoGanador As Byte) As String
    Dim loopIntegrante As Byte
    Dim cantidadIntegrantes As Byte
    
    cantidadIntegrantes = equipos(idEquipoGanador).cantidadIntegrantes
    
    If cantidadIntegrantes > 1 Then 'Es un equipo?
    
    generarMensajeGanador = "Ganador el equipo de "
    
    For loopIntegrante = 1 To cantidadIntegrantes
        generarMensajeGanador = generarMensajeGanador & equipos(idEquipoGanador).integrantes(loopIntegrante).nick
        If loopIntegrante = cantidadIntegrantes - 1 Then
            generarMensajeGanador = generarMensajeGanador & " y "
        ElseIf loopIntegrante < cantidadIntegrantes Then
            generarMensajeGanador = generarMensajeGanador & ", "
        End If
    Next
    
    generarMensajeGanador = generarMensajeGanador & "."

    Else 'No es solo una persona, entonces no lo consideramos "equipo"
        generarMensajeGanador = "Ganador " & equipos(idEquipoGanador).integrantes(1).nick & "."
    End If
    
End Function

'Envia al usuario a su casa
Private Sub deshacermeUsuario(idEquipo As Byte, NumIntegrante As Byte)
    
    Dim UserIndex As Integer
        
    UserIndex = equipos(idEquipo).integrantes(NumIntegrante).UserIndex
    equipos(idEquipoGanador).integrantes(NumIntegrante).Acepto = 0
    Call WarpUserChar(UserIndex, equipos(idEquipo).integrantes(NumIntegrante).posOrig.map, equipos(idEquipo).integrantes(NumIntegrante).posOrig.x, equipos(idEquipo).integrantes(NumIntegrante).posOrig.y)
    Set UserList(UserIndex).evento = Nothing
    
End Sub

'Obtiene el equipo al cual pertenece el personaje
Private Function obtenerEquipo(idPersonaje As Long) As Byte
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte

    For loopEquipo = 1 To cantidadEquipos
        For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            If equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante = idPersonaje Then
                obtenerEquipo = loopEquipo
                Exit Function
            End If
        Next
    Next loopEquipo
    
    obtenerEquipo = 0

End Function

Private Sub desasAsignarEvento()

    Dim loopIntegrante As Byte
    Dim loopEquipo As Byte
    Dim UserIndex As Integer
    
    For loopEquipo = 1 To cantidadEquipos
        For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
        
            UserIndex = IDIndex(equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante)
                    
            'Le quito la referencia del evento
            If UserIndex > 0 Then
                If UserList(UserIndex).evento Is Me Then
                    Set UserList(UserIndex).evento = Nothing
                End If
            End If
        Next
    Next loopEquipo

End Sub
'Descalifica a todos los equipos.
'Todos perdieron.
'Esto pasa cuando no se matan antes de los 15 minutos.
Private Sub descalificarTodosEquipos()
    
    Dim loopIntegrante As Byte
    Dim loopEquipo As Byte
    Dim UserIndex As Integer
    
    
    'Me deshagado de los equipos
    'Si esta descalificado no lo proceso porque ya lo procese antes
    
    For loopEquipo = 1 To cantidadEquipos
    
        If equipos(loopEquipo).descalificado = 0 Then
            For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
        
                UserIndex = equipos(loopEquipo).integrantes(loopIntegrante).UserIndex
                
                'No actualizo las estadisticas porque todos perdieron.
                'Cantidad de retos perdidos
                'UserList(UserIndex).Stats.RetosPerdidosB = UserList(UserIndex).Stats.RetosPerdidosB + 1
                'Actualizo el oro perdido
                'UserList(UserIndex).Stats.OroPerdido = UserList(UserIndex).Stats.OroPerdido + cantidadOroApostado
                
                If modoApuesta = eModoApuesta.Items Or modoApuesta = eModoApuesta.oroitems Then
                    'No le quito los items. Ya que es al pedo
                    'Call TirarTodo(UserIndex)
                    'Lo mando del repositorio a donde estaba.
                    Call WarpUserChar(UserIndex, equipos(loopEquipo).integrantes(loopIntegrante).posOrig.map, equipos(loopEquipo).integrantes(loopIntegrante).posOrig.x, equipos(loopEquipo).integrantes(loopIntegrante).posOrig.y, True)
                End If
    
                'Le quito la referencia del evento
                Set UserList(UserIndex).evento = Nothing
                
            Next
        End If
    Next loopEquipo
    
End Sub

Private Function obtenerStringPuntos() As String
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    Dim cantidadIntegrantes As Byte
    
    obtenerStringPuntos = "Resultado parcial: "
    
    For loopEquipo = 1 To cantidadEquipos
    
        cantidadIntegrantes = equipos(loopEquipo).cantidadIntegrantes
        
        For loopIntegrante = 1 To cantidadIntegrantes
            obtenerStringPuntos = obtenerStringPuntos & equipos(loopEquipo).integrantes(loopIntegrante).nick
        
            If loopIntegrante = cantidadIntegrantes - 1 Then
                obtenerStringPuntos = obtenerStringPuntos & " y "
            ElseIf loopIntegrante < cantidadIntegrantes Then
                obtenerStringPuntos = obtenerStringPuntos & " , "
            End If
        
        Next
        
        If equipos(loopEquipo).descalificado = 0 Then
            obtenerStringPuntos = obtenerStringPuntos & " " & equipos(loopEquipo).cantidadRoundsGanados
        Else
            obtenerStringPuntos = obtenerStringPuntos & " (des)"
        End If
        
        'Cuando proceso el ultimo equipo no le pongo el ultimo "-" ya que no quda nada más.
        If loopEquipo < cantidadEquipos Then
            obtenerStringPuntos = obtenerStringPuntos & " - "
        Else
            obtenerStringPuntos = obtenerStringPuntos & "."
        End If
        
    Next

End Function

'Obtiene los integrantes del equipo separados por  "," e "y"
Private Function obtenerStringEquipo(idEquipo As Byte) As String
    Dim loopIntegrante As Byte
    Dim cantidadIntegrantes As Byte

    cantidadIntegrantes = equipos(idEquipo).cantidadIntegrantes
    
    
    For loopIntegrante = 1 To cantidadIntegrantes
        
        obtenerStringEquipo = obtenerStringEquipo & equipos(idEquipo).integrantes(loopIntegrante).nick
        
        If loopIntegrante = cantidadIntegrantes - 1 Then
            obtenerStringEquipo = obtenerStringEquipo & " y "
        ElseIf loopIntegrante < cantidadIntegrantes Then
            obtenerStringEquipo = obtenerStringEquipo & " , "
        End If

    Next

End Function

'Obtiene el nombre de los equipos separados por " Vs "
Private Function obtenerStringPrensetacion() As String

    Dim loopEquipo As Byte
    
    obtenerStringPrensetacion = ""
    
    For loopEquipo = 1 To cantidadEquipos
    
        obtenerStringPrensetacion = obtenerStringPrensetacion & obtenerStringEquipo(loopEquipo)
        
        If loopEquipo < cantidadEquipos Then
            obtenerStringPrensetacion = obtenerStringPrensetacion & " Vs "
        Else
            obtenerStringPrensetacion = obtenerStringPrensetacion & "."
        End If
        
    Next loopEquipo

End Function

' TODO. Esta funcion solo existe porque el "Acepto" de la estructura se usa para dos cosas
' Esto deberia eliminarse
Private Sub enviarMensajeParticipantesQueAceptaron(ByRef mensaje As String)
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    Dim UserIndex As Integer
    
    For loopEquipo = 1 To cantidadEquipos
        'Si fue descalificado no esta más con nosotros
        For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
        
            If equipos(loopEquipo).integrantes(loopIntegrante).Acepto = 1 Then
                UserIndex = equipos(loopEquipo).integrantes(loopIntegrante).UserIndex
                    
                If UserIndex = 0 Then
                    UserIndex = IDIndex(equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante)
                End If
                    
                If UserIndex > 0 Then
                    EnviarPaquete Paquetes.MensajeGuild, nombreReto & "> " & mensaje, UserIndex, ToIndex
                End If
                
            End If
        Next loopIntegrante
    Next loopEquipo
    
    'EnviarPaquete Paquetes.MensajeGuild, nombreReto & "> " & mensaje, 0, ToAdmins
End Sub

Private Sub enviarMensajeParticipantes(ByRef mensaje As String)
    Dim loopEquipo As Byte
    Dim loopIntegrante As Byte
    Dim UserIndex As Integer
    For loopEquipo = 1 To cantidadEquipos
        'Si fue descalificado no esta más con nosotros
        If equipos(loopEquipo).descalificado = 0 Then
            For loopIntegrante = 1 To equipos(loopEquipo).cantidadIntegrantes
            
                UserIndex = equipos(loopEquipo).integrantes(loopIntegrante).UserIndex
                
                If UserIndex = 0 Then
                    UserIndex = IDIndex(equipos(loopEquipo).integrantes(loopIntegrante).IDIntegrante)
                End If
                
                If UserIndex > 0 Then
                    EnviarPaquete Paquetes.MensajeGuild, nombreReto & "> " & mensaje, UserIndex, ToIndex
                End If
            Next loopIntegrante
        End If
    Next loopEquipo
    
    'EnviarPaquete Paquetes.MensajeGuild, nombreReto & "> " & mensaje, 0, ToAdmins
End Sub
'Termina el evento
Private Sub terminar()

'Si inicie un combate lo libero
If Not combate Is Nothing Then
    Set combate = Nothing
End If

'Llibero el ring
If ring.id > 0 Then
    Call modRings.liberarRing(ring)
End If

'Libero el descanso
If descanso.id > 0 Then
    Call modDescansos.liberarZonaDescanso(descanso)
End If

'Este evento se estaba jugando y termino
If Estado = eEstadoEvento.Desarrollandose Then
    Call modRetos.terminoReto
    Call LogNuevosRetos("Termino el reto " & obtenerStringPrensetacion())
End If

'Finalmente el evento termino.
Estado = eEstadoEvento.Terminado
End Sub
'Funciones de la interfaz
Public Sub iEvento_usuarioAbandono(UserIndex As Integer)
    'Es lo mismo que cerrar
    Call iEvento_usuarioCerro(UserIndex)
End Sub
Public Sub iEvento_usuarioCerro(UserIndex As Integer)

    Dim IDEquipoGanador_ As Byte
    Dim tipoResultado As eResultadoCombate
    Dim mensaje As String
    Dim equipoDescalificado As Byte
    
    If idEquipoGanador = 0 Then
        
        IDEquipoGanador_ = combate.descalificarEquipoDelUsuario(UserIndex, tipoResultado)
       
        equipoDescalificado = obtenerEquipo(UserList(UserIndex).id)
        
        'Es una persona o un equipo el que quedo afuera?
        If equipos(equipoDescalificado).cantidadIntegrantes > 1 Then
            mensaje = "Equipo de " & obtenerStringEquipo(obtenerEquipo(UserList(UserIndex).id)) & " descalificado porque " & UserList(UserIndex).Name & " cerro el juego."
        Else
            mensaje = UserList(UserIndex).Name & " descalificado porque cerro el juego."
        End If
        
        Call enviarMensajeParticipantes(mensaje)
        'Call LogNuevosRetos(mensaje)
        
        If IDEquipoGanador_ > 0 Then
        
            'Le sumo del round
            equipos(IDEquipoGanador_).cantidadRoundsGanados = equipos(IDEquipoGanador_).cantidadRoundsGanados + 1
            'Este round le sirvio para ganar?
            If tipoResultado = eResultadoCombate.final Then
                'Terminando...
                Call procesarGanador(IDEquipoGanador_)
            Else
                'Ya descalifique al equipo del combate. Como esta descalificado todo el equipo
                'ya es un equipo perdedor
                Call procesarEquipoPerdedor(obtenerEquipo(UserList(UserIndex).id))
                'Aviso...
                Call enviarMensajeParticipantes("Gana el round el equipo de " & obtenerStringEquipo(IDEquipoGanador_) & ".")
                Call enviarMensajeParticipantes(obtenerStringPuntos())
                'Vamos, que sigue!
                Call combate.continuar
            End If
        Else
            'Ya descalifique al equipo del combate. Como esta descalificado todo el equipo
            'ya es un equipo perdedor
            Call procesarEquipoPerdedor(obtenerEquipo(UserList(UserIndex).id))

        End If
    Else
        'El reto ya estaba terminado. El usuario estaba en el descanso recogiendo los items
        'Lo saco del repositorio
        Dim loopIntegrante As Byte
        
        For loopIntegrante = 1 To equipos(idEquipoGanador).cantidadIntegrantes
            If equipos(idEquipoGanador).integrantes(loopIntegrante).IDIntegrante = UserList(UserIndex).id Then
                Call deshacermeUsuario(idEquipoGanador, loopIntegrante)
                Exit Sub
            End If
        Next loopIntegrante
    End If

End Sub

Public Function iEvento_getEstadoEvento() As eEstadoEvento
    iEvento_getEstadoEvento = Estado
End Function

Public Function iEvento_getTipoEvento() As eTipoEvento
    iEvento_getTipoEvento = eTipoEvento.reto
End Function

Public Sub iEvento_timeOutMinuto()
    'Depende el estado es lo que pasa.

    Debug.Print "Descanso " & descanso.id
    tiempoTranscurrido = tiempoTranscurrido + 1
    
    If Estado = eEstadoEvento.Preparacion Then
    
        If tiempoTranscurrido = 4 Then
        
            'Se cancela la creacion del reto
            'TODO. Se le deberia enviar un mensaje a todos los que aceptaron.
            'Pudo no iniciarse porque no se aceptarn todas las solicitudes
            'O ubo alguna falla cuando se los quizo sumonear a todos
            '(faltaba el oro, no estaba online, estaba en otro evento ya).
            'Si no todas las solicitudes fueron aceptadas entonces
            'envio el mensaje de que no todos aceptaron el combate
            'Envio mensaje de que se cancelo
            'Call enviarMensajeParticipantes("El reto fue cancelado debido a que no todos los invitados aceptaron la solicitud para ingresar.")
                
            Call iEvento_cancelar
        End If
        
    ElseIf Estado = eEstadoEvento.Desarrollandose Then
       
        If idEquipoGanador = 0 Then
            
            If tiempoTranscurrido + 5 = tiempoMaximo Then
                'Envio mensajae de que en 5 minutos termina
                Call enviarMensajeParticipantes("Últimos 5 minutos de combate. Si no hay ganador el reto será cancelado y perderan sus apuestas.")
            ElseIf tiempoTranscurrido >= tiempoMaximo Then
                'Cancelo el reto. Nadie gana.
                'Log
                Call LogNuevosRetos("Cancelado por time out. Hasta el momento " & obtenerStringPuntos())
                'Envio mensaje de que se cancelo
                Call enviarMensajeParticipantes("El combate fue cancelado por exceder el tiempo máximo de duración. Los participantes pierden sus apuestas.")
                'Finalmente lo cancelo
                Call iEvento_cancelar
            End If
        
        Else
        
            'El reto ya termino, hay un ganador, pero esta en el repositorio
            'sacando los items
            If tiempoTranscurrido = 2 Then
                'Los saco del repositorio
                Dim loopIntegrante As Byte
                
                For loopIntegrante = 1 To equipos(idEquipoGanador).cantidadIntegrantes
                    If equipos(idEquipoGanador).integrantes(loopIntegrante).Acepto = 1 Then
                        Call deshacermeUsuario(idEquipoGanador, loopIntegrante)
                    End If
                Next
                
                'Listo, terminamos
                Call terminar
            End If
        End If
        
    End If
   
End Sub

Public Sub iEvento_timeOutSegundo()
End Sub

Public Function iEvento_puedeLanzar(indexHechizo As Byte) As Boolean
    
    'Si no esta desarrollandose no lo jodo
    If Estado = eEstadoEvento.Desarrollandose Then
        iEvento_puedeLanzar = reglaHechizos(indexHechizo)
    Else
        iEvento_puedeLanzar = True
    End If

End Function

Public Function iEvento_getIDGanador() As Byte
    iEvento_getIDGanador = idEquipoGanador
End Function

Public Function iEvento_getTimeTranscurrido() As Integer
    iEvento_getTimeTranscurrido = tiempoTranscurrido
End Function

Public Function iEvento_getNombre() As String
    iEvento_getNombre = nombreReto
End Function

Public Function iEvento_isInscripcionesAbiertas() As Boolean
    iEvento_isInscripcionesAbiertas = False
End Function

Public Function iEvento_getCantidadParticipantesEquipo() As Byte
    iEvento_getCantidadParticipantesEquipo = 0
End Function
Public Function iEvento_agregarEquipo(IntergantesIndexs() As Integer) As eResultadoInscripcion
    'Dummy
End Function

Public Sub iEvento_usuarioIngreso(UserIndex As Integer, UserID As Long)
End Sub

Public Function iEvento_getDescripcion() As String
    iEvento_getDescripcion = ""
End Function

Private Sub iEvento_usuarioRevive(UserIndexRevivido As Integer, UserIndexRevividor As Integer)
    If Not combate Is Nothing Then
        'If Not esperandoIngreso Then 'TODO Mejorar
            Call combate.reviviruUsuario(UserIndexRevivido)
        ' End If
    End If
End Sub
