VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iEvento_DeathMach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**
' Es la hoja de ruta del evento. Va armando los distintos modos de torneos
' con los equipos que correspondan.
'
'@see
Implements iEvento

Option Explicit

Private Enum eSubEstadoDesarrollo
    indefinido  'No hay ningun sub estado del desarrollo del evento
    inscripcionesabiertas  'Los usuarios pueden enviar su inscripcion
    inscripcionesCerradas  'Los usuarios ya no pueden enviar su inscripcion. Ya estan todos listos, prepanadose para pelear
    EnCombate 'Los personajes se encuentran peleando
    combateTerminado 'Se termino el combate
    esperandoConfirmacionDeFin 'Espera a que se confirme el exito del evento (para pagar las apuestas)
End Enum

'*************************************************************
Private Estado As eEstadoEvento
Private subEstadoDesarrollo As eSubEstadoDesarrollo
'*************************************************************
Private nombreEvento As String                              ' Nombre del evento
Private descripcionEvento As String                         ' Descripción del evento
Private cantidadEstrellas As Byte                           ' Importancia del evento
Private comoIdentificarEquipo As eIdentificacionEquipos     ' ¿Como se debe identificar a los equipos? Nombre del equipo.
'*************************************************************
Private MAX_TIEMPO_AVISO As Long            '  Cantidad de tiempo durante el cual se va a estar anunciando el evento
Private MAX_TIEMPO_INSCRIPCION As Long      '  Cantidad de minutos maximo que puede durar la iscripcion al evento
Private MAX_TIEMPO_PREPARACION As Long      '  Cantidad de minutos que tienen los usuarios para preparse antes del comienzo de los combates
Private MAX_TIEMPO_DURACION As Long         '  Cantidad de minutos que puede durar el evento
Private MAX_TIEMPO_TOLERANCIA As Long       '  Cantidad de minutos que puede estar un usuario Offline cuando se lo llama antes de ser expulsado
Private cantidadInscripcionesTerminadas As Integer  ' Cantidad de veces que se cerraron inscripciones (toma sentido en los eventos circulares, luego siempre va a ser 1)
'*************************************************************
Private cantidadEquiposMaxima As Byte       '  Cantidad maxima de tablaequipos que soporta el evento
Private cantidadEquiposMinima As Byte       '  Cantidad minima de tablaequipos que debe haber antes para poder comenzar y que no sea cancelado por falta de publico.
Private cantidadEquiposPorCombate As Byte   '  ¿Cuantos equipos hay dentro del mismo combate?

Private cantidadParticipantesEquipo As Byte ' Cantidad de personas que integran cada equipo
Private cantidadEquiposInscriptos As Byte   ' Cantidad de tablaequipos que finalmente se inscribieron

Private oroInscripcion As Long              '  Cuanto oro tienen que pagar los usuarios para poder participar

Private tiempoTranscurrido As Integer       'Minutos que transcurrieron
'************************************************************
' Recursos
Private tipoRing As eRingTipo
Private tipoDescanso As eDescansoTipo

Private descansos() As modDescansos.tZonaDescanso   ' Lista de descansos donde voy a poner a los usuarios mientras esperan que comience el combate
Private rings() As modRings.tRing                   ' Rings que utiliza el evento
Private tengoDescansos As Boolean                   ' Booleano necesario para saber si trate de tomar algun descanso
Private tengoRings As Boolean                       ' Booleano necesario para saber si trate de tomar algun descanso

Private subEvento As iModoTorneo                    ' Modos de torneo
Private eventosDeEquipos() As iModoTorneo           ' Relacion eventos equipos
'************************************************************
' Apuestas
Private MAX_TIEMPO_APUESTAS_ABIERTAS As Byte    ' Tiempo que tienen los users para apostar
Private apuestaID As Integer                    ' Cada subEvento puede tener su apuesta
Private apuestasActivadas As Boolean
Private pozoBaseApuestas As Long
'************************************************************
' Premio
Private tablaPagos() As Long                    ' Lo  que corresponde de oro depende del puesto
Private tablaPagosExpresion As eEventoPremio    ' Formato en el cual esta expresada la tabla de pagos
Private oroRecaudado As Long                    ' Cantidad de monedas de oro que se recaudo a través de la inscripcion
'************************************************************
Private reglaHechizos() As Boolean              ' ¿Se puede utilizar determinado hechizo?

Private tablaEquipos() As tEquipoTablaTorneo    ' Equipos participando. El ID del equipo coincide con la posicon en esta tabla

Private condicionCantidadRoundsGanados As Byte  ' Variables de condiciones.. ver si da para pasarlo a un objeto

'Condiciones para ingresar al evento
Private listaCondiciones As Collection

' La hoja de ruta establece cual es el Modo con el cual se organizaran los enfrentamientos.
' Los parametros son informacion adicional. Por ejemplo si una liga es de ida o vuelta.
Private hojaDeRuta As String
Private parametros As String

'TODO Re Veer
Private soloInscripcion As Boolean

'TODO re ver
Private IDPersistencia As Long

'Cuento cada vez que ingresa alguien y guardo esta informacion en cada equipo para saber quien mando primero
'Me parece que es mejor que un gettickcount. No se puede confiar en el ID del equipo para ordenarlos
'por el tema de que hay equipos que pueden abandonar durante la inscripcio
Private ultimoTickDeInscripcion As Integer

Private Sub Class_Terminate()
    Call LogTorneos(nombreEvento & "-> Se elimina de memoria.")
End Sub

Public Sub iEvento_agregarCondicionIngreso(condicion As iCondicionEvento)
    Call listaCondiciones.Add(condicion)
End Sub

Public Sub setTiempos(ByVal tiempoAviso As Integer, ByVal tiempoMaxInscripcion As Byte, ByVal tiempoTolerancia As Byte)
    MAX_TIEMPO_INSCRIPCION = tiempoMaxInscripcion
    MAX_TIEMPO_TOLERANCIA = tiempoTolerancia
    MAX_TIEMPO_AVISO = tiempoAviso
End Sub
Private Function desincribirEquipo(equipo As tEquipoTablaTorneo)
    Dim condicion As iCondicionEvento
    Dim loopIntegrante As Byte
    Dim integrantesid() As Long
    
    equipo.idEquipo = 0
    cantidadEquiposInscriptos = cantidadEquiposInscriptos - 1
    
    ReDim integrantesid(1 To equipo.cantidadIntegrantes) As Long
    
    For loopIntegrante = 1 To equipo.cantidadIntegrantes
        integrantesid(loopIntegrante) = equipo.integrantes(loopIntegrante).id
    Next
    
    For Each condicion In listaCondiciones
        Call condicion.removerEquipoInscripto(integrantesid)
    Next
    
End Function

Public Sub iEvento_configurarApuestas(ByVal activadas As Boolean, ByVal pozoInicial As Long, ByVal tiempo As Byte)
    apuestasActivadas = activadas
    pozoBaseApuestas = pozoInicial
    MAX_TIEMPO_APUESTAS_ABIERTAS = tiempo
End Sub

Public Sub iEvento_establecerTablaDePagos(tablaPagos_() As Long, tipoExpresion As eEventoPremio)
    tablaPagos = tablaPagos_
    tablaPagosExpresion = tipoExpresion
End Sub

Private Function iEvento_getCantidadEquiposMax() As Byte
        iEvento_getCantidadEquiposMax = cantidadEquiposMaxima
End Function

Public Function iEvento_isCumpleCondicionEquipo(userIndexs() As Integer) As Boolean
    Dim loopIntegrante As Byte
    Dim condicion As iCondicionEvento
    
    'Condicion fija, de tener el oro.
    For loopIntegrante = 1 To UBound(userIndexs)
        If Not isCumpleCondicion(UserList(userIndexs(loopIntegrante))) Then
            iEvento_isCumpleCondicionEquipo = False
            Exit Function
        End If
    Next

    'Condiciones dinamicas
    For Each condicion In listaCondiciones
        If Not condicion.puedeIngresarEquipo(userIndexs) Then
            iEvento_isCumpleCondicionEquipo = False
            Exit Function
        End If
    Next
    
    'Les digo a las condiciones que tengan en cuenta que entro
    'este equipo
    For Each condicion In listaCondiciones
        Call condicion.asentarIngresoEquipo(userIndexs)
    Next
        
    iEvento_isCumpleCondicionEquipo = True
End Function

'Devuelve true si los tablaequipos se pueden inscribir a este evento
Public Function iEvento_isInscripcionesAbiertas() As Boolean
    iEvento_isInscripcionesAbiertas = (Estado = eEstadoEvento.Desarrollandose And subEstadoDesarrollo = eSubEstadoDesarrollo.inscripcionesabiertas)
End Function

' ¿El usuario puede tirar un objeto?
Public Function iEvento_puedeTirarObjeto(UserIndex As Integer, objeto As Integer, ByVal cantidad As Long, destino As eDestinoObjeto, destinoIndex As Integer) As Boolean
    
 ' No se pueden tirar objetos al suelo
 ' Se pueden tirar objetos a criaturas
 ' Se puede tirar objetos a personajes siempre que sean del mismo equipo
 If destino = eDestinoObjeto.suelo Then
    EnviarPaquete Paquetes.mensajeinfo, "No puedes tirar objetos al suelo en este momento, revisá las reglas del evento.", UserIndex, ToIndex
    iEvento_puedeTirarObjeto = False
 ElseIf destino = eDestinoObjeto.criatura Then
    iEvento_puedeTirarObjeto = True
 ElseIf destino = eDestinoObjeto.Usuario Then
    ' ¿El destino index es el mismo equipo?
    If modTorneos.sonCompaneros(tablaEquipos, UserList(UserIndex).id, UserList(destinoIndex).id) Then
        iEvento_puedeTirarObjeto = True
    Else
        iEvento_puedeTirarObjeto = False
        EnviarPaquete Paquetes.mensajeinfo, "No puedes darle objetos a un personaje que no sea de tu equipo.", UserIndex, ToIndex
    End If
 End If
 
End Function


Private Sub Class_Initialize()
    Estado = eEstadoEvento.esperandoConfirmacionInicio
    
#If testeo = 1 Then
    MAX_TIEMPO_AVISO = 1
#Else
    MAX_TIEMPO_AVISO = 10
#End If
    
    MAX_TIEMPO_INSCRIPCION = 10
    MAX_TIEMPO_PREPARACION = 1
    MAX_TIEMPO_DURACION = 30
    
    MAX_TIEMPO_APUESTAS_ABIERTAS = 10
    
    cantidadEstrellas = 3
    
    cantidadEquiposMinima = 2
    cantidadEquiposMaxima = 2
    cantidadEquiposPorCombate = 2
    
    cantidadParticipantesEquipo = 2
    
    cantidadInscripcionesTerminadas = 0
        
    comoIdentificarEquipo = eIdentificacionEquipos.identificaPersonajes
        
    ' Por defecto
    tipoRing = eRingTipo.ringTorneo
    tipoDescanso = eDescansoTipo.torneo
    
    ultimoTickDeInscripcion = 0
    
    Call setSoloInscripcion(False)
    
    Call modTorneos.cargarReglasbasicasHechizos(reglaHechizos)
    

    
    Set listaCondiciones = New Collection
    Call LogTorneos("Se crea un nuevo torneo")
End Sub

Public Sub iEvento_publicar()
    ' Cambiamos el estado
    Estado = eEstadoEvento.Preparacion
    tiempoTranscurrido = 0
    
    'Me fijo si ya se ocuparon todos los cupos (porque se inscribieron manualmente)
    ' en la etapa de confirmacion
    ' y estamos listos para empezar
    If cantidadEquiposInscriptos = cantidadEquiposMaxima Then
        ' Cerramos las inscripciones automaticamente y comenzamos el evento
        ' Solo si el evento esta confirmado. Sino lo voy a hacer cuando se confirme la publicacion
        Call preComienzo
        Exit Sub
    End If
    
    ' Sino arranca directamente, paso a la etapa de inscripcion
    If MAX_TIEMPO_AVISO > 0 Then
        Call modTorneos.enviarMensajeGlobal(descripcionEvento & " La inscripción comenzará en " & HelperTiempo.segundosAHoras(MAX_TIEMPO_AVISO * 60) & ".", nombreEvento)
    Else
        Call abrirInscripcion
    End If

End Sub
' Toma los recursos necesario para desarrollar el evento
Public Function iniciar() As Boolean
    iniciar = False
    
    'Obtengo el espacio necesario para almacenar a los equipos
    ReDim tablaEquipos(1 To cantidadEquiposMaxima) As tEquipoTablaTorneo
    
    'Estructura guardo la referencia de donde este cada equipo
    ReDim eventosDeEquipos(1 To cantidadEquiposMaxima) As iModoTorneo

    ' Obtengo los recursos
    If soloInscripcion = False Then
        If hojaDeRuta = "NADA" Then 'No hay sub evento, solo neceisot descansos
            If Not obtenerRecursos(cantidadEquiposMaxima, cantidadParticipantesEquipo, 0) Then
                Call LogTorneos("ERROR. " & nombreEvento & "-> No se pudieron obtener los recursos.")
                Call terminar
                Exit Function
            End If
        Else 'Neceisto descansos y rings
            If Not obtenerRecursos(cantidadEquiposMaxima, cantidadParticipantesEquipo, cantidadEquiposMaxima \ cantidadEquiposPorCombate) Then
                Call LogTorneos("ERROR. " & nombreEvento & "-> No se pudieron obtener los recursos.")
                Call terminar
                Exit Function
            End If
        End If
    Else
        ReDim descansos(0)
        ReDim rings(0)
    End If
    
    Call LogTorneos("Se crea el evento con los siguientes datos:" & vbCrLf & iEvento_obtenerInfoExtendida())
    Call LogTorneos(nombreEvento & "-> Se inicia correctamente")
    iniciar = True
End Function

Private Function obtenerRecursos(cantidadEquiposMaxima As Byte, cantidadParticipantesEquipo As Byte, cantidadRings As Byte) As Boolean
    'Obtengo los descansos necesarios
    Dim loopC As Byte
        
    'Obtengo los descanso que necesito
    ReDim descansos(1 To cantidadEquiposMaxima) As modDescansos.tZonaDescanso
    tengoDescansos = True
    
    For loopC = 1 To cantidadEquiposMaxima
    
        ' Obtengo el descanso respectivo
        descansos(loopC) = modDescansos.getZonaDescanso(cantidadParticipantesEquipo, tipoDescanso)
        
        'Pude obtener el descanso?
        If descansos(loopC).id = 0 Then
            Call LogTorneos("ERROR. " & nombreEvento & "-> No se pudo obtener el descanso " & loopC)
            obtenerRecursos = False
            Exit Function
        Else
            Call LogTorneos(nombreEvento & "-> SE obtuvo el descanso " & descansos(loopC).id)
        End If
        
    Next loopC
    
    
    'Obtengo los rings
    If cantidadRings > 0 Then
    
        ReDim rings(1 To cantidadRings) As tRing
        tengoRings = True
        
        For loopC = 1 To cantidadRings

            rings(loopC) = modRings.obtenerRing(cantidadEquiposPorCombate, cantidadParticipantesEquipo, tipoRing)
    
            If rings(loopC).id = 0 Then
                Call LogTorneos("ERROR. " & nombreEvento & "-> No se pudo obtener el ring " & loopC & " para el torneo.")
                obtenerRecursos = False
                Exit Function
            End If
        
            Call LogTorneos(nombreEvento & "-> Se obtuvo el ring " & rings(loopC).id)
        Next loopC
    Else
        ReDim rings(0)
    End If
    
    obtenerRecursos = True

End Function

Private Function liberarRecursos()

    Dim loopC As Byte
    
    ' ¿Tengo algun descanso?
    If tengoDescansos = True Then
        tengoDescansos = False
        For loopC = 1 To UBound(descansos)
              'Si hay un descanso, lo libero
            If descansos(loopC).id > 0 Then
                Call LogTorneos(nombreEvento & "-> Se libera el descanso " & descansos(loopC).id)
                Call modDescansos.liberarZonaDescanso(descansos(loopC))
            End If
        Next loopC
    End If

    '¿Tengo algun ring?
    If tengoRings = True Then
        tengoRings = False
        For loopC = 1 To UBound(rings)
            If rings(loopC).id > 0 Then
                Call LogTorneos(nombreEvento & "-> Se libera el ring " & rings(loopC).id)
                Call modRings.liberarRing(rings(loopC))
            End If
        Next loopC
    End If
    
End Function

Private Function log(desc As String)
    Call LogTorneos(nombreEvento & "-> " & desc)
End Function
Public Function iEvento_agregarEquipo(integrantesIndexs() As Integer) As eResultadoInscripcion
    Dim loopIntegrante As Byte
    Dim presentacion As String
    Dim UserIndex As Integer
    Dim numeroEquipo As Byte
    
    Call LogTorneos(nombreEvento & "-> Se intenta inscribir al equipo liderador por " & UserList(integrantesIndexs(1)).Name)
    
    
    If Not cantidadEquiposInscriptos < cantidadEquiposMaxima Then
        iEvento_agregarEquipo = eResultadoInscripcion.inscripcionCerrada
        Call LogTorneos("ERROR. " & nombreEvento & " se intenta inscribir cuando no hay más lugar.")
        Exit Function
    End If
    
    'Chequeo lo más importante que es que tenga el oro.
    For loopIntegrante = 1 To cantidadParticipantesEquipo
        
        UserIndex = integrantesIndexs(loopIntegrante)
        
         If Not isCumpleCondicion(UserList(UserIndex)) Then
            Call LogTorneos(nombreEvento & "-> El participante " & UserList(UserIndex).Name & " no tiene el oro.")
            iEvento_agregarEquipo = noTieneOro
            Exit Function
         End If
         
    Next loopIntegrante

    'Busco un lugar.
    'Esto se puso por si alguno abandona en mitad de la inscripcion.
    For numeroEquipo = 1 To cantidadEquiposMaxima
        If tablaEquipos(numeroEquipo).idEquipo = 0 Then
            Exit For
        End If
    Next numeroEquipo
    
    If numeroEquipo > cantidadEquiposMaxima Then
        iEvento_agregarEquipo = eResultadoInscripcion.inscripcionCerrada
        Call LogTorneos("ERROR. " & nombreEvento & " no encontro lugar cuando deberia haber.")
        Exit Function
    End If
    
    'ESTA OK
    
    'Contabilizo la nueva inscripcion
    ultimoTickDeInscripcion = ultimoTickDeInscripcion + 1
    
    ' Aumentamos el Oro recaudado
    oroRecaudado = oroRecaudado + oroInscripcion * UBound(integrantesIndexs)
    
    'Redimensiono donde guardo los integrates
    ReDim tablaEquipos(numeroEquipo).integrantes(1 To UBound(integrantesIndexs)) As tIntegrantesEquipoTorneo

    Dim nPos As WorldPos

    With tablaEquipos(numeroEquipo)

        'Asignaciones generales
        .Estado = eEstadoEquipoTorneo.participando

        .idEquipo = numeroEquipo
        
        .NombreEquipo = obtenerNombreIdentificacionPersonajes(integrantesIndexs, comoIdentificarEquipo)
           
        .cantidadCombatesEmpatados = 0
        .cantidadCombatesGanados = 0
        .cantidadCombatesJugados = 0

        .cantidadRoundsGanados = 0
        .cantidadRoundsJugados = 0

        .cantidadIntegrantes = cantidadParticipantesEquipo
        .cantidadIntegrantesDescalificados = 0
        
        'Para saber en que orden se inscribio
        .tickInscripcion = ultimoTickDeInscripcion
        
        For loopIntegrante = 1 To cantidadParticipantesEquipo
            
            UserIndex = integrantesIndexs(loopIntegrante)
            'Obtengo los dats del usuario
            .integrantes(loopIntegrante).UserIndex = UserIndex
            .integrantes(loopIntegrante).nick = UserList(UserIndex).Name
            .integrantes(loopIntegrante).posOriginal = UserList(UserIndex).pos
            .integrantes(loopIntegrante).id = UserList(UserIndex).id
            .integrantes(loopIntegrante).cantidadAdvertencias = 0

            .integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Jugando
   
            .integrantes(loopIntegrante).cantidadOroPagadoInscripcion = 0
            
            'Le saco el oro
            #If TDSFacil = 0 Then
                If Not UserList(UserIndex).Premium Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - oroInscripcion
                    .integrantes(loopIntegrante).cantidadOroPagadoInscripcion = oroInscripcion
                End If
            #Else
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - oroInscripcion
                .integrantes(loopIntegrante).cantidadOroPagadoInscripcion = oroInscripcion
            #End If
            
            'Lo actualizo
            EnviarPaquete Paquetes.EnviarOro, Codify(UserList(UserIndex).Stats.GLD), UserIndex, ToIndex

            'Lo mando al descanso
            If Not soloInscripcion Then
                Call ClosestLegalPos(descansos(numeroEquipo).centro, nPos, UserList(UserIndex))
                Call WarpUserChar(UserIndex, nPos.map, nPos.x, nPos.y, False)
    
                'Lo doy el mensaje de bienvenida
                If UserList(UserIndex).Genero = eGeneros.Mujer Then
                    EnviarPaquete Paquetes.MensajeTalk, nombreEvento & "-> Bienvenida al evento " & UserList(UserIndex).Name & ".", UserIndex, ToIndex
                Else
                    EnviarPaquete Paquetes.MensajeTalk, nombreEvento & "-> Bienvenido al evento " & UserList(UserIndex).Name & ".", UserIndex, ToIndex
                End If

                'No abandones eh
                 EnviarPaquete Paquetes.MensajeTalk, nombreEvento & "-> Puedes abandonar el evento con el comando /ABANDONAR. Esto se considera una falta grave de respeto, perderás la inscripción y puedes ser penado.", UserIndex
            
                'Le relaciono el evento
                Set UserList(UserIndex).evento = Me
            Else
                'Lo doy el mensaje de bienvenida
                If UserList(UserIndex).Genero = eGeneros.Mujer Then
                    EnviarPaquete Paquetes.MensajeTalk, nombreEvento & "-> Felicitaciones. Has sido inscripta en el evento. Conformas el equipo número " & numeroEquipo & ". Puedes seguir jugando, pero estate listo cuando este por comenzar el evento.", UserIndex, ToIndex
                Else
                    EnviarPaquete Paquetes.MensajeTalk, nombreEvento & "-> Felicitaciones. Has sido inscripto en el evento. Conformas el equipo número " & numeroEquipo & ". Puedes seguir jugando, pero estate listo cuando este por comenzar el evento.", UserIndex, ToIndex
                End If
            End If
        Next loopIntegrante

    End With
    
    Call LogTorneos(nombreEvento & "-> Se inscribe al equipo " & modTorneos.obtenerStringEquipo(tablaEquipos(numeroEquipo), False, eFormatoDisplayEquipo.completo))
       
    'Si llegue acá quiere decir que se isncribio
    cantidadEquiposInscriptos = cantidadEquiposInscriptos + 1
     
    'Me fijo si ya se ocuparon todos los cupos y estamos listos para empezar
    If cantidadEquiposInscriptos = cantidadEquiposMaxima Then
        ' Cerramos las inscripciones automaticamente y comenzamos el evento
        ' Solo si el evento esta confirmado. Sino lo voy a hacer cuando se confirme la publicacion
        If Not Estado = eEstadoEvento.esperandoConfirmacionInicio Then Call preComienzo
    End If 'Seguimos esperando más participantes
    
    iEvento_agregarEquipo = eResultadoInscripcion.correcta
        
End Function

'Devuelve true si el usuario tiene las condiciones necesarias para entrar al evento
'TODO Ver donde meter la condicion del oro y condiciones que le quitan cosas a las personas
Private Function isCumpleCondicion(Usuario As User) As Boolean
   
    isCumpleCondicion = False
        
    'Tiene el oro?
    #If TDSFacil = 0 Then
    If Usuario.Stats.GLD < oroInscripcion And Not Usuario.Premium Then
    #Else
    If Usuario.Stats.GLD < oroInscripcion Then
    #End If
        Exit Function
    End If
     
    isCumpleCondicion = True

End Function

'Los usuarios se deben preparar para combatir
Private Sub preComienzo()
    
    Dim loopDescanso As Byte
    Dim loopEquipo As Byte
    Dim mensajeGlobal As String
    
    Call LogTorneos(nombreEvento & "-> Se cierran las inscripciones. Cuenta regresiva para el comienzo.")
    
    Estado = eEstadoEvento.Desarrollandose
    subEstadoDesarrollo = eSubEstadoDesarrollo.inscripcionesCerradas
    tiempoTranscurrido = 0
    
    'Me fijo si no alcace el maximo de equipos inscriptos.
    If cantidadEquiposInscriptos < cantidadEquiposMaxima Then
    
        Call LogTorneos(nombreEvento & "-> Se comprime lista.")
            
        Dim pivot As Byte 'El equipo que ira al slot vacio
        
        loopEquipo = 1
        pivot = UBound(tablaEquipos)
        
        'Obtengo el equipo que esta en el ultimo slot
        Do While loopEquipo <= cantidadEquiposInscriptos
            If tablaEquipos(loopEquipo).idEquipo = 0 Then
                Do While Not tablaEquipos(pivot).idEquipo > 0
                    pivot = pivot - 1
                Loop
                
                'Libero el descanso que no se uso. No quiero teletrapsortar a la gente al descanso que quedo libre
                If Not soloInscripcion Then
                    Call modDescansos.liberarZonaDescanso(descansos(loopEquipo))
                
                    descansos(loopEquipo) = descansos(pivot)
                End If
                
                'Pongo al equipo en el slot libre
                tablaEquipos(loopEquipo) = tablaEquipos(pivot)
                tablaEquipos(loopEquipo).idEquipo = loopEquipo
                
                pivot = pivot - 1
            End If
            
            'Paso al siguiente equipo
            loopEquipo = loopEquipo + 1

        Loop
        
        'Redimensiono el array de equipos
        ReDim Preserve tablaEquipos(1 To cantidadEquiposInscriptos)
    End If
    
    If cantidadParticipantesEquipo = 1 Then
        mensajeGlobal = "Inscripciones finalizadas. Personajes inscriptos (" & cantidadEquiposInscriptos & "): " & modTorneos.obtenerStringPrensetacion(tablaEquipos, " - ", False, eFormatoDisplayEquipo.completo)
    Else
        mensajeGlobal = "Inscripciones finalizadas. Equipos inscriptos (" & cantidadEquiposInscriptos & "): " & modTorneos.obtenerStringPrensetacion(tablaEquipos, " - ", False, eFormatoDisplayEquipo.completo)
    End If
    
    If tablaPagosExpresion = eEventoPremio.porcentajeSobreAcumulado Then
        mensajeGlobal = mensajeGlobal & " Recaudación: " & HelperDinero.formatearDinero(oroRecaudado) & "."
    End If
        
    Call modTorneos.enviarMensajeGlobal(mensajeGlobal, nombreEvento)
    
    'Cancelo la solicitud de todos los usuarios que no llegaron a ingresar
    Call modEventos.cancelarSolicitudesPendientesParaElEvento(Me)
    
    'Redimensiono dependiendo la cantidad de sub eventos en esta fase
    apuestaID = 0
    
    Set subEvento = Nothing
    
    If Not hojaDeRuta = "NADA" Then
    
        Set subEvento = modFacadeModos.crearModo(hojaDeRuta)
        
        Call subEvento.setNombreEvento(nombreEvento)
        Call subEvento.setNecesidadRoundGanadosGanador(condicionCantidadRoundsGanados)
        Call subEvento.setParametros(parametros)
        Call subEvento.setMaximaCantidadAdvertencias(MAX_TIEMPO_TOLERANCIA)
        
        Call modCasteoTorneos.setRings(subEvento, rings)
    
        Call modCasteoTorneos.setTablaEquipos(subEvento, tablaEquipos)
        
        If IDPersistencia > 0 Then
            Call subEvento.cargar(IDPersistencia)
        End If
        
        'Guardo la referencia de donde esta cada equipo
        For loopEquipo = 1 To cantidadEquiposInscriptos
            Set eventosDeEquipos(loopEquipo) = subEvento
        Next
    End If
    
    '¿Estan las apuestas activadas?
    If apuestasActivadas Then
        apuestaID = modApuestasInterface.crearApuesta(Me, pozoBaseApuestas, cantidadEstrellas, tablaEquipos)
   
        'Le aviso a la gente que puede apostar
        Call modTorneos.enviarMensajeGlobal("Las apuestas estan abiertas. Pozo inicial " & pozoBaseApuestas & " monedas de oro. Tienen " & MAX_TIEMPO_APUESTAS_ABIERTAS & " minutos para apostar desde su [URL=Cuenta Premium -> Seccion Apuestas]http://tierrasdelsur.cc/cpremium.html?a=apuestas .", nombreEvento)
           
        MAX_TIEMPO_PREPARACION = MAX_TIEMPO_APUESTAS_ABIERTAS
   End If
   
   If Not hojaDeRuta = "NADA" Then
        If MAX_TIEMPO_PREPARACION = 1 Then
            Call modTorneos.enviarMensajeEquipos(tablaEquipos, "Tienen un minuto para prepararse.", nombreEvento)
        Else
            Call modTorneos.enviarMensajeEquipos(tablaEquipos, "Tienen " & MAX_TIEMPO_PREPARACION & " minutos para prepararse.", nombreEvento)
        End If
    End If
End Sub

'Comienzan los combates. Se le da el poder al ModoTorneo para que haga lo que corresponda con los equipos
Private Sub comenzar()
    
    Call LogTorneos(nombreEvento & "-> Comienza el torneo.")
    
    'Empezamos!!!
    Call subEvento.comenzar(0)
    subEstadoDesarrollo = eSubEstadoDesarrollo.EnCombate
    
    If subEvento.isTerminado Then Call procesarGanador(subEvento)
End Sub

Private Sub preTerminar()
    If apuestasActivadas And apuestaID > 0 Then
        'Si en este evento estan las apuestas activadas, las cierro
        subEstadoDesarrollo = eSubEstadoDesarrollo.esperandoConfirmacionDeFin
    Else
        'Termimos directo
        Call terminar
    End If
End Sub
Private Sub terminar()
    
    Call liberarRecursos
    'El evento termino
    Estado = eEstadoEvento.Terminado

    Call LogTorneos(nombreEvento & "-> Termina.")
End Sub

Private Sub procesarGanador(ByVal subEvento As iModoTorneo)
    Dim loopC As Byte
    Dim loopIntegrante As Byte
    Dim premioOro As Long
    Dim mensajeGlobal As String
    Dim mensajePremio As String
    Dim primerPremioOro As Long
    
    tablaEquipos = modCasteoTorneos.obtenerTabla(subEvento)
    
    ' ¿Hay una condicion especial de finalización?
    ' Si la hay, me fijo si se cumplio. Si se cumplió, le aviso a la gente que se cumplió.
    ' Si no se cumplio, pregunto a cuantos debo mantener.
    ' Mantengo a esos y al resto los llevo al descanso.
    ' Vuelvo a abrir las inscripciones y el primero en entrar fueron los N primeros.
    
    '**************************************************************
    'Le entrego el premio
    For loopC = 1 To UBound(tablaPagos)
    
        If tablaPagosExpresion = monedasDeOro Then
            premioOro = Int(tablaPagos(loopC) / tablaEquipos(loopC).cantidadIntegrantes)
        Else ' Porcentaje
            premioOro = Int(((tablaPagos(loopC) / 100) * oroRecaudado) / tablaEquipos(loopC).cantidadIntegrantes)
        End If
        
        If loopC = 1 Then
            primerPremioOro = premioOro * tablaEquipos(loopC).cantidadIntegrantes
        End If
        
        If premioOro > 0 Then
            mensajePremio = nombreEvento & "-> ¡Felicitaciones por el " & loopC & "º lugar!. Se te entregó el premio de " & FormatNumber(premioOro, 0, vbTrue, vbFalse, vbTrue) & " monedas de oro."
        Else
            mensajePremio = nombreEvento & "-> ¡Felicitaciones por el " & loopC & "º lugar!."
        End If
        
        For loopIntegrante = 1 To tablaEquipos(loopC).cantidadIntegrantes
            Call modTorneos.entregarOro(tablaEquipos(loopC).integrantes(loopIntegrante).id, premioOro, mensajePremio)
        Next loopIntegrante
        
    Next loopC
    
    Call LogTorneos(nombreEvento & "-> Se entregaron los premios")
    '*************************************************************
    
    'Retorno los usuarios que me quedaron a sus pociones.
    For loopC = 1 To cantidadEquiposInscriptos
        If tablaEquipos(loopC).Estado = participando Then
            tablaEquipos(loopC).Estado = eEstadoEquipoTorneo.termino
            Call modTorneos.desecharEquipo(tablaEquipos(loopC), 0)
        End If
    Next loopC
    
   '¿Quien gano?
    If cantidadParticipantesEquipo = 1 Then
        mensajeGlobal = "¡¡Ganó el evento " & modTorneos.obtenerStringEquipo(tablaEquipos(1), False, eFormatoDisplayEquipo.completo) & "!!."
    Else
        mensajeGlobal = "¡¡Ganador del evento el equipo de " & modTorneos.obtenerStringEquipo(tablaEquipos(1), False, eFormatoDisplayEquipo.completo) & "!!."
    End If
    
    '¿Gano monedas de oro?
    If primerPremioOro > 0 Then
        mensajeGlobal = mensajeGlobal & " Premio " & HelperDinero.formatearDinero(primerPremioOro) & " monedas de oro."
    End If
    
    'Aviso a todos los usuarios que hubo un ganador
    Call modTorneos.enviarMensajeGlobal(mensajeGlobal, nombreEvento)
        
    'El torneo llego a lo ultimo...
    subEstadoDesarrollo = eSubEstadoDesarrollo.combateTerminado

End Sub
' *************************************************************************************
' SETERS
Public Sub setCantidadRoundGanadosGanador(cantidad As Byte)
    condicionCantidadRoundsGanados = cantidad
End Sub

Public Sub setComoIdentificarEquipos(tipo As eIdentificacionEquipos)
    comoIdentificarEquipo = tipo
End Sub

Public Sub setNombre(nombre_ As String)
    nombreEvento = nombre_
End Sub

Public Sub setCantidadParticipantesEquipo(cantidad As Byte)
    cantidadParticipantesEquipo = cantidad
End Sub
Public Sub setCantidadMaxMinEquipos(Minima As Byte, Maxima As Byte)
    cantidadEquiposMaxima = Maxima
    cantidadEquiposMinima = Minima
End Sub
Public Sub setPrecioInscripcion(Precio As Long)
    oroInscripcion = Precio
End Sub
Public Sub setDescripcion(descripcion_ As String)
    descripcionEvento = descripcion_
End Sub

Private Function iEvento_obtenerIDPersistencia() As Integer
    Dim subEvento As iModoTorneo
    Dim loopSubEvento As Byte
    
    iEvento_obtenerIDPersistencia = 0
    
    If Estado = eEstadoEvento.Desarrollandose Then
        If subEstadoDesarrollo = EnCombate Then
            iEvento_obtenerIDPersistencia = subEvento.obtenerIDPersistencia
        End If
    End If
End Function

Private Function iEvento_obtenerInfoEstado() As String
    If Estado = eEstadoEvento.Preparacion Then
        iEvento_obtenerInfoEstado = "La inscripción comenzará en " & HelperTiempo.segundosAHoras((MAX_TIEMPO_AVISO - tiempoTranscurrido) * 60) & "."
    ElseIf Estado = eEstadoEvento.Desarrollandose Then
        If iEvento_isInscripcionesAbiertas Then
            iEvento_obtenerInfoEstado = "Las inscripciones están abiertas. " & modTorneos.generarMensajeParticipar(nombreEvento, cantidadParticipantesEquipo)
        Else
            iEvento_obtenerInfoEstado = "El evento se está desarrollando. Las inscripciones están cerradas."
        End If
    ElseIf Estado = eEstadoEvento.Terminado Then
        iEvento_obtenerInfoEstado = "El evento acaba de terminar."
    End If
End Function

Private Function iEvento_obtenerInfoExtendida() As String
    Dim tempbyte As Byte
    Dim contador As Byte
    Dim auxEquipo As tEquipoTablaTorneo
    
    iEvento_obtenerInfoExtendida = nombreEvento & vbCrLf & vbCrLf & descripcionEvento & vbCrLf & iEvento_obtenerInfoEstado
        
    If soloInscripcion Then
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "SIN EVENTO AUTOMATICO. Solo se generará la lista de inscriptos y ni siquiear se los sumone."
    End If
    
    iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Tipo: " & hojaDeRuta
    iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Precio inscripción: " & oroInscripcion
    iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "De " & cantidadEquiposMinima & " a " & cantidadEquiposMaxima & " equipos de " & cantidadParticipantesEquipo & " participantes"
        
    If soloInscripcion Then
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "No se utilizan Descansos ni Rings."
    Else
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Descansos del tipo: " & modDescansos.getDescTipoDescanso(tipoDescanso)

        If Not hojaDeRuta = "NADA" Then  'No hay sub evento, solo neceisot descansos
            iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Rings del tipo: " & modRings.getDescTipoRing(tipoRing)
        Else
            iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "No se necesitan rings."
        End If
    End If
    
    If Not hojaDeRuta = "NADA" Then
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Combates al mejor de " & condicionCantidadRoundsGanados
    End If
    
    If apuestasActivadas Then
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Apuestas activadas. Pozo inicial: " & pozoBaseApuestas
    End If
    
    ' Premios
    If tablaPagosExpresion = eEventoPremio.monedasDeOro Then
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Premios (monedas de oro): "
    Else
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Premios (porcentaje sobre la recaudacion: " & HelperDinero.formatearDinero(oroRecaudado) & "): "
    End If
    
    For tempbyte = LBound(tablaPagos) To UBound(tablaPagos)
            iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & tempbyte & "º) " & tablaPagos(tempbyte)
    Next tempbyte

    iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Cantidad inscriptos: " & cantidadEquiposInscriptos
    
    If cantidadEquiposInscriptos > 0 Then
            
        If comoIdentificarEquipo = eIdentificacionEquipos.identificaPersonajes Then
            iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Equipos (identificados con los nombres de los personajes):"
        ElseIf comoIdentificarEquipo = eIdentificacionEquipos.identificaClan Then
            iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Equipos (identificados con el nombre del clan):"
        ElseIf comoIdentificarEquipo = eIdentificacionEquipos.identificaFaccion Then
            iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "Equipos (identificados por la Alineacion):"
        End If

        For tempbyte = 1 To UBound(tablaEquipos)
            If tablaEquipos(tempbyte).idEquipo > 0 Then
            
                contador = contador + 1
                
                If Estado = eEstadoEvento.Desarrollandose And subEstadoDesarrollo = eSubEstadoDesarrollo.EnCombate And Not eventosDeEquipos(tempbyte) Is Nothing Then
                    auxEquipo = modCasteoTorneos.obtenerEquipo(eventosDeEquipos(tempbyte), tablaEquipos(tempbyte).idEquipo)
                        
                    iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & contador & ") " & _
                    modTorneos.obtenerStringEquipoConEstado(auxEquipo, eFormatoDisplayEquipo.completo)
                Else
                    iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & contador & ") " & _
                    modTorneos.obtenerStringEquipo(tablaEquipos(tempbyte), False, eFormatoDisplayEquipo.completo)
                End If
            End If
        Next tempbyte
    End If
    
    If listaCondiciones.Count > 0 Then
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & vbCrLf & "Restricciones:"
        
        Dim condicion As iCondicionEvento
        
        For Each condicion In listaCondiciones
            iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & "- " & condicion.obtenerDescripcion
        Next
    Else
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & vbCrLf & "No hay restricciones para ingresar a este evento."
    End If
    
    Debug.Print iEvento_obtenerInfoExtendida
    iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & vbCrLf & vbCrLf & "Hechizos permitidos:"
    
    Dim hechizosPermitidos As String
    hechizosPermitidos = ""
    
    For tempbyte = 1 To UBound(reglaHechizos)
        If reglaHechizos(tempbyte) Then
            hechizosPermitidos = hechizosPermitidos & "- " & hechizos(tempbyte).nombre
        End If
    Next tempbyte
    
    If hechizosPermitidos = "" Then
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & " Ninguno"
    Else
        iEvento_obtenerInfoExtendida = iEvento_obtenerInfoExtendida & hechizosPermitidos
    End If
    
    
    Debug.Print iEvento_obtenerInfoExtendida
    
End Function

Public Sub iEvento_setHechizosPermitidos(reglas() As Boolean)
    reglaHechizos = reglas
End Sub

Public Sub iEvento_setHojaDeRuta(hojaRuta As String, parametros_ As String)
    hojaDeRuta = hojaRuta
    parametros = parametros_
    
    If hojaDeRuta = "DEATHMATCH" Then
        cantidadEquiposPorCombate = cantidadEquiposMaxima
    End If
        
End Sub

Public Sub iEvento_setTiporing(tipoRing_ As eRingTipo)
    tipoRing = tipoRing_
End Sub

Public Sub iEvento_setTipoDescanso(tipoDescanso_ As eDescansoTipo)
    tipoDescanso = tipoDescanso_
End Sub


Private Sub abrirInscripcion()
    Call LogTorneos(nombreEvento & "-> Se abren las inscripciones.")
    Estado = eEstadoEvento.Desarrollandose
    subEstadoDesarrollo = eSubEstadoDesarrollo.inscripcionesabiertas
    tiempoTranscurrido = 0
    'Envio el mensaje de que las inscripciones estan abiertas
    Call modTorneos.enviarMensajeGlobal(descripcionEvento, nombreEvento)
    Call modTorneos.enviarMensajeGlobal("INSCRIPCIONES ABIERTAS." & modTorneos.generarMensajeParticipar(nombreEvento, cantidadParticipantesEquipo), nombreEvento)
End Sub

' ************************************************************************************
' IMPLEMENTACIONES DE LA INTERFACE
Public Sub iEvento_timeOutMinuto()
    'Depende el estado es lo que pasa.

    tiempoTranscurrido = tiempoTranscurrido + 1
    
    If Estado = eEstadoEvento.Preparacion Then
    
            'Si falta 1 minuto, aviso.
            'Si faltan menos de 10 minutos aviso cada 3 minutos
            'Si faltan menos de 60 minutos aviso cada 10 minutos.
            'Si faltan 60 o más minutos aviso cada 50 minutos
            If tiempoTranscurrido = MAX_TIEMPO_AVISO Then
                'Termino el tiempo de aviso de torneo, que empiecen las inscripciones
                Call abrirInscripcion
            ElseIf (MAX_TIEMPO_AVISO - tiempoTranscurrido) = 1 Then
                Call modTorneos.enviarMensajeGlobal(descripcionEvento & " La inscripción comenzará en un minuto.", nombreEvento)
            ElseIf (MAX_TIEMPO_AVISO - tiempoTranscurrido) < 10 And (tiempoTranscurrido Mod 3) = 0 Then
                Call modTorneos.enviarMensajeGlobal(descripcionEvento & " La inscripción comenzará en " & MAX_TIEMPO_AVISO - tiempoTranscurrido & " minutos.", nombreEvento)
            ElseIf (MAX_TIEMPO_AVISO - tiempoTranscurrido) < 60 And (tiempoTranscurrido Mod 10) = 0 Then
                Call modTorneos.enviarMensajeGlobal(descripcionEvento & " La inscripción comenzará en " & MAX_TIEMPO_AVISO - tiempoTranscurrido & " minutos.", nombreEvento)
            ElseIf tiempoTranscurrido Mod 50 = 0 Then
                 Call modTorneos.enviarMensajeGlobal(descripcionEvento & " La inscripción comenzará en " & HelperTiempo.segundosAHoras((MAX_TIEMPO_AVISO - tiempoTranscurrido) * 60) & ".", nombreEvento)
            End If
    
    ElseIf Estado = eEstadoEvento.Desarrollandose Then
    
            If subEstadoDesarrollo = inscripcionesabiertas Then
                'Paso el tiempo maximo de espera para las inscripciones?
                If tiempoTranscurrido = MAX_TIEMPO_INSCRIPCION Then
                    'El tiempo maximo paso. Entonces..
                    'Si no alcanzamos la cantidad minima de inscriptos entonces
                    'se cancela el evento
                    'caso contrario se inicia con los que hay
                    cantidadInscripcionesTerminadas = cantidadInscripcionesTerminadas + 1
                    
                    If cantidadEquiposInscriptos >= cantidadEquiposMinima Then
                        Call preComienzo
                    Else 'No hay la suficiente cantidad de inscriptos
                        Call modTorneos.enviarMensajeGlobal("El evento se cancela por falta de participantes.", nombreEvento)
                        Call iEvento_cancelar
                    End If
                ElseIf ((MAX_TIEMPO_INSCRIPCION - tiempoTranscurrido) Mod 3) = 0 Then
                    'Cada 3 minutos aviso que las inscripciones estan abiertas
                    Call modTorneos.enviarMensajeGlobal(descripcionEvento, nombreEvento)
                    
                    If cantidadInscripcionesTerminadas > 0 And tablaPagosExpresion = eEventoPremio.porcentajeSobreAcumulado Then
                        Call modTorneos.enviarMensajeGlobal("Recaudación de monedas de oro hasta el momento: " & HelperDinero.formatearDinero(oroRecaudado) & ".", nombreEvento)
                    End If
                    
                    Call modTorneos.enviarMensajeGlobal("INSCRIPCIONES ABIERTAS." & modTorneos.generarMensajeParticipar(nombreEvento, cantidadParticipantesEquipo), nombreEvento)
                End If
            ElseIf subEstadoDesarrollo = inscripcionesCerradas Then
                'Termino el tiempo para que se preparen?
                    If tiempoTranscurrido >= MAX_TIEMPO_PREPARACION Then
                        
                        'Aviso que las apuestas estan cerradas
                        If apuestasActivadas Then
                            Call modApuestasInterface.cerrarApuesta(apuestaID)
                            Call modTorneos.enviarMensajeGlobal("Apuestas cerradas.", nombreEvento)
                        End If
                    
                        'Empezamos!!! Siempre y cuando tenga un evento automatico, sino practicamente muere
                        If hojaDeRuta <> "NADA" Then
                            Call comenzar
                        Else
                            Call preTerminar
                        End If
                        
                    ElseIf tiempoTranscurrido = MAX_TIEMPO_PREPARACION - 1 Then
                    
                        'Les aviso que en un minutito empieza
                        If apuestasActivadas Then
                            Call modTorneos.enviarMensajeGlobal("En un minuto comenzará el evento y se cerrarán las apuestas.", nombreEvento)
                        Else
                            Call modTorneos.enviarMensajeEquipos(tablaEquipos, "En un minuto comenzará el evento.", nombreEvento)
                        End If
                        
                    ElseIf tiempoTranscurrido Mod 3 = 0 And apuestasActivadas Then
                    
                        'Les aviso que apuesten
                        Call modTorneos.enviarMensajeGlobal("Las apuestas están abiertas. Tienen " & (MAX_TIEMPO_PREPARACION - tiempoTranscurrido) & " minutos para apostar desde la [URL=Cuenta Premium -> Seccion Apuestas.]http://tierrasdelsur.cc/cpremium.html?a=apuestas Allí se encuentra toda la información del evento.", nombreEvento)
                    
                    End If
            ElseIf subEstadoDesarrollo = EnCombate Then
            
                Call pasarMinutoEventos
                
            ElseIf subEstadoDesarrollo = eSubEstadoDesarrollo.combateTerminado Then
                'Veo la tabla de resultados final
                Call modTorneos.enviarTabla(tablaEquipos, nombreEvento, ToAll)
            
                'Terminamos
                Call preTerminar
            End If
    End If

   
End Sub

Private Sub pasarMinutoEventos()
    
    Call subEvento.pasarMinuto
    
    If subEvento.isTerminado Then Call procesarGanador(subEvento)
        
End Sub
'Funciones de la interfaz
Public Sub iEvento_usuarioMuere(UserIndex As Integer)
    Dim idEquipoJugador As Byte
    
    Call LogTorneos(nombreEvento & "-> Muere " & UserList(UserIndex).Name)

    idEquipoJugador = modTorneos.obtenerIDEquipoPersonaje(tablaEquipos, UserList(UserIndex).id)

    If Not eventosDeEquipos(idEquipoJugador) Is Nothing Then
        Call eventosDeEquipos(idEquipoJugador).usuarioMuere(UserIndex)
        If eventosDeEquipos(idEquipoJugador).isTerminado Then Call procesarGanador(eventosDeEquipos(idEquipoJugador))
    End If ' No deberia pasar esto
    
End Sub

Public Sub iEvento_usuarioRevive(UserIndexRevivido As Integer, UserIndexResucitor As Integer)
    Dim idEquipoJugador As Byte
    
    If UserIndexResucitor = 0 Then
        Call LogTorneos(nombreEvento & "-> Revivido " & UserList(UserIndexRevivido).Name)
    Else
        Call LogTorneos(nombreEvento & "-> Revivido " & UserList(UserIndexRevivido).Name & " gracias a " & UserList(UserIndexResucitor).Name)
    End If
    
    'Obtengo el equipo del usuario que revivi
    idEquipoJugador = modTorneos.obtenerIDEquipoPersonaje(tablaEquipos, UserList(UserIndexRevivido).id)

    If Not eventosDeEquipos(idEquipoJugador) Is Nothing Then
        Call eventosDeEquipos(idEquipoJugador).usuarioRevive(UserIndexRevivido, UserIndexResucitor)
    Else
        If hojaDeRuta <> "NADA" Then 'TODO Mejorar
            Call modUtilitarios.Preparando(UserIndexRevivido)
        End If
    End If
    
End Sub

Public Sub iEvento_usuarioAbandono(UserIndex As Integer)

    Call LogTorneos(nombreEvento & "-> Abandono " & UserList(UserIndex).Name)
    
    Dim idEquipoJugador As Byte
    
    'Los usuarios son dueños de otro evento?
    'De cual?
    idEquipoJugador = modTorneos.obtenerIDEquipoPersonaje(tablaEquipos, UserList(UserIndex).id)
        
    'El jugador no esta en ningun sub evento, es prioridad mia
    If eventosDeEquipos(idEquipoJugador) Is Nothing Then
               
        'Si n usuario abandona el evento cuando estamos en la inscripciones
        'desechamos a todo el equipo ya que no vamos a arrancar con un equipo muleto..
        If Estado = eEstadoEvento.Desarrollandose And subEstadoDesarrollo = eSubEstadoDesarrollo.inscripcionesabiertas Then
         
            'Le aviso a su equipo que el pibe abandono
            Call modTorneos.enviarMensajeEquipo(tablaEquipos(idEquipoJugador), UserList(UserIndex).Name & " abandono el evento.", nombreEvento)
 
            'Le aviso al usuario que es un pancho por abandonar
            EnviarPaquete Paquetes.MensajeTalk, "Abandonaste el evento. Esto se considera una falta de respeto grave para los usuarios participantes y los no participantes. Perdiste la inscripción y podrás ser penado.", UserIndex
                    
            If tablaEquipos(idEquipoJugador).cantidadIntegrantes = 1 Then
                Call modTorneos.enviarMensajeGlobal("El participante " & UserList(UserIndex).Name & " abandono el evento. Se liberá su cupo para otro participante.", nombreEvento)
            Else
                Call modTorneos.enviarMensajeGlobal("El equipo de " & modTorneos.obtenerStringEquipo(tablaEquipos(idEquipoJugador), False, eFormatoDisplayEquipo.completo) & " abandono el evento. Se liberá su cupo para otro equipo.", nombreEvento)
            End If
            
            Call modTorneos.desecharEquipo(tablaEquipos(idEquipoJugador), 0)
            Call desincribirEquipo(tablaEquipos(idEquipoJugador))
                
        Else 'Esta en mitad del torneo
            'Obtengo el integrante
            'Lo marco como descalificado
            'Me fijo si tengo que descalificar a todo el equipo
            'Si descalifico al equipo me fijo si queda más de un equipo
            '   Si queda un solo equipo, declaro al ganador
            Dim loopIntegrante As Byte
        
            For loopIntegrante = 1 To tablaEquipos(idEquipoJugador).cantidadIntegrantes
                If tablaEquipos(idEquipoJugador).integrantes(loopIntegrante).id = UserList(UserIndex).id Then
                            
                    tablaEquipos(idEquipoJugador).cantidadIntegrantesDescalificados = tablaEquipos(idEquipoJugador).cantidadIntegrantesDescalificados + 1
                            
                    tablaEquipos(idEquipoJugador).integrantes(loopIntegrante).Estado = eEstadoIntegranteEquipo.Descalificando
                            
                    Call modTorneos.desecharIntegranteEquipo(tablaEquipos(idEquipoJugador).integrantes(loopIntegrante), 0)
                        
                    'Le aviso al mundo
                    Call modTorneos.enviarMensajeGlobal(UserList(UserIndex).Name & " abandono el evento.", nombreEvento)
                            
                    'Le aviso al usuario que es un pancho por abandonar
                    EnviarPaquete Paquetes.MensajeTalk, "Abandonaste el evento. Esto se considera una falta de respeto grave para los usuarios participantes y los no participantes. Perdiste la inscripción y podrás ser penado.", UserIndex
                End If
            Next
        
            If Not tieneCantidadMinimaDeParticipantes(tablaEquipos(idEquipoJugador)) Then
                If tablaEquipos(idEquipoJugador).cantidadIntegrantes = 1 Then
                    Call modTorneos.enviarMensajeEquipos(tablaEquipos, "Equipo de " & modTorneos.obtenerStringEquipo(tablaEquipos(idEquipoJugador), False, eFormatoDisplayEquipo.NombreEquipo) & " descalificado por no contar con la cantidad minima de integrantes.", nombreEvento)
                Else
                    Call modTorneos.desecharEquipo(tablaEquipos(idEquipoJugador), 0)
                End If
                
                tablaEquipos(idEquipoJugador).Estado = eEstadoEquipoTorneo.descalificado
            End If
        End If

    Else 'No es de mi propiedad el jugador.
        'Le indico al evento que el jugador cerro
        Call eventosDeEquipos(idEquipoJugador).usuarioAbandono(UserIndex)
        
        '¿Este evento causo que termine el sub evento?
        If eventosDeEquipos(idEquipoJugador).isTerminado Then Call procesarGanador(eventosDeEquipos(idEquipoJugador))
    End If
    
End Sub

Public Sub iEvento_usuarioCerro(UserIndex As Integer)
    'Si el evento esta en pleno desarrollo y el tipo esta jugando
    'directamente se lo descalifica. No puede cerrar en medio de un combate
    'Sino se lo agrega a la lista de usuarios offline. No se lo descalifica de una
    'Si luego no vuelve a loguear en el plazo especificado, ahí lo eliminamos.
    Call LogTorneos(nombreEvento & "-> Cerro " & UserList(UserIndex).Name)
    
    Dim idEquipoJugador As Byte
    
    'Los usuarios son dueños de otro evento?
    'De cual?
    idEquipoJugador = modTorneos.obtenerIDEquipoPersonaje(tablaEquipos, UserList(UserIndex).id)
        
    'El jugador no esta en ningun sub evento, es prioridad mia
    If eventosDeEquipos(idEquipoJugador) Is Nothing Then
        
        'Lo marco como offline
        Call modTorneos.actualizarUserIndexPersonaje(tablaEquipos, UserList(UserIndex).id, 0)
        
        Call modEventos.agregarUsuarioOffline(UserList(UserIndex).id, Me)
        
    Else 'No es de mi propiedad el jugador.
        'Le indico al evento que el jugador cerro
        Call eventosDeEquipos(idEquipoJugador).usuarioCerro(UserIndex)
        
        '¿Este evento causo que termine el sub evento?
        If eventosDeEquipos(idEquipoJugador).isTerminado Then Call procesarGanador(eventosDeEquipos(idEquipoJugador))
    End If

End Sub

'Alguien que particiba de este evento cerro y ahora volvio a loguear
Public Sub iEvento_usuarioIngreso(UserIndex As Integer, UserID As Long)
    
    Dim idEquipoJugador As Byte
        
    Call LogTorneos(nombreEvento & "-> Ingreso " & UserList(UserIndex).Name)
    
    idEquipoJugador = modTorneos.obtenerIDEquipoPersonaje(tablaEquipos, UserList(UserIndex).id)
         
    'Lo doy el mensaje de bienvenida
    If UserList(UserIndex).Genero = eGeneros.Mujer Then
        EnviarPaquete Paquetes.MensajeTalk, nombreEvento & "> Bienvenida de nuevo al evento " & UserList(UserIndex).Name & ". Por favor, estate atento y no cierres el juego cuando este por comenzar un combate. Puedes quedar eliminado.", UserIndex, ToIndex
    Else
        EnviarPaquete Paquetes.MensajeTalk, nombreEvento & "> Bienvenido de nuevo al evento " & UserList(UserIndex).Name & ". Por favor, estate atento y no cierres el juego cuando este por comenzar un combate. Puedes quedar eliminado.", UserIndex, ToIndex
    End If
    
    'El jugador no esta en ningun sub evento, es prioridad mia
    If eventosDeEquipos(idEquipoJugador) Is Nothing Then
        Call modTorneos.actualizarUserIndexPersonaje(tablaEquipos, UserID, UserIndex)
    Else
        Call eventosDeEquipos(idEquipoJugador).usuarioIngreso(UserIndex, UserID)
    End If

End Sub

Public Sub iEvento_cancelar()

    Dim loopEquipo  As Byte
    Dim loopIntegrante As Byte
    Dim tabla() As tEquipoTablaTorneo
    
    Call LogTorneos(nombreEvento & "-> es cancelado.")
    
    If Estado = eEstadoEvento.Preparacion Then
    
        'Estaen preparaicon todavia no se trajo a nadie
        
    ElseIf Estado = eEstadoEvento.Desarrollandose Then
    
        'Cancelo los sub eventos
        If subEstadoDesarrollo = eSubEstadoDesarrollo.EnCombate Then
        
            Call subEvento.cancelar
            
            tabla = modCasteoTorneos.obtenerTabla(subEvento)
            
            'Hay algunos usuarios
            For loopEquipo = 1 To UBound(tabla)
                If tabla(loopEquipo).Estado = eEstadoEquipoTorneo.participando Then
                    tabla(loopEquipo).Estado = termino
                    Call modTorneos.desecharEquipo(tablaEquipos(loopEquipo), oroInscripcion, True)
                End If
            Next loopEquipo
            
        Else 'Insc Abierras/ Cerradas / Combate terminado
        
            For loopEquipo = 1 To UBound(tablaEquipos)
                If tablaEquipos(loopEquipo).idEquipo > 0 Then 'Hay un equipo
                    If tablaEquipos(loopEquipo).Estado = eEstadoEquipoTorneo.participando Then
                        tablaEquipos(loopEquipo).Estado = eEstadoEquipoTorneo.termino
                        Call modTorneos.desecharEquipo(tablaEquipos(loopEquipo), oroInscripcion, True)
                    End If
                End If
            Next loopEquipo
    
        End If

        If apuestasActivadas And apuestaID > 0 Then
            Call modApuestasInterface.cancelarApuesta(apuestaID)
        End If
     
    End If 'Terminado. No queda nadie en el evento.


    Call terminar
End Sub

Public Sub iEvento_timeOutSegundo()
End Sub

Public Function iEvento_getEstadoEvento() As eEstadoEvento
    iEvento_getEstadoEvento = Estado
End Function

Public Function iEvento_getTipoEvento() As eTipoEvento
    iEvento_getTipoEvento = eTipoEvento.torneo
End Function

Public Function iEvento_puedeLanzar(indexHechizo As Byte) As Boolean
    'Si no esta desarrollandose no lo jodo
    If Estado = eEstadoEvento.Desarrollandose Then
        iEvento_puedeLanzar = reglaHechizos(indexHechizo)
    Else
        iEvento_puedeLanzar = True
    End If

   Exit Function
End Function

Public Function iEvento_getIDGanador() As Byte
    iEvento_getIDGanador = 0
End Function

Public Function iEvento_getTimeTranscurrido() As Integer
    iEvento_getTimeTranscurrido = tiempoTranscurrido
End Function

Public Function iEvento_getNombre() As String
    iEvento_getNombre = nombreEvento
End Function

Public Function iEvento_getCantidadParticipantesEquipo() As Byte
    iEvento_getCantidadParticipantesEquipo = cantidadParticipantesEquipo
End Function

Public Function iEvento_getDescripcion() As String
    iEvento_getDescripcion = descripcionEvento
End Function


Private Function tieneCantidadMinimaDeParticipantes(equipo As tEquipoTablaTorneo) As Boolean
    tieneCantidadMinimaDeParticipantes = (equipo.cantidadIntegrantesDescalificados < equipo.cantidadIntegrantes)
End Function

Public Sub setSoloInscripcion(si As Boolean)
    soloInscripcion = si
End Sub

Private Sub guardarEquiposEnArchivo()
    Dim nfile As Integer
    Dim numeroEquipo As Byte
    
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\" & nombreEvento & "-Insc.txt" For Append Shared As #nfile
    
    Print #nfile, "-------------" & nombreEvento & " " & Date & " " & Time & " " & "---------------"
  
    For numeroEquipo = 1 To UBound(tablaEquipos)
        If tablaEquipos(numeroEquipo).idEquipo > 0 Then
            Print #nfile, tablaEquipos(numeroEquipo).idEquipo & ")" & modTorneos.obtenerStringEquipo(tablaEquipos(numeroEquipo), False, eFormatoDisplayEquipo.completo)
        End If
    Next numeroEquipo
    
    Close #nfile
End Sub

Public Sub iEvento_guardar()
    Dim loopSubEvento As Byte
    
    If Estado = eEstadoEvento.Desarrollandose Then
        If subEstadoDesarrollo = EnCombate Then
            subEvento.guardar
        End If
    End If
End Sub

Public Sub iEvento_setIDPersistencia(IDPersistencia_ As Long)
    IDPersistencia = IDPersistencia_
End Sub

'ATENCION: Solo sirve para pagar las apuestas o confirmar el ganador en el caso
'de que el evento sea 100% automatico y no solo una inscripcion
Public Function iEvento_establecerGanadorManualmente(nombreEquipoGanador As String) As Boolean
    Dim ok As Boolean
    ok = False
    'Evento asutomatico?
    If soloInscripcion Or (soloInscripcion = False And hojaDeRuta = "NADA") Then
        Dim loopEquipo As Byte
        
        For loopEquipo = 1 To UBound(tablaEquipos)
            If modTorneos.obtenerStringEquipo(tablaEquipos(loopEquipo), False, eFormatoDisplayEquipo.completo) = nombreEquipoGanador Then
                Call modApuestasInterface.establecerGanador(apuestaID, tablaEquipos(loopEquipo).idEquipo)
                ok = True
                Exit For
            End If
        Next loopEquipo

    Else
        ok = True
        Call modApuestasInterface.establecerGanador(apuestaID, tablaEquipos(1).idEquipo)
    End If
    
    'Ahora si terminamos
    If ok Then
        terminar
    End If
    
    iEvento_establecerGanadorManualmente = ok
End Function
